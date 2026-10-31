VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvoice 
   Caption         =   "Invoice Batch Processing - Clinic"
   ClientHeight    =   8256
   ClientLeft      =   0
   ClientTop       =   240
   ClientWidth     =   13308
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8256
   ScaleWidth      =   13308
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtLGNo 
      Height          =   285
      Left            =   9240
      TabIndex        =   275
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdLGSearch 
      Caption         =   "&Go"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10560
      TabIndex        =   274
      Top             =   2400
      Width           =   585
   End
   Begin MSComctlLib.ListView lstInvoice 
      Height          =   345
      Left            =   0
      TabIndex        =   146
      Top             =   7800
      Visible         =   0   'False
      Width           =   11895
      _ExtentX        =   20976
      _ExtentY        =   614
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
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   480
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Batch No"
         Object.Width           =   1411
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
         Text            =   "Inv. No"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bill No."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Inv. Date"
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
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Clinic"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Bill Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "PreparedBy"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "101"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "102"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "103"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Text            =   "104"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Text            =   "105"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Text            =   "106"
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
         Text            =   "POth Desc"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Text            =   "LOth Desc"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Text            =   "IOth Desc"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Text            =   "MedInjType1"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   48
         Text            =   "MedInjType2"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Text            =   "MedInjType3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Text            =   "MedInjType4"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Text            =   "MedInjType5"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   52
         Text            =   "MedInjDay1"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   53
         Text            =   "MedInjDay2"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   54
         Text            =   "MedInjDay3"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   55
         Text            =   "MedInjDay4"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   56
         Text            =   "MedInjDay5"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   57
         Text            =   "MedInjQty1"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   58
         Text            =   "MedInjQty2"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   59
         Text            =   "MedInjQty3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   60
         Text            =   "MedInjQty4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   61
         Text            =   "MedInjQty5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   62
         Text            =   "Diagnosis1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(64) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   63
         Text            =   "Diagnosis2"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(65) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   64
         Text            =   "Diagnosis3"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(66) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   65
         Text            =   "Diagnosis4"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(67) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   66
         Text            =   "Payor"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(68) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   67
         Text            =   "CLN Plan"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(69) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   68
         Text            =   "Cln Ins"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(70) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   69
         Text            =   "HLT Code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(71) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   70
         Text            =   "Bal Limit"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(72) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   71
         Text            =   "Visit Lmt"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(73) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   72
         Text            =   "Card Type"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(74) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   73
         Text            =   "Visit Ind"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(75) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   74
         Text            =   "Annual Ind"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(76) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   75
         Text            =   "Co-pay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(77) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   76
         Text            =   "Pol No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(78) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   77
         Text            =   "Patient IC"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(79) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   78
         Text            =   "Eff Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(80) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   79
         Text            =   "Exp Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(81) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   80
         Text            =   "Payto"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(82) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   81
         Text            =   "Payee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(83) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   82
         Text            =   "Supp Lmt Type"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(84) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   83
         Text            =   "GP Fee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(85) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   84
         Text            =   "GP+LT Med"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(86) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   85
         Text            =   "GP + Inj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(87) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   86
         Text            =   "GP + Med"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(88) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   87
         Text            =   "SP Fee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(89) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   88
         Text            =   "SP+ LT Med"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(90) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   89
         Text            =   "SP + Inj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(91) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   90
         Text            =   "SP + Med"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(92) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   91
         Text            =   "total Gp(c1)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(93) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   92
         Text            =   "total SP(C2)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(94) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   93
         Text            =   "OtherDiagCode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(95) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   94
         Text            =   "OtherDiagind"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(96) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   95
         Text            =   "OthDiagDesc"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(97) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   96
         Text            =   "DiagDesc1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(98) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   97
         Text            =   "DiagDesc2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(99) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   98
         Text            =   "DiagDesc3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(100) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   99
         Text            =   "DiagDesc4"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(101) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   100
         Text            =   "DiagDesc5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(102) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   101
         Text            =   "Vision"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(103) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   102
         Text            =   "MedType1(5)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(104) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   103
         Text            =   "MedType1(6)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(105) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   104
         Text            =   "MedDay1(5)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(106) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   105
         Text            =   "MedDay1(6)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(107) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   106
         Text            =   "MedQty1(5)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(108) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   107
         Text            =   "MedQty1(6)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(109) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   108
         Text            =   "Reserve4InvoiceMBMPostageStatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(110) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   109
         Text            =   "Reserve4InvoiceMBMPostageFees"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(111) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   110
         Text            =   "InvReceiptDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(112) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   111
         Text            =   "InvReceiptMode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(113) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   112
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(114) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   113
         Text            =   "Excess"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(115) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   114
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(116) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   115
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(117) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   116
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(118) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   117
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(119) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   118
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(120) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   119
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(121) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   120
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(122) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   121
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(123) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   122
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(124) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   123
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(125) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   124
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(126) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   125
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(127) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   126
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(128) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   127
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(129) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   128
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(130) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   129
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(131) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   130
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(132) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   131
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(133) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   132
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(134) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   133
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(135) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   134
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(136) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   135
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(137) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   136
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(138) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   137
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(139) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   138
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(140) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   139
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(141) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   140
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(142) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   141
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(143) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   142
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(144) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   143
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(145) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   144
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(146) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   145
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(147) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   146
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(148) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   147
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(149) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   148
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(150) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   149
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(151) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   150
         Text            =   "CLMOpenDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(152) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   151
         Text            =   "CLMOpenBy"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(153) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   152
         Text            =   "CLMUpdateBy"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(154) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   153
         Text            =   "CLMUpdateDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(155) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   154
         Text            =   "CLMBordxStatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(156) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   155
         Text            =   "CLMStatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(157) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   156
         Text            =   "CLMCloseBy"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(158) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   157
         Text            =   "CLMCloseDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(159) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   158
         Text            =   "TotFCLN"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(160) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   159
         Text            =   "TotCLNExc"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(161) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   160
         Text            =   "TotFinalAct"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(162) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   161
         Text            =   "TotFinalAPP"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(163) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   162
         Text            =   "TotFinalMem"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(164) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   163
         Text            =   "CLMBalLmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(165) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   164
         Text            =   "CLMPrevBal"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(166) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   165
         Text            =   "CLMPaidByM2M"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(167) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   166
         Text            =   "CLMPaidByM2C"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(168) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   167
         Text            =   "CLMActGPBMed"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(169) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   168
         Text            =   "CLMActGPMedInj"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(170) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   169
         Text            =   "CLMActGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(171) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   170
         Text            =   "CLMActSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(172) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   171
         Text            =   "CLMActSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(173) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   172
         Text            =   "CLMActSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(174) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   173
         Text            =   "CLMActXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(175) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   174
         Text            =   "CLMActUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(176) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   175
         Text            =   "CLMActECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(177) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   176
         Text            =   "CLMActDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(178) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   177
         Text            =   "CLMActNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(179) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   178
         Text            =   "CLMActID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(180) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   179
         Text            =   "CLMActTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(181) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   180
         Text            =   "CLMActMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(182) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   181
         Text            =   "CLMActPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(183) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   182
         Text            =   "CLMActBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(184) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   183
         Text            =   "CLMActBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(185) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   184
         Text            =   "CLMActUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(186) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   185
         Text            =   "CLMActPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(187) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   186
         Text            =   "CLMActLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(188) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   187
         Text            =   "CLMActDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(189) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   188
         Text            =   "CLMActDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(190) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   189
         Text            =   "CLMActHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(191) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   190
         Text            =   "CLMActMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(192) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   191
         Text            =   "CLMActIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(193) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   192
         Text            =   "CLMActDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(194) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   193
         Text            =   "CLMActMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(195) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   194
         Text            =   "CLMActPRVCheckup"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(196) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   195
         Text            =   "CLMActVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(197) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   196
         Text            =   "ActTyphoidNonDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(198) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   197
         Text            =   "ActOthersNonDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(199) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   198
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(200) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   199
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(201) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   200
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(202) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   201
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(203) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   202
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(204) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   203
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(205) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   204
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(206) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   205
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(207) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   206
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(208) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   207
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(209) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   208
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(210) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   209
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(211) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   210
         Text            =   "CLMContGPBMed"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(212) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   211
         Text            =   "CLMContGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(213) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   212
         Text            =   "CLMContGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(214) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   213
         Text            =   "CLMContSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(215) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   214
         Text            =   "CLMContSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(216) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   215
         Text            =   "CLMContSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(217) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   216
         Text            =   "CLMContXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(218) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   217
         Text            =   "CLMContUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(219) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   218
         Text            =   "CLMContECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(220) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   219
         Text            =   "CLMContDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(221) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   220
         Text            =   "CLMContNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(222) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   221
         Text            =   "CLMContID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(223) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   222
         Text            =   "CLMContTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(224) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   223
         Text            =   "CLMContMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(225) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   224
         Text            =   "CLMContPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(226) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   225
         Text            =   "CLMContBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(227) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   226
         Text            =   "CLMContBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(228) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   227
         Text            =   "CLMContUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(229) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   228
         Text            =   "CLMContPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(230) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   229
         Text            =   "CLMContLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(231) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   230
         Text            =   "CLMContDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(232) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   231
         Text            =   "CLMContDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(233) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   232
         Text            =   "CLMContHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(234) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   233
         Text            =   "CLMContMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(235) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   234
         Text            =   "CLMContIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(236) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   235
         Text            =   "CLMContDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(237) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   236
         Text            =   "CLMContMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(238) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   237
         Text            =   "CLMContPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(239) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   238
         Text            =   "CLMContVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(240) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   239
         Text            =   "CLMContTyphoidNonDec"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(241) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   240
         Text            =   "CLMContOthersNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(242) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   241
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(243) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   242
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(244) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   243
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(245) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   244
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(246) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   245
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(247) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   246
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(248) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   247
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(249) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   248
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(250) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   249
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(251) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   250
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(252) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   251
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(253) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   252
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(254) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   253
         Text            =   "CLMAppGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(255) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   254
         Text            =   "CLMAppGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(256) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   255
         Text            =   "CLMAppGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(257) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   256
         Text            =   "CLMAppSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(258) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   257
         Text            =   "CLMAppSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(259) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   258
         Text            =   "CLMAppSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(260) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   259
         Text            =   "CLMAppXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(261) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   260
         Text            =   "CLMAppUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(262) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   261
         Text            =   "CLMAppECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(263) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   262
         Text            =   "CLMAppDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(264) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   263
         Text            =   "CLMAppNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(265) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   264
         Text            =   "CLMAppID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(266) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   265
         Text            =   "CLMAppTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(267) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   266
         Text            =   "CLMAppMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(268) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   267
         Text            =   "CLMAppPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(269) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   268
         Text            =   "CLMAppBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(270) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   269
         Text            =   "CLMAppBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(271) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   270
         Text            =   "CLMAppUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(272) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   271
         Text            =   "CLMAppPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(273) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   272
         Text            =   "CLMAppLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(274) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   273
         Text            =   "CLMAppDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(275) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   274
         Text            =   "CLMAppDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(276) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   275
         Text            =   "CLMAppHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(277) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   276
         Text            =   "CLMAppMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(278) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   277
         Text            =   "CLMAppIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(279) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   278
         Text            =   "CLMAppDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(280) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   279
         Text            =   "CLMAppMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(281) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   280
         Text            =   "CLMAppPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(282) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   281
         Text            =   "CLMAppVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(283) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   282
         Text            =   "CLMAppTyphoidNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(284) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   283
         Text            =   "CLMAppOthersNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(285) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   284
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(286) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   285
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(287) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   286
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(288) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   287
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(289) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   288
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(290) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   289
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(291) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   290
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(292) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   291
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(293) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   292
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(294) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   293
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(295) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   294
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(296) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   295
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(297) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   296
         Text            =   "CLMBnfGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(298) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   297
         Text            =   "CLMBnfGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(299) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   298
         Text            =   "CLMBnfGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(300) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   299
         Text            =   "CLMBnfSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(301) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   300
         Text            =   "CLMBnfSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(302) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   301
         Text            =   "CLMBnfSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(303) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   302
         Text            =   "CLMBnfXray"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(304) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   303
         Text            =   "CLMBnfUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(305) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   304
         Text            =   "CLMBnfECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(306) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   305
         Text            =   "CLMBnfDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(307) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   306
         Text            =   "CLMBnfNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(308) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   307
         Text            =   "CLMBnfID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(309) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   308
         Text            =   "CLMBnfTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(310) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   309
         Text            =   "CLMBnfMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(311) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   310
         Text            =   "CLMBnfPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(312) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   311
         Text            =   "CLMBnfBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(313) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   312
         Text            =   "CLMBnfBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(314) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   313
         Text            =   "CLMBnfUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(315) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   314
         Text            =   "CLMBnfPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(316) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   315
         Text            =   "CLMBnfLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(317) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   316
         Text            =   "CLMBnfDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(318) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   317
         Text            =   "CLMBnfDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(319) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   318
         Text            =   "CLMBnfHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(320) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   319
         Text            =   "CLMBnfMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(321) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   320
         Text            =   "CLMBnfIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(322) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   321
         Text            =   "CLMBnfDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(323) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   322
         Text            =   "CLMBnfMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(324) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   323
         Text            =   "CLMBnfPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(325) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   324
         Text            =   "CLMBnfVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(326) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   325
         Text            =   "CLMBnfTyphoidNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(327) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   326
         Text            =   "CLMBnfOthersNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(328) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   327
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(329) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   328
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(330) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   329
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(331) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   330
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(332) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   331
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(333) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   332
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(334) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   333
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(335) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   334
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(336) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   335
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(337) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   336
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(338) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   337
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(339) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   338
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(340) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   339
         Text            =   "CLMMBMExcGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(341) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   340
         Text            =   "CLMMBMExcGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(342) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   341
         Text            =   "CLMMBMExcGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(343) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   342
         Text            =   "CLMMBMExcSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(344) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   343
         Text            =   "CLMMBMExcSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(345) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   344
         Text            =   "CLMMBMExcSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(346) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   345
         Text            =   "CLMMBMExcXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(347) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   346
         Text            =   "CLMMBMExcUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(348) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   347
         Text            =   "CLMMBMExcECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(349) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   348
         Text            =   "CLMMBMExcDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(350) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   349
         Text            =   "CLMMBMExcNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(351) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   350
         Text            =   "CLMMBMExcID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(352) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   351
         Text            =   "CLMMBMExcTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(353) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   352
         Text            =   "CLMMBMExcMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(354) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   353
         Text            =   "CLMMBMExcPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(355) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   354
         Text            =   "CLMMBMExcBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(356) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   355
         Text            =   "CLMMBMExcBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(357) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   356
         Text            =   "CLMMBMExcUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(358) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   357
         Text            =   "CLMMBMExcPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(359) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   358
         Text            =   "CLMMBMExcLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(360) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   359
         Text            =   "CLMMBMExcDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(361) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   360
         Text            =   "CLMMBMExcDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(362) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   361
         Text            =   "CLMMBMExcHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(363) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   362
         Text            =   "CLMMBMExcMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(364) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   363
         Text            =   "CLMMBMExcIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(365) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   364
         Text            =   "CLMMBMExcDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(366) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   365
         Text            =   "CLMMBMExcMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(367) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   366
         Text            =   "CLMMBMExcPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(368) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   367
         Text            =   "CLMMBMExcVision"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(369) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   368
         Text            =   "CLMMBMTyphoidNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(370) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   369
         Text            =   "CLMMBMOthersNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(371) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   370
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(372) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   371
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(373) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   372
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(374) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   373
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(375) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   374
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(376) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   375
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(377) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   376
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(378) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   377
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(379) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   378
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(380) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   379
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(381) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   380
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(382) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   381
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(383) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   382
         Text            =   "CLMNPayCLNGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(384) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   383
         Text            =   "CLMNPayCLNGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(385) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   384
         Text            =   "CLMNPayCLNGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(386) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   385
         Text            =   "CLMNPayCLNSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(387) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   386
         Text            =   "CLMNPayCLNSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(388) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   387
         Text            =   "CLMNPayCLNSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(389) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   388
         Text            =   "CLMNPayCLNXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(390) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   389
         Text            =   "CLMNPayCLNUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(391) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   390
         Text            =   "CLMNPayCLNECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(392) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   391
         Text            =   "CLMNPayCLNDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(393) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   392
         Text            =   "CLMNPayCLNNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(394) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   393
         Text            =   "CLMNPayCLNID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(395) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   394
         Text            =   "CLMNPayCLNTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(396) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   395
         Text            =   "CLMNPayCLNMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(397) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   396
         Text            =   "CLMNPayCLNPThers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(398) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   397
         Text            =   "CLMNPayCLNBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(399) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   398
         Text            =   "CLMNPayCLNBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(400) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   399
         Text            =   "CLMNPayCLNUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(401) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   400
         Text            =   "CLMNPayCLNPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(402) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   401
         Text            =   "CLMNPayCLNLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(403) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   402
         Text            =   "CLMNPayCLNDTPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(404) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   403
         Text            =   "CLMNPayCLNDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(405) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   404
         Text            =   "CLMNPayCLNHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(406) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   405
         Text            =   "CLMNPayCLNMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(407) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   406
         Text            =   "CLMNPayCLNIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(408) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   407
         Text            =   "CLMNPayCLNDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(409) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   408
         Text            =   "CLMNPayCLNMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(410) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   409
         Text            =   "CLMNPayCLNPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(411) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   410
         Text            =   "CLMNPayCLNVision"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(412) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   411
         Text            =   "CLMNPayCLNTyphoid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(413) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   412
         Text            =   "CLMNPayCLNOthersNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(414) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   413
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(415) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   414
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(416) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   415
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(417) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   416
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(418) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   417
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(419) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   418
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(420) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   419
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(421) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   420
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(422) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   421
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(423) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   422
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(424) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   423
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(425) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   424
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(426) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   425
         Text            =   "CLMPayCLNGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(427) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   426
         Text            =   "CLMPayCLNGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(428) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   427
         Text            =   "CLMPayCLNGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(429) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   428
         Text            =   "CLMPayCLNSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(430) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   429
         Text            =   "CLMPayCLNSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(431) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   430
         Text            =   "CLMPayCLNSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(432) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   431
         Text            =   "CLMPayCLNXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(433) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   432
         Text            =   "CLMPayCLNUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(434) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   433
         Text            =   "CLMPayCLNECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(435) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   434
         Text            =   "CLMPayCLNDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(436) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   435
         Text            =   "CLMPayCLNNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(437) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   436
         Text            =   "CLMPayCLNID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(438) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   437
         Text            =   "CLMPayCLNTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(439) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   438
         Text            =   "CLMPayCLNMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(440) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   439
         Text            =   "CLMPayCLNPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(441) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   440
         Text            =   "CLMPayCLNBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(442) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   441
         Text            =   "CLMPayCLNBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(443) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   442
         Text            =   "CLMPayCLNUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(444) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   443
         Text            =   "CLMPayCLNPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(445) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   444
         Text            =   "CLMPayCLNLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(446) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   445
         Text            =   "CLMPayCLNDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(447) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   446
         Text            =   "CLMPayCLNDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(448) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   447
         Text            =   "CLMPayCLNHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(449) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   448
         Text            =   "CLMPayCLNMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(450) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   449
         Text            =   "CLMPayCLNIOhters"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(451) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   450
         Text            =   "CLMPayCLNDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(452) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   451
         Text            =   "CLMPayCLNMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(453) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   452
         Text            =   "CLMPayCLNPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(454) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   453
         Text            =   "CLMPayCLNVision"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(455) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   454
         Text            =   "CLMPayCLNTyphoidNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(456) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   455
         Text            =   "CLMPayCLNOthersNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(457) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   456
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(458) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   457
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(459) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   458
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(460) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   459
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(461) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   460
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(462) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   461
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(463) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   462
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(464) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   463
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(465) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   464
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(466) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   465
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(467) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   466
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(468) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   467
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(469) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   468
         Text            =   "MCOProcFee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(470) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   469
         Text            =   "MCOProcFeeStatus"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(471) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   470
         Text            =   "InvProcPeriod"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(472) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   471
         Text            =   "CoPayCLNColl"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(473) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   472
         Text            =   "CLNBillAmt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(474) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   473
         Text            =   "Claimability"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(475) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   474
         Text            =   "TyphoidNonDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(476) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   475
         Text            =   "OthersNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(477) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   476
         Text            =   "RBAmt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(478) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   477
         Text            =   "RBCoPay"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(479) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   478
         Text            =   "LimitExceeded"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(480) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   479
         Text            =   "Excess"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame FrameInvoice 
      Height          =   2115
      Left            =   4440
      TabIndex        =   156
      Top             =   240
      Width           =   4095
      Begin VB.TextBox txtMBMNo 
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
         Left            =   750
         TabIndex        =   6
         Top             =   230
         Width           =   3260
      End
      Begin VB.TextBox txtPatientName 
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
         Left            =   750
         TabIndex        =   189
         Top             =   480
         Width           =   3260
      End
      Begin VB.Frame FrameMBMInfo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2235
         Left            =   600
         TabIndex        =   173
         Top             =   2400
         Visible         =   0   'False
         Width           =   2370
         Begin VB.TextBox TxtPayor 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   181
            Top             =   1440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtPolNo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   180
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtTabMBMNo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   179
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtdob 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   178
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtPlanCode 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   177
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox TxtInsType 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   176
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtMBMStatus 
            Height          =   285
            Left            =   1920
            TabIndex        =   175
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtTabCoverID 
            Height          =   285
            Left            =   1920
            TabIndex        =   174
            Text            =   "Text1"
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label64 
            Caption         =   "Plan Code"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   188
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label62 
            Caption         =   "InsType"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   187
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label61 
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
            TabIndex        =   186
            Top             =   1440
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label60 
            Caption         =   "Member No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   185
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label37 
            Caption         =   "Patient DOB"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   184
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Policy No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   183
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label102 
            Caption         =   "Status"
            Height          =   255
            Left            =   1560
            TabIndex        =   182
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.TextBox txtPatientIC 
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
         Left            =   750
         TabIndex        =   7
         Top             =   770
         Width           =   3260
      End
      Begin VB.TextBox txtDept 
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
         Left            =   750
         TabIndex        =   241
         Top             =   1050
         Width           =   3260
      End
      Begin MSMask.MaskEdBox txtEffDate 
         Height          =   285
         Left            =   750
         TabIndex        =   238
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
      Begin MSMask.MaskEdBox txtExpDate 
         Height          =   285
         Left            =   750
         TabIndex        =   237
         Top             =   1560
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
      Begin VB.TextBox txtMCDaysBal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   750
         TabIndex        =   243
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2400
         TabIndex        =   249
         Top             =   1320
         Width           =   1500
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&Search"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   645
            TabIndex        =   252
            Top             =   140
            Width           =   705
         End
         Begin VB.CommandButton cmdGo 
            Caption         =   "&Go"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   80
            TabIndex        =   251
            Top             =   120
            Width           =   585
         End
         Begin VB.CommandButton cmdUpdateDept 
            Caption         =   "Dept"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   645
            TabIndex        =   250
            Top             =   360
            Width           =   705
         End
         Begin VB.CommandButton cmdRegMBM 
            Caption         =   "&MBM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   80
            TabIndex        =   253
            Top             =   360
            Width           =   585
         End
      End
      Begin VB.Label Label2 
         Caption         =   "MC"
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
         Left            =   75
         TabIndex        =   244
         Top             =   1835
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Dept"
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
         Left            =   75
         TabIndex        =   242
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Eff Date"
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
         Left            =   75
         TabIndex        =   240
         Top             =   1370
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "Exp Date"
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
         Left            =   75
         TabIndex        =   239
         Top             =   1605
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
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
         Left            =   75
         TabIndex        =   170
         Top             =   520
         Width           =   615
      End
      Begin VB.Label Label75 
         Caption         =   "IC No"
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
         Left            =   75
         TabIndex        =   172
         Top             =   800
         Width           =   735
      End
      Begin VB.Label Label54 
         Caption         =   "MBM No"
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
         Left            =   80
         TabIndex        =   171
         Top             =   250
         Width           =   855
      End
   End
   Begin VB.Frame FrameDis 
      Enabled         =   0   'False
      Height          =   2055
      Left            =   8520
      TabIndex        =   191
      Top             =   240
      Width           =   3735
      Begin VB.TextBox txtGPLimit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   291
         Top             =   1680
         Width           =   660
      End
      Begin VB.TextBox txtSPLimit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   289
         Top             =   1680
         Width           =   660
      End
      Begin VB.TextBox txtDentalLmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   285
         Top             =   1200
         Width           =   660
      End
      Begin VB.TextBox txtDentBal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   284
         Top             =   1440
         Width           =   660
      End
      Begin VB.TextBox txtSpecBal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   283
         Top             =   960
         Width           =   660
      End
      Begin VB.TextBox txtSpecLmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   281
         Top             =   720
         Width           =   660
      End
      Begin VB.TextBox txtMatBal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   279
         Top             =   480
         Width           =   660
      End
      Begin VB.TextBox txtMatLmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   277
         Top             =   160
         Width           =   660
      End
      Begin VB.TextBox txtPolStatus 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   192
         Top             =   160
         Width           =   660
      End
      Begin VB.TextBox txtAnnLmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   193
         Top             =   480
         Width           =   660
      End
      Begin VB.TextBox txtAccumAmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   194
         Top             =   720
         Width           =   660
      End
      Begin VB.TextBox TxtBalLimit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   195
         Top             =   960
         Width           =   660
      End
      Begin VB.TextBox TxtVisitLimit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   199
         Top             =   1200
         Width           =   660
      End
      Begin VB.TextBox txtDentStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   196
         Top             =   680
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox txtLTM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   197
         Top             =   920
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox TxtCoPay 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   198
         Top             =   1170
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label38 
         Caption         =   "SP/GP Limit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   290
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label34 
         Caption         =   "Dental Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   288
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label34 
         Caption         =   "Specialist Bal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   286
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label34 
         Caption         =   "Specialist Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   282
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label34 
         Caption         =   "Maternity  Bal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   280
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label34 
         Caption         =   "Maternity Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   278
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label34 
         Caption         =   "Ann. Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   203
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label38 
         Caption         =   "Vst Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   202
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label40 
         Caption         =   "Bal. Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   201
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label68 
         Caption         =   "Pol. Sta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   200
         Top             =   200
         Width           =   735
      End
      Begin VB.Label Label112 
         Caption         =   "Co-Pay"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   206
         Top             =   1250
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label39 
         Caption         =   "Acc Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   204
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label69 
         BackColor       =   &H80000013&
         Caption         =   "LTM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   135
         Left            =   120
         TabIndex        =   205
         Top             =   1000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label111 
         BackColor       =   &H80000013&
         Caption         =   "Dent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   207
         Top             =   735
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label38 
         Caption         =   "Dental Bal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   287
         Top             =   1440
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstErr 
      Height          =   255
      Left            =   0
      TabIndex        =   99
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   445
      _ExtentY        =   445
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "MBM No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cov"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Policy No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Patient IC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Eff Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Exp. Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bal. Limit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Visit Lmt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Err Msg"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame FrameMBMCLN 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   0
      TabIndex        =   98
      Top             =   240
      Width           =   4425
      Begin VB.TextBox txtClinicName 
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
         Height          =   315
         Left            =   1320
         TabIndex        =   208
         Top             =   230
         Width           =   3045
      End
      Begin VB.TextBox TxtClinicCode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   720
         MaxLength       =   8
         TabIndex        =   0
         Top             =   230
         Width           =   615
      End
      Begin VB.ComboBox CboInvRecMode 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         TabIndex        =   164
         Top             =   1275
         Visible         =   0   'False
         Width           =   3110
      End
      Begin MSMask.MaskEdBox txtInvDate 
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Width           =   915
         _ExtentX        =   1609
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
      Begin MSMask.MaskEdBox txtInvRecDate 
         Height          =   300
         Left            =   2280
         TabIndex        =   5
         Top             =   720
         Width           =   915
         _ExtentX        =   1609
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
      Begin VB.TextBox txtInvNo 
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
         Left            =   720
         MaxLength       =   30
         TabIndex        =   1
         Top             =   480
         Width           =   945
      End
      Begin VB.TextBox TxtTotInvoiceAmt 
         Alignment       =   1  'Right Justify
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
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   3
         Top             =   480
         Width           =   765
      End
      Begin VB.ComboBox cboProcMth 
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
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label101 
         Caption         =   "Mth"
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
         Index           =   5
         Left            =   120
         TabIndex        =   223
         ToolTipText     =   "Processing Month"
         Top             =   795
         Width           =   735
      End
      Begin VB.Label lblClnType 
         Caption         =   "lblClnType"
         Height          =   255
         Left            =   3360
         TabIndex        =   209
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Amt"
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
         TabIndex        =   169
         ToolTipText     =   "Invoice Amount"
         Top             =   555
         Width           =   735
      End
      Begin VB.Label Label35 
         Caption         =   "Inv. No"
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
         TabIndex        =   168
         Top             =   550
         Width           =   615
      End
      Begin VB.Label Label67 
         Caption         =   "InvDate"
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
         Left            =   1680
         TabIndex        =   167
         Top             =   555
         Width           =   615
      End
      Begin VB.Label Label45 
         Caption         =   "RecDate"
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
         Left            =   1680
         TabIndex        =   166
         Top             =   795
         Width           =   735
      End
      Begin VB.Label Label70 
         Caption         =   "RecMode"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   165
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Clinic"
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
         TabIndex        =   101
         Top             =   280
         Width           =   495
      End
   End
   Begin VB.Frame FrameInv 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   0
      TabIndex        =   102
      Top             =   1400
      Width           =   4425
      Begin VB.ComboBox cboClaimability 
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
         Left            =   3120
         TabIndex        =   232
         Top             =   1140
         Visible         =   0   'False
         Width           =   1875
      End
      Begin MSMask.MaskEdBox txtVisitTime 
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   495
         _ExtentX        =   868
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtMC 
         Alignment       =   1  'Right Justify
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
         Left            =   3360
         MaxLength       =   3
         TabIndex        =   12
         Top             =   225
         Width           =   875
      End
      Begin VB.TextBox txtCLNBillAmt 
         Alignment       =   1  'Right Justify
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
         Left            =   720
         MaxLength       =   20
         TabIndex        =   9
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox txtBillNo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         MaxLength       =   20
         TabIndex        =   8
         Top             =   180
         Visible         =   0   'False
         Width           =   645
      End
      Begin MSMask.MaskEdBox txtVisitDate 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Top             =   465
         Width           =   1005
         _ExtentX        =   1778
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
      Begin VB.ComboBox cboFFS 
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
         Left            =   2400
         TabIndex        =   247
         Top             =   480
         Width           =   885
      End
      Begin VB.Label Label101 
         Caption         =   "FFS"
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
         Index           =   9
         Left            =   1800
         TabIndex        =   248
         Top             =   555
         Width           =   495
      End
      Begin VB.Label Label101 
         Caption         =   "Clm'lity"
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
         Index           =   16
         Left            =   2400
         TabIndex        =   233
         Top             =   1185
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label101 
         Caption         =   "Bill Amt"
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
         Index           =   15
         Left            =   120
         TabIndex        =   231
         Top             =   225
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "MC"
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
         Left            =   3000
         TabIndex        =   103
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label66 
         Caption         =   "Vst Time"
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
         Left            =   1800
         TabIndex        =   104
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Vst Date"
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
         TabIndex        =   106
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Bill No"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   220
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame FrameDet 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   3720
      TabIndex        =   120
      Top             =   2760
      Width           =   8625
      Begin VB.TextBox txtExcess 
         Alignment       =   1  'Right Justify
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
         Left            =   7550
         MaxLength       =   8
         TabIndex        =   264
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtRBAmt 
         Alignment       =   1  'Right Justify
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
         Left            =   7550
         MaxLength       =   8
         TabIndex        =   254
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox A32 
         Alignment       =   1  'Right Justify
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
         Left            =   3180
         MaxLength       =   8
         TabIndex        =   61
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox A33 
         Alignment       =   1  'Right Justify
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
         Left            =   3180
         MaxLength       =   8
         TabIndex        =   62
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox TxtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   180
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   78
         Top             =   3480
         Width           =   3495
      End
      Begin VB.TextBox A10 
         Alignment       =   1  'Right Justify
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
         Left            =   3165
         MaxLength       =   8
         TabIndex        =   51
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox A19 
         Alignment       =   1  'Right Justify
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
         Left            =   5565
         MaxLength       =   8
         TabIndex        =   63
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox A20 
         Alignment       =   1  'Right Justify
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
         Left            =   5565
         MaxLength       =   8
         TabIndex        =   64
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox A21 
         Alignment       =   1  'Right Justify
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
         Left            =   5565
         MaxLength       =   8
         TabIndex        =   65
         Top             =   855
         Width           =   495
      End
      Begin VB.TextBox A22 
         Alignment       =   1  'Right Justify
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
         Left            =   5565
         MaxLength       =   8
         TabIndex        =   66
         Top             =   1110
         Width           =   495
      End
      Begin VB.TextBox A24 
         Alignment       =   1  'Right Justify
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
         Left            =   5565
         MaxLength       =   8
         TabIndex        =   69
         Top             =   1755
         Width           =   495
      End
      Begin VB.TextBox A25 
         Alignment       =   1  'Right Justify
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
         Left            =   5565
         MaxLength       =   8
         TabIndex        =   70
         Top             =   1995
         Width           =   495
      End
      Begin VB.TextBox A26 
         Alignment       =   1  'Right Justify
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
         Left            =   5565
         MaxLength       =   8
         TabIndex        =   71
         Top             =   2265
         Width           =   495
      End
      Begin VB.TextBox A27 
         Alignment       =   1  'Right Justify
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
         Left            =   5565
         MaxLength       =   8
         TabIndex        =   72
         Top             =   2505
         Width           =   495
      End
      Begin VB.TextBox A28 
         Alignment       =   1  'Right Justify
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
         Left            =   5565
         MaxLength       =   8
         TabIndex        =   73
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtIOthDesc 
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
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   74
         Top             =   2790
         Width           =   1335
      End
      Begin VB.TextBox A11 
         Alignment       =   1  'Right Justify
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
         Left            =   3165
         MaxLength       =   8
         TabIndex        =   52
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox A12 
         Alignment       =   1  'Right Justify
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
         Left            =   3165
         MaxLength       =   8
         TabIndex        =   53
         Top             =   855
         Width           =   495
      End
      Begin VB.TextBox A13 
         Alignment       =   1  'Right Justify
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
         Left            =   3165
         MaxLength       =   8
         TabIndex        =   54
         Top             =   1095
         Width           =   495
      End
      Begin VB.TextBox A14 
         Alignment       =   1  'Right Justify
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
         Left            =   3165
         MaxLength       =   8
         TabIndex        =   55
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox B1 
         Alignment       =   1  'Right Justify
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
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   41
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox B5 
         Alignment       =   1  'Right Justify
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
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   46
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox A15 
         Alignment       =   1  'Right Justify
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
         Left            =   3165
         MaxLength       =   8
         TabIndex        =   56
         Top             =   1605
         Width           =   495
      End
      Begin VB.TextBox A16 
         Alignment       =   1  'Right Justify
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
         Left            =   3165
         MaxLength       =   8
         TabIndex        =   57
         Top             =   1845
         Width           =   495
      End
      Begin VB.TextBox A17 
         Alignment       =   1  'Right Justify
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
         Left            =   3165
         MaxLength       =   8
         TabIndex        =   58
         Top             =   2115
         Width           =   495
      End
      Begin VB.TextBox A18 
         Alignment       =   1  'Right Justify
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
         Left            =   3165
         MaxLength       =   8
         TabIndex        =   59
         Top             =   2355
         Width           =   495
      End
      Begin VB.TextBox B4 
         Alignment       =   1  'Right Justify
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
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   42
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox B3 
         Alignment       =   1  'Right Justify
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
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   43
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox B2 
         Alignment       =   1  'Right Justify
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
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   44
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox B8 
         Alignment       =   1  'Right Justify
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
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   47
         Top             =   2190
         Width           =   495
      End
      Begin VB.TextBox B7 
         Alignment       =   1  'Right Justify
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
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   48
         Top             =   2430
         Width           =   495
      End
      Begin VB.TextBox B6 
         Alignment       =   1  'Right Justify
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
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   49
         Top             =   2670
         Width           =   495
      End
      Begin VB.TextBox txtPOthDesc 
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   60
         Top             =   2400
         Width           =   1140
      End
      Begin VB.TextBox C1 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   45
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox C2 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   50
         Top             =   2955
         Width           =   495
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1975
         Left            =   6240
         TabIndex        =   210
         Top             =   2160
         Width           =   2295
         Begin VB.TextBox A29 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   880
            MaxLength       =   8
            TabIndex        =   77
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label101 
            Caption         =   "Actual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   1605
            TabIndex        =   230
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label101 
            Caption         =   "Clinic"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   885
            TabIndex        =   229
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblActCoPay 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1560
            TabIndex        =   228
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblDispCLNBillAmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   880
            TabIndex        =   227
            Top             =   420
            Width           =   615
         End
         Begin VB.Label lblCLNBillAmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   880
            TabIndex        =   226
            Top             =   1005
            Width           =   615
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   1575
            X2              =   2175
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label lblPay2CLN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1560
            TabIndex        =   219
            Top             =   1605
            Width           =   615
         End
         Begin VB.Label Label76 
            Caption         =   "Amt Pay"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   285
            Index           =   3
            Left            =   75
            TabIndex        =   218
            Top             =   1620
            Width           =   1095
         End
         Begin VB.Label lblFFS 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1560
            TabIndex        =   217
            Top             =   1275
            Width           =   615
         End
         Begin VB.Label Label76 
            Caption         =   "Less:FFS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   2
            Left            =   75
            TabIndex        =   216
            Top             =   1260
            Width           =   855
         End
         Begin VB.Label Label76 
            Caption         =   "Bill Amt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   1
            Left            =   75
            TabIndex        =   215
            Top             =   975
            Width           =   735
         End
         Begin VB.Label lblBillAmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1560
            TabIndex        =   214
            Top             =   1005
            Width           =   615
         End
         Begin VB.Label Label101 
            Caption         =   "Co-Pymt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   75
            TabIndex        =   213
            Top             =   690
            Width           =   735
         End
         Begin VB.Label lblDispBillAmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1560
            TabIndex        =   212
            Top             =   420
            Width           =   615
         End
         Begin VB.Label Label76 
            Caption         =   "Total Bill "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   0
            Left            =   75
            TabIndex        =   211
            Top             =   420
            Width           =   735
         End
      End
      Begin VB.TextBox A2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   85
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox A1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   84
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox A3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   86
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox A6 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   89
         Top             =   2880
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox A5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   88
         Top             =   2880
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox A4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   87
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtMBMAdjAnnLmt 
         Alignment       =   1  'Right Justify
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
         Left            =   7080
         MaxLength       =   12
         TabIndex        =   76
         Top             =   3240
         Width           =   555
      End
      Begin VB.TextBox A31 
         Alignment       =   1  'Right Justify
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
         Left            =   6840
         MaxLength       =   12
         TabIndex        =   75
         Top             =   2640
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   7550
         MaxLength       =   8
         TabIndex        =   263
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtRBAmtExc 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   7550
         MaxLength       =   8
         TabIndex        =   260
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox A7 
         Alignment       =   1  'Right Justify
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
         Left            =   5535
         MaxLength       =   8
         TabIndex        =   266
         Top             =   3120
         Width           =   555
      End
      Begin VB.TextBox A8 
         Alignment       =   1  'Right Justify
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
         Left            =   5535
         MaxLength       =   8
         TabIndex        =   267
         Top             =   3360
         Width           =   555
      End
      Begin VB.TextBox A9 
         Alignment       =   1  'Right Justify
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
         Left            =   5535
         MaxLength       =   8
         TabIndex        =   268
         Top             =   3615
         Width           =   555
      End
      Begin VB.TextBox A30 
         Alignment       =   1  'Right Justify
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
         Left            =   5535
         MaxLength       =   8
         TabIndex        =   269
         Top             =   3885
         Width           =   555
      End
      Begin VB.TextBox txtLOthDesc 
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
         Left            =   6720
         MaxLength       =   50
         TabIndex        =   68
         Top             =   2400
         Width           =   1365
      End
      Begin VB.TextBox A23 
         Alignment       =   1  'Right Justify
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
         Left            =   7680
         MaxLength       =   8
         TabIndex        =   67
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtRBCoPay 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   7550
         MaxLength       =   8
         TabIndex        =   256
         Top             =   1125
         Width           =   975
      End
      Begin VB.TextBox txtLmtExceeded 
         Alignment       =   1  'Right Justify
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
         Left            =   7550
         MaxLength       =   8
         TabIndex        =   258
         Top             =   1395
         Width           =   975
      End
      Begin VB.Label Label101 
         Caption         =   "Dental"
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
         Index           =   3
         Left            =   3840
         TabIndex        =   273
         Top             =   3165
         Width           =   735
      End
      Begin VB.Label Label101 
         Caption         =   "Maternity"
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
         Index           =   4
         Left            =   3840
         TabIndex        =   272
         Top             =   3405
         Width           =   855
      End
      Begin VB.Label Label101 
         Caption         =   "PRV Chk."
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
         Index           =   11
         Left            =   3840
         TabIndex        =   271
         Top             =   3645
         Width           =   855
      End
      Begin VB.Label Label101 
         Caption         =   "Vision"
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
         Index           =   12
         Left            =   3840
         TabIndex        =   270
         Top             =   3915
         Width           =   855
      End
      Begin VB.Label Label85 
         Caption         =   "Total Excess"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   6240
         TabIndex        =   265
         Top             =   1800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label85 
         Caption         =   "R&&B Excess"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   6240
         TabIndex        =   262
         Top             =   615
         Width           =   1575
      End
      Begin VB.Label Label85 
         Caption         =   "R&&B Co-Pay"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6240
         TabIndex        =   261
         Top             =   1155
         Width           =   1575
      End
      Begin VB.Label Label85 
         Caption         =   "Limit Exceeded"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   6240
         TabIndex        =   259
         Top             =   1425
         Width           =   1575
      End
      Begin VB.Label Label85 
         Caption         =   "Other Excess"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6240
         TabIndex        =   257
         Top             =   870
         Width           =   1575
      End
      Begin VB.Label Label85 
         Caption         =   "Room && Board"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   255
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label90 
         Caption         =   "Typhoid Jab"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1620
         TabIndex        =   236
         Top             =   2940
         Width           =   915
      End
      Begin VB.Label txtOther 
         Caption         =   "Others"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1620
         TabIndex        =   235
         Top             =   3180
         Width           =   1095
      End
      Begin VB.Label Label73 
         Caption         =   "No Limit Deduction"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1620
         TabIndex        =   234
         ToolTipText     =   "Geberal Pratitioner"
         Top             =   2700
         Width           =   1875
      End
      Begin VB.Label Label88 
         Caption         =   "B. Drawing"
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
         Left            =   3840
         TabIndex        =   139
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label101 
         Caption         =   "Specialist"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   222
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label101 
         Caption         =   "Total"
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
         Index           =   7
         Left            =   240
         TabIndex        =   221
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label101 
         Caption         =   "Total"
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
         Index           =   6
         Left            =   240
         TabIndex        =   220
         Top             =   1370
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   190
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label118 
         Caption         =   "Med."
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
         TabIndex        =   154
         Top             =   2205
         Width           =   375
      End
      Begin VB.Label Label117 
         Caption         =   "Inj."
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
         TabIndex        =   153
         Top             =   2475
         Width           =   315
      End
      Begin VB.Label Label116 
         Caption         =   "Med."
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
         TabIndex        =   152
         Top             =   640
         Width           =   375
      End
      Begin VB.Label Label115 
         Caption         =   "Inj."
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
         TabIndex        =   151
         Top             =   880
         Width           =   375
      End
      Begin VB.Label Label77 
         Caption         =   "Fee + LT Med."
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
         TabIndex        =   150
         Top             =   1950
         Width           =   255
      End
      Begin VB.Label Label109 
         Caption         =   "LT Med."
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
         TabIndex        =   149
         Top             =   2715
         Width           =   735
      End
      Begin VB.Label Label74 
         Caption         =   "Fee + LT Med."
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
         TabIndex        =   148
         Top             =   380
         Width           =   255
      End
      Begin VB.Label Label108 
         Caption         =   "LT Med."
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
         TabIndex        =   147
         Top             =   1120
         Width           =   735
      End
      Begin VB.Label Label32 
         Caption         =   "DTP&&Polio"
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
         Left            =   3840
         TabIndex        =   144
         Top             =   1755
         Width           =   1575
      End
      Begin VB.Label Label65 
         Caption         =   "Pap Smear"
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
         Left            =   3840
         TabIndex        =   143
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label Label82 
         Caption         =   "Urine"
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
         Left            =   3840
         TabIndex        =   142
         Top             =   915
         Width           =   735
      End
      Begin VB.Label Label83 
         Caption         =   "B. Glucose"
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
         Left            =   3840
         TabIndex        =   141
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label84 
         Caption         =   "Laboratory"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   140
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "M. Surg."
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
         Left            =   1680
         TabIndex        =   138
         Top             =   2070
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "Oth"
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
         Left            =   1680
         TabIndex        =   137
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label27 
         Caption         =   "I&&D"
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
         Left            =   1680
         TabIndex        =   136
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label Label31 
         Caption         =   "T&&S/Sto"
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
         Left            =   1680
         TabIndex        =   135
         Top             =   1830
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Nebulizer"
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
         Left            =   1680
         TabIndex        =   134
         Top             =   1395
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Dressing"
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
         Left            =   1680
         TabIndex        =   133
         Top             =   1155
         Width           =   735
      End
      Begin VB.Label Label58 
         Caption         =   "ECG"
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
         Left            =   1680
         TabIndex        =   129
         Top             =   915
         Width           =   1095
      End
      Begin VB.Label Label63 
         Caption         =   "Utr Sound"
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
         Left            =   1680
         TabIndex        =   128
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Lbl45days 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "45 Days Late"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   1005
         TabIndex        =   127
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label LblExpDate 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Expired"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1005
         TabIndex        =   126
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label85 
         Caption         =   "MMR"
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
         Left            =   3840
         TabIndex        =   125
         Top             =   2505
         Width           =   1095
      End
      Begin VB.Label Label86 
         Caption         =   "Hepatitis B"
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
         Left            =   3840
         TabIndex        =   124
         Top             =   2235
         Width           =   975
      End
      Begin VB.Label Label87 
         Caption         =   "DTP&&Polio&&HIB"
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
         Left            =   3840
         TabIndex        =   123
         Top             =   2010
         Width           =   1455
      End
      Begin VB.Label Label89 
         Caption         =   "Immunization"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   122
         Top             =   1485
         Width           =   1455
      End
      Begin VB.Label Label90 
         Caption         =   "Oth"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   3840
         TabIndex        =   121
         Top             =   2805
         Width           =   1095
      End
      Begin VB.Label Label73 
         Caption         =   "GP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   131
         ToolTipText     =   "Geberal Pratitioner"
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   "X-Ray "
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
         Left            =   1680
         TabIndex        =   132
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label79 
         Caption         =   "Procedures"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   130
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame FrameDiag 
      Caption         =   "Diagnosis"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   0
      TabIndex        =   107
      Top             =   2520
      Width           =   3495
      Begin VB.CommandButton cmdSearchICD5 
         Caption         =   ">"
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
         Left            =   2925
         TabIndex        =   27
         Top             =   1365
         Width           =   255
      End
      Begin VB.CommandButton cmdSearchICD4 
         Caption         =   ">"
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
         Left            =   2925
         TabIndex        =   24
         Top             =   1050
         Width           =   255
      End
      Begin VB.CommandButton cmdSearchICD1 
         Caption         =   ">"
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
         Left            =   2925
         TabIndex        =   15
         Top             =   205
         Width           =   255
      End
      Begin VB.CommandButton cmdSearchICD3 
         Caption         =   ">"
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
         Left            =   2925
         TabIndex        =   21
         Top             =   765
         Width           =   255
      End
      Begin VB.CommandButton cmdSearchICD2 
         Caption         =   ">"
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
         Left            =   2925
         TabIndex        =   18
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtDiagnosis1 
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
         Left            =   315
         MaxLength       =   10
         TabIndex        =   13
         Top             =   205
         Width           =   495
      End
      Begin VB.TextBox txtDiagnosis2 
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
         Left            =   315
         MaxLength       =   10
         TabIndex        =   16
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtDiagnosis3 
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
         Left            =   315
         MaxLength       =   10
         TabIndex        =   19
         Top             =   765
         Width           =   495
      End
      Begin VB.TextBox txtDiagnosis4 
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
         Left            =   315
         MaxLength       =   10
         TabIndex        =   22
         Top             =   1035
         Width           =   495
      End
      Begin VB.TextBox TxtDiagnosis5 
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
         Left            =   320
         MaxLength       =   100
         TabIndex        =   25
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox TxtDiagDesc1 
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
         Height          =   315
         Left            =   800
         MaxLength       =   10
         TabIndex        =   14
         Top             =   205
         Width           =   2055
      End
      Begin VB.TextBox TxtDiagDesc2 
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
         Height          =   315
         Left            =   800
         MaxLength       =   10
         TabIndex        =   17
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox TxtDiagDesc3 
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
         Height          =   315
         Left            =   800
         MaxLength       =   10
         TabIndex        =   20
         Top             =   765
         Width           =   2055
      End
      Begin VB.TextBox TxtDiagDesc4 
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
         Height          =   315
         Left            =   800
         MaxLength       =   10
         TabIndex        =   23
         Top             =   1035
         Width           =   2055
      End
      Begin VB.TextBox TxtDiagnosisOth 
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
         Height          =   315
         Left            =   800
         MaxLength       =   100
         TabIndex        =   26
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox TxtDiagDesc5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   800
         MaxLength       =   10
         TabIndex        =   225
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "5"
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
         TabIndex        =   155
         Top             =   1335
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "4"
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
         TabIndex        =   111
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label52 
         Caption         =   "3"
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
         TabIndex        =   110
         Top             =   795
         Width           =   255
      End
      Begin VB.Label Label51 
         Caption         =   "2"
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
         TabIndex        =   109
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "1"
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
         TabIndex        =   108
         Top             =   195
         Width           =   135
      End
   End
   Begin VB.Frame FrameMed 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2435
      Left            =   0
      TabIndex        =   112
      Top             =   4320
      Width           =   3495
      Begin VB.TextBox txtMedDay1 
         Alignment       =   2  'Center
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
         Index           =   6
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   306
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtMedDay1 
         Alignment       =   2  'Center
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
         Index           =   5
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   305
         Top             =   1730
         Width           =   495
      End
      Begin VB.TextBox txtrate 
         Height          =   285
         Index           =   6
         Left            =   2880
         TabIndex        =   304
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtrate 
         Height          =   285
         Index           =   5
         Left            =   2880
         TabIndex        =   303
         Top             =   1730
         Width           =   495
      End
      Begin VB.TextBox txtrate 
         Height          =   285
         Index           =   4
         Left            =   2880
         TabIndex        =   302
         Top             =   1460
         Width           =   495
      End
      Begin VB.TextBox txtrate 
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   301
         Top             =   1180
         Width           =   495
      End
      Begin VB.TextBox txtrate 
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   300
         Top             =   900
         Width           =   495
      End
      Begin VB.TextBox txtrate 
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   299
         Top             =   620
         Width           =   495
      End
      Begin VB.TextBox txtrate 
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   298
         Top             =   360
         Width           =   495
      End
      Begin VB.ComboBox cboMedType1 
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   297
         Top             =   2000
         Width           =   1650
      End
      Begin VB.ComboBox cboMedType1 
         Height          =   315
         Index           =   5
         Left            =   240
         TabIndex        =   296
         Top             =   1730
         Width           =   1650
      End
      Begin VB.ComboBox cboMedType1 
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   295
         Top             =   1460
         Width           =   1650
      End
      Begin VB.ComboBox cboMedType1 
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   294
         Top             =   1180
         Width           =   1650
      End
      Begin VB.ComboBox cboMedType1 
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   293
         Top             =   900
         Width           =   1650
      End
      Begin VB.ComboBox cboMedType1 
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   292
         Top             =   620
         Width           =   1650
      End
      Begin VB.TextBox txtMedDay1 
         Alignment       =   2  'Center
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
         Index           =   0
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   29
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtMedDay1 
         Alignment       =   2  'Center
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
         Index           =   1
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   31
         Top             =   620
         Width           =   495
      End
      Begin VB.TextBox txtMedDay1 
         Alignment       =   2  'Center
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
         Index           =   2
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   33
         Top             =   900
         Width           =   495
      End
      Begin VB.TextBox txtMedDay1 
         Alignment       =   2  'Center
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
         Index           =   3
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   35
         Top             =   1180
         Width           =   495
      End
      Begin VB.TextBox txtMedDay1 
         Alignment       =   2  'Center
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
         Index           =   4
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   37
         Top             =   1460
         Width           =   495
      End
      Begin VB.TextBox txtMedQty1 
         Alignment       =   2  'Center
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
         Index           =   0
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   30
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtMedQty1 
         Alignment       =   2  'Center
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
         Index           =   1
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   32
         Top             =   620
         Width           =   255
      End
      Begin VB.TextBox txtMedQty1 
         Alignment       =   2  'Center
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
         Index           =   2
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   34
         Top             =   900
         Width           =   255
      End
      Begin VB.TextBox txtMedQty1 
         Alignment       =   2  'Center
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
         Index           =   3
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   36
         Top             =   1180
         Width           =   255
      End
      Begin VB.TextBox txtMedQty1 
         Alignment       =   2  'Center
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
         Index           =   4
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   38
         Top             =   1460
         Width           =   255
      End
      Begin VB.TextBox txtMedQty1 
         Alignment       =   2  'Center
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
         Index           =   5
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   39
         Top             =   1730
         Width           =   255
      End
      Begin VB.TextBox txtMedQty1 
         Alignment       =   2  'Center
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
         Index           =   6
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   40
         Top             =   2040
         Width           =   255
      End
      Begin VB.ComboBox cboMedType1 
         Height          =   315
         Index           =   0
         Left            =   260
         TabIndex        =   28
         Top             =   340
         Width           =   1650
      End
      Begin VB.Label Label3 
         Caption         =   "Price"
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
         Left            =   2880
         TabIndex        =   307
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label101 
         Caption         =   "Medication && Injection"
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
         Index           =   10
         Left            =   360
         TabIndex        =   224
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label43 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   159
         Top             =   2140
         Width           =   135
      End
      Begin VB.Label Label42 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   158
         Top             =   1830
         Width           =   135
      End
      Begin VB.Label Label50 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   1545
         Width           =   975
      End
      Begin VB.Label Label41 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   116
         Top             =   405
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   705
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   113
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label72 
         Caption         =   "Qty"
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
         Left            =   2520
         TabIndex        =   119
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label71 
         Caption         =   " Days"
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
         Left            =   2040
         TabIndex        =   118
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame Frame7 
      Height          =   675
      Left            =   0
      TabIndex        =   145
      Top             =   7080
      Width           =   11895
      Begin VB.CommandButton cmdCompute 
         Caption         =   "&Compute"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6600
         TabIndex        =   91
         Top             =   160
         Width           =   855
      End
      Begin VB.TextBox txtAllBillAmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   161
         Top             =   165
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAllNoOfBill 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   880
         MaxLength       =   20
         TabIndex        =   160
         Top             =   165
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdAddRec 
         Caption         =   "&Verify"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7440
         TabIndex        =   79
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11160
         TabIndex        =   97
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton CmdClearInv 
         Caption         =   "Clear &Bill"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10200
         TabIndex        =   94
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton CmdMBMInfo 
         Caption         =   "&MBM Info"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9240
         TabIndex        =   93
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8640
         TabIndex        =   92
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         TabIndex        =   81
         Top             =   160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5400
         TabIndex        =   95
         Top             =   160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10080
         TabIndex        =   90
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8160
         TabIndex        =   82
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8040
         TabIndex        =   80
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton CmdTotInv 
         Caption         =   "&Tot Inv."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8760
         TabIndex        =   157
         Top             =   160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdVerify 
         Caption         =   "&Verify"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7560
         TabIndex        =   83
         Top             =   160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdClearAll 
         Caption         =   "&Clear All"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8880
         TabIndex        =   96
         Top             =   160
         Width           =   855
      End
      Begin VB.Label lblCAS 
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
         Left            =   4200
         TabIndex        =   246
         Top             =   195
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblCaseNumber 
         Caption         =   "Case No :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   245
         Top             =   195
         Width           =   855
      End
      Begin VB.Label lblAllBillAmt 
         Caption         =   "Tot. Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1560
         TabIndex        =   163
         Top             =   210
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label LblAllNoOfBill 
         Caption         =   "No.of.Bill"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   162
         Top             =   210
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Label Label38 
      Caption         =   "LG No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   8640
      TabIndex        =   276
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblInvoice 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   " Invoice Batch Processing - Clinic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Left            =   0
      TabIndex        =   100
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bDataStatus As Boolean
Dim bExit As Boolean
'Dim IntIndex As String
Dim EditFlag As Boolean
Dim ScanCode As String
Dim AA, CC As Integer
Dim strPayTo, strPayee As String
Dim strMBMno As String
Dim ASOInd As String
Dim UtilisedAmt As Double
Dim IntInvAmt As Double
Dim CASENumber, NewCaseNo As String
Dim MembershipNo As String
Dim CovPersonID As String
Dim strPolNo As String
Dim strPatientName As String
Dim strMBMDataStatus As String
Dim Strrelationtype As String
Dim strMBMCancelDate  As String
Dim strCMBMCancelDate As String
Dim strCMBMCCLNExpDate As String
Dim strPOlStatus As String
Dim strEffdate As String
Dim strExpDate As String
Dim strClnPln As String
Dim strClnIns As String
Dim strPatientDOB As String
Dim strPayor As String
Dim strPatientID As String
Dim strPatientIC As String
Dim strBalLimit As String
Dim strVisitLimit As String
Dim strDentLmt As String
Dim strSpecLmt As String
Dim strMEPLmt As String
Dim strCLNContract As String
Dim strMBMCLNLTM As String
Dim strHLTCode As String
Dim strVisitIND, strAnnualInd As String
Dim strCoPay As String
Dim strSpecStatus As String
Dim strMEPStatus As String
Dim StrDentLmtStatus As String
Dim StrSpecLmtStatus As String
Dim strBNFStatus As String
Dim strDentExtLmt As String
Dim str1stSpecLmt As String
Dim strAccumAmt As String
Dim strGRPMCOStatus As String
Dim strPLNGrossStatus As String
Dim strMBMDepartment As String
Dim strMCStatus As String
Dim strMBMMCDaysBal As String
Dim strGRPPostageFees As String
Dim MBM As New ADODB.Recordset
Dim lstErrRec As ListItem
Dim VPlan, VExpDate, VEffDate, VDiagCode, VDiagIND As String
Dim SuppLmtType As String
Dim Z1, Z2, Z3, Z4, Z5, Z6, Z7, Z8 As Double
Dim OPBatchNo As String
Dim ReasonOtherIND As String
Dim StrDentStatus As String
Private m_blnDirection(1 To 23) As Boolean
Public lstRec As ListItem
Dim tmpCaseNo As String
Dim tmpCLMGPFeeBasicMed As Double
Dim tmpCLMGPFeeLTMed As Double
Dim tmpCLMGPFeeMedInj As Double
Dim tmpCLMSPFeeBasicMed As Double
Dim tmpCLMSPFeeLTMed As Double
Dim tmpCLMSPFeeMedInj As Double
Dim tmpCLMDental As Double
Dim tmpCLMMaternity As Double
Dim tmpCLMPRVCheckup As Double
Dim tmpCLMXRay As Double
Dim tmpCLMUltraSound As Double
Dim tmpCLMECG As Double
Dim tmpCLMDressing As Double
Dim tmpCLMNebulizer As Double
Dim tmpCLMID As Double
Dim tmpCLMTSSTO As Double
Dim tmpCLMMinorSurgery As Double
Dim tmpCLMPOthers As Double
Dim tmpCLMBloodDrawing As Double
Dim tmpCLMBloodGlucose As Double
Dim tmpCLMUrine As Double
Dim tmpCLMPapSmear As Double
Dim tmpCLMLOthers As Double
Dim tmpCLMDTPPolio As Double
Dim tmpCLMDTPPolioHIB As Double
Dim tmpCLMHepatitisB As Double
Dim tmpCLMMMR As Double
Dim tmpCLMIOthers As Double
Dim tmpCLMPOthersDesc As String
Dim tmpCLMLOthersDesc As String
Dim tmpCLMIOthersDesc As String
Dim tmpCLMVision As Double
Dim tmpCLMTyphiad As Double
Dim tmpCLMOtherNoBalDed As Double
Dim tmpInvINdex As String
Dim tmpCASnumber As String
Dim tmpMBMNumber As String
Dim tmpCoverID As String
Dim tmpCASStatus As String
Dim tmpClaimability As String
Dim tmpCaseType As String
Dim TmpPending As String
Dim tmpRegisterDate As String
Dim tmpREgisterBy As String
Dim tmpLastUpdateDate As String
Dim tmpLastUpdateBy As String
Dim tmpClinicCode As String
Dim tmpInvoiceNo As String
Dim tmpInvoiceDate As String
Dim tmpInvoiceStmtAmt As String
Dim tmpDateVisit As String
Dim tmpMCDays As String
Dim tmpOPBatchNo As String
Dim tmpTimeVisit As String
Dim tmpMedInjType1 As String
Dim tmpMedInjType2 As String
Dim tmpMedInjType3 As String
Dim tmpMedInjType4 As String
Dim tmpMedInjType5 As String
Dim tmpMedInjDays1 As String
Dim tmpMedInjDays2 As String
Dim tmpMedInjDays3 As String
Dim tmpMedInjDays4 As String
Dim tmpMedInjDays5 As String
Dim tmpMedInjQty1 As String
Dim tmpMedInjQty2 As String
Dim tmpMedInjQty3 As String
Dim tmpMedInjQty4 As String
Dim tmpMedInjQty5 As String
Dim tmpMedInjType6 As String
Dim tmpMedInjType7 As String
Dim tmpMedInjDays6 As String
Dim tmpMedInjDays7 As String
Dim tmpMedInjQty6 As String
Dim tmpMedInjQty7 As String
Dim tmpInvoiceReceiptDate As String
Dim tmpInvoiceReceiptMode As String
Dim tmpCasRemark As String
Dim tmpCASTreatReason1 As String
Dim tmpCASTreatReason2 As String
Dim tmpCASTreatReason3 As String
Dim tmpCASTreatReason4 As String
Dim tmpCASTreatReason5 As String
Dim tmpCASDiagnosis5 As String
Dim tmpMCStatus As String
Dim tmpNoOfDiag As String
Dim tmpNoOfMedication As String
Dim tmpINVEnteredDate As String
Dim tmpINVEnteredBy As String
Dim tmpINVLastUpdateBy As String
Dim tmpINVLastUpdateDate As String
Dim tmpBillNo As String
Dim tmpBillDate As String
Dim tmpINVAmount As String
Dim tmpINVPayTo As String
Dim tmpINVPayee As String
Dim tmpINVBillAmt As String
Dim tmpGPFee As String
Dim tmpGPLTMed As String
Dim tmpGPInj As String
Dim tmpGPMed As String
Dim tmpSPFee As String
Dim tmpSPLTMed As String
Dim tmpSPInj As String
Dim tmpSPMed As String
Dim tmpClaimOpenDate As String
Dim tmpClaimOpenBy As String
Dim tmpClaimCloseDate As String
Dim tmpClaimClosedBy As String
Dim tmpClaimLastUpdateDate As String
Dim tmpClaimLastUpdateUser As String
Dim tmpCLMCoPayColl As String
Dim tmpCLMPayableByMedix2C As String
Dim tmpCLMCLNExcess As Double
Dim tmpCLMTotalActual As Double
Dim tmpCLMFinalApprove As Double
Dim tmpCLMExcess As Double
Dim tmpCLMBalanceLimit As Double
Dim tmpCLMPreviousBalance As Double
Dim tmpCLMPaidBYMedix2M As Double
Dim tmpCLMPaidBYMedix2C As Double
Dim tmpCLMAdjAnnLmt As Double
Dim tmpClmNonGPFee As Double
Dim tmpNonGPFee As Double
Dim tmpCLMMCOProcFee As Double
Dim tmpCLMMCOProcStatus As String
Dim tmpCLMProcMth As String
Dim tmpCLMProcYR As String
Dim tmpCLMCoPayCLNColl As Double
Dim tmpCLMBordxAmt As Double
Dim tmpCLMNo As String
Dim strGRPBordxFees As String
Dim CASENumberMonth As String
Dim MontlyStatus As String
Dim FinalAmountForMonthLimit As String
Dim tmpcboDrugList As Integer
Dim tmpcboDrugcnt As String
Dim tmpcboDrugLen As String
Dim tmpSaveStatus As Boolean
Dim strPLNMMPerc As String
Dim tmpLimitExceed As String
Dim tmpvisit As Double
Dim TmpSPVisitLimit As Double
Dim TmpMEVisitLimit As Double
Dim VisitInd As String
Dim Tmpprepostmatellimit As Double
Dim tmpLmt As String
Dim PDentalStatus As String
Dim tmpPDentalAnnLimit As String
Dim GPLimit As Double
Dim SPLimit As Double
Dim GPCoPay As Double
Dim SPCoPay As Double
Dim tmpTopupAmt, SPtmpTopupAmt As Double
Dim tmpGrpCompany As String
Dim tmpChildDepLimit As Double
Dim tmpChildLimitInd As String

Private Sub A1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A29_Change()
Dim a As Double
Dim B As Double
Dim LimitValue As Double
Dim copay As Double
Dim DispVal As Double
copay = 0
LimitValue = 0

If tmpGrpCompany = "C014" Then   ' Group company for Prince Court
    If Len(Trim(C1)) Then
        LimitValue = tmpvisit
    Else
        LimitValue = TmpSPVisitLimit
    End If
End If


'Call GetClinicLimit
'If Len(Trim(C1)) Then
'    LimitValue = GPLimit
'    CoPay = GPCoPay
'Else
'    LimitValue = SPLimit
'    CoPay = SPCoPay
'End If

If LimitValue > 0 Then
    If IsNumeric(lblDispCLNBillAmt) > LimitValue Then
        a = IIf(IsNumeric(lblDispCLNBillAmt), lblDispCLNBillAmt, 0)
        B = lblDispCLNBillAmt - LimitValue
        lblActCoPay = B
        
        If Len(A29) Then
            If CDbl(A29) = CDbl(B) Then
            
            Else
                A29 = B
            End If
        Else
            A29 = B
        End If
        
        If CDbl(a) <> "0" Then
            lblCLNBillAmt = CDbl(a) - CDbl(B)
        End If
    Else
        a = IIf(IsNumeric(lblDispCLNBillAmt), lblDispCLNBillAmt, 0)
        B = IIf(IsNumeric(A29), A29, 0)
        If CDbl(a) <> "0" Then
            lblCLNBillAmt = CDbl(a) - CDbl(B)
        End If
    End If
Else
    a = IIf(IsNumeric(lblDispCLNBillAmt), lblDispCLNBillAmt, 0)
    B = IIf(IsNumeric(A29), A29, 0)
    If CDbl(a) <> "0" Then
        lblCLNBillAmt = CDbl(a) - CDbl(B)
    End If
End If
End Sub

Private Sub A32_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A32_LostFocus()
    Call DisplayBillAmt
End Sub


Private Sub A33_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A33_LostFocus()
    Call DisplayBillAmt
End Sub



Private Sub cmdLGSearch_Click()
Dim cmd As New ADODB.Command
Dim rst2 As New ADODB.Recordset
Dim CLN As New ADODB.Recordset
Dim prm As ADODB.Parameter
Dim strAppend As String
Dim strsqlCLN As String
Dim Record As ListItem
Dim tmpLGNo As String

Dim tmpstring As String
If txtLGNo <> "" Then
    
    strAppend = "  LGNo = '" & txtLGNo & "'"
    
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "LGCaseReg"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    Do Until rst2.EOF
        tmpLGNo = txtLGNo
        txtMBMNo = rst2!MBMNumber & "-" & rst2!coverid
        cmdGo.Default = True
        CmdGo_Click
        
        txtLGNo = tmpLGNo
        txtVisitDate = rst2!LGdate
       
        
        If Len(rst2!DoctorName) Then
            TxtRemark = rst2!DoctorName & "/" & rst2!Diagnosis
        End If
        
       
        TxtClinicCode = Trim(rst2!CLNName)
       

        'If Len(rst2!LGRemarks) Then
        '    Record.SubItems(6) = rst2!LGRemarks
        'End If
        
        
        rst2.MoveNext
    Loop
    rst2.Close
    Set rst2 = Nothing
End If
End Sub

Private Sub cmdRegMBM_Click()
    Load frmMemberRegistrationCLN
    frmMemberRegistrationCLN.Show vbModal
End Sub

Private Sub cmdUpdateDept_Click()
Dim StrSql As String
Dim MBM As New ADODB.Recordset
Dim tmpMBMNo As String
Dim tmpZZ As String

    If txtMBMNo <> "" And txtDept <> "" Then
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            tmpZZ = InStr(1, txtMBMNo, "-")
            tmpMBMNo = Trim(Left(txtMBMNo, IIf(tmpZZ - 1 <= 0, Len(Trim(txtMBMNo)), tmpZZ - 1)))
            MBM.Open "SELECT MBMDepartment from MBMOthers where MBMNumber='" & Trim(tmpMBMNo) & "'", medixmasterconn, adOpenStatic, adLockOptimistic
                MBM!MBMDepartment = txtDept
            MBM.update
            MBM.Close
            Set MBM = Nothing
     '   medixmasterconn.Close
     '   Set medixmasterconn = Nothing
            MsgBox "Record Updated! ", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub Command1_Click()

End Sub



Private Sub lblActCoPay_Change()
Dim a As Double
Dim B As Double

    a = IIf(IsNumeric(lblDispBillAmt), lblDispBillAmt, 0)
    B = IIf(IsNumeric(lblActCoPay), lblActCoPay, 0)
    If CDbl(a) <> "0" Then
        lblBillAmt = Format(CDbl(a) - CDbl(B), "#,##0.00")
    End If
End Sub

Private Sub A31_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A31_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A7_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A8_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A9_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A10_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A11_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A12_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A13_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A14_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A15_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A16_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A17_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A18_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A19_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A20_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A21_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A22_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A23_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A24_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A25_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A26_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A27_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A28_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A30_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A30_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B1_LostFocus()
    Z1 = IIf(IsNumeric(B1), B1, 0)
    Z2 = IIf(IsNumeric(B2), B2, 0)
    Z3 = IIf(IsNumeric(B3), B3, 0)
    Z4 = IIf(IsNumeric(B4), B4, 0)
    C1 = CDbl(Z1) + CDbl(Z2) + CDbl(Z3) + CDbl(Z4)
    Call DisplayBillAmt
End Sub

Private Sub B2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B2_LostFocus()
    Z1 = IIf(IsNumeric(B1), B1, 0)
    Z2 = IIf(IsNumeric(B2), B2, 0)
    Z3 = IIf(IsNumeric(B3), B3, 0)
    Z4 = IIf(IsNumeric(B4), B4, 0)
    C1 = CDbl(Z1) + CDbl(Z2) + CDbl(Z3) + CDbl(Z4)
    Call DisplayBillAmt
End Sub

Private Sub B3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B3_LostFocus()
    Z1 = IIf(IsNumeric(B1), B1, 0)
    Z2 = IIf(IsNumeric(B2), B2, 0)
    Z3 = IIf(IsNumeric(B3), B3, 0)
    Z4 = IIf(IsNumeric(B4), B4, 0)
    C1 = CDbl(Z1) + CDbl(Z2) + CDbl(Z3) + CDbl(Z4)
    Call DisplayBillAmt
End Sub

Private Sub B4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B4_LostFocus()
    Z1 = IIf(IsNumeric(B1), B1, 0)
    Z2 = IIf(IsNumeric(B2), B2, 0)
    Z3 = IIf(IsNumeric(B3), B3, 0)
    Z4 = IIf(IsNumeric(B4), B4, 0)
    C1 = CDbl(Z1) + CDbl(Z2) + CDbl(Z3) + CDbl(Z4)
    Call DisplayBillAmt
End Sub

Private Sub B5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B5_LostFocus()
   ' Call A29_Change
    Z5 = IIf(IsNumeric(B5), B5, 0)
    Z6 = IIf(IsNumeric(B6), B6, 0)
    Z7 = IIf(IsNumeric(B7), B7, 0)
    Z8 = IIf(IsNumeric(B8), B8, 0)
    C2 = CDbl(Z5) + CDbl(Z6) + CDbl(Z7) + CDbl(Z8)
    
    Call DisplayBillAmt
End Sub

Private Sub B6_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B6_LostFocus()
   ' Call A29_Change
    Z5 = IIf(IsNumeric(B5), B5, 0)
    Z6 = IIf(IsNumeric(B6), B6, 0)
    Z7 = IIf(IsNumeric(B7), B7, 0)
    Z8 = IIf(IsNumeric(B8), B8, 0)
    C2 = CDbl(Z5) + CDbl(Z6) + CDbl(Z7) + CDbl(Z8)
    
    Call DisplayBillAmt
End Sub

Private Sub B7_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B7_LostFocus()
    Z5 = IIf(IsNumeric(B5), B5, 0)
    Z6 = IIf(IsNumeric(B6), B6, 0)
    Z7 = IIf(IsNumeric(B7), B7, 0)
    Z8 = IIf(IsNumeric(B8), B8, 0)
    C2 = CDbl(Z5) + CDbl(Z6) + CDbl(Z7) + CDbl(Z8)
    Call DisplayBillAmt
    Call A29_Change
End Sub

Private Sub B8_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B8_LostFocus()
    Z5 = IIf(IsNumeric(B5), B5, 0)
    Z6 = IIf(IsNumeric(B6), B6, 0)
    Z7 = IIf(IsNumeric(B7), B7, 0)
    Z8 = IIf(IsNumeric(B8), B8, 0)
    C2 = CDbl(Z5) + CDbl(Z6) + CDbl(Z7) + CDbl(Z8)
    Call DisplayBillAmt
    Call A29_Change
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
Private Sub A19_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A21_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A22_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A23_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A24_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A25_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A26_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A27_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A28_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A29_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub CboInvRecMode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboInvRecMode.Text, CboInvRecMode.SelStart) + Chr(KeyAscii) + Mid(CboInvRecMode.Text, CboInvRecMode.SelStart + CboInvRecMode.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboInvRecMode.ListCount - 1
       
        If NewText = LCase(Left(CboInvRecMode.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboInvRecMode.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboInvRecMode.Text = ValidValue
            CboInvRecMode.SelStart = 0
            CboInvRecMode.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CmdAddRec_Click()
Dim isVisit As Boolean
isVisit = False
tmpLmt = ""

Call A29_Change

'1. Check visit status if Visit status =y then check case type (GP,SP,ME)
    If VisitInd = "Y" Or PDentalStatus = "Y" Then
       If txtDiagnosis1 = "211" Or txtDiagnosis1 = "210" Or txtDiagnosis2 = "211" Or txtDiagnosis2 = "210" Or txtDiagnosis3 = "211" Or txtDiagnosis3 = "210" Or txtDiagnosis4 = "211" Or txtDiagnosis4 = "210" Or TxtDiagnosis5 = "211" Or TxtDiagnosis5 = "210" Then
            tmpLmt = Tmpprepostmatellimit
            If lblDispBillAmt <= Tmpprepostmatellimit Then
            isVisit = True
            Else
            Tmpprepostmatellimit = 0
            isVisit = True
            End If
       ElseIf Len(A8) Then
       tmpLmt = TmpMEVisitLimit
        If lblDispBillAmt <= TmpMEVisitLimit Then
            isVisit = True
            Else
            TmpMEVisitLimit = 0
            isVisit = True
        End If
        ElseIf txtDiagnosis1 = "303" Or txtDiagnosis2 = "303" Or txtDiagnosis3 = "303" Or txtDiagnosis4 = "303" Or TxtDiagnosis5 = "303" Then
        tmpLmt = tmpPDentalAnnLimit
        If lblDispBillAmt <= tmpPDentalAnnLimit Then
          isVisit = True
          Else
          tmpPDentalAnnLimit = 0
          isVisit = True
        End If
  
      ElseIf Mid(lblClnType, 1, 1) = "G" Then
      tmpLmt = tmpvisit
        If lblDispBillAmt <= tmpvisit Then
            isVisit = True
            Else
            tmpvisit = 0
            isVisit = True
        End If
      ElseIf Mid(lblClnType, 1, 1) = "S" Then
      tmpLmt = TmpSPVisitLimit
        If lblDispBillAmt <= TmpSPVisitLimit Then
            isVisit = True
            Else
            TmpSPVisitLimit = 0
            isVisit = True
        End If
     End If
    Else
        isVisit = True
    End If
    If isVisit Then
    tmpSaveStatus = False
    frmWorksheet.InvoiceClick = 1
    lstInvoice.ListItems.Clear
    Call ChkData
    Else
        MsgBox ("Please check claim amount more then visit limit " + tmpLmt)
    End If
End Sub

Private Sub CmdClearInv_Click()
    Call ClearEntries
    Call ClearMBMInfo
    FrameMBMCLN.Enabled = True
    EditFlag = False
    Call EntriesEnabled(False)
    txtMBMNo.SetFocus
End Sub

Private Sub CmdClearAll_Click()
    Call ClearEntries
    Call ClearMBMInfo
    txtInvDate = "__/__/____"
    txtInvNo = ""
    txtClinicName = ""
    TxtClinicCode = ""
    txtAllBillAmt = 0
    txtAllNoOfBill = 0
    TxtTotInvoiceAmt = ""
    txtInvRecDate = "__/__/____"
    CboInvRecMode.Text = ""
    CmdMBMInfo.Enabled = False
    Call EntriesEnabled(False)
    lstInvoice.ListItems.Clear
    lstErr.ListItems.Clear
    EditFlag = False
    FrameMBMCLN.Enabled = True
    lblClnType = ""
    TxtClinicCode.SetFocus
End Sub

Private Sub cmdCompute_Click()
        frmWorksheet.InvoiceClick = 1
        frmWorksheet.INVCLNForm = 1
        frmWorksheet.Show
End Sub

Private Sub cmdEdit_Click()
    If Trim(txtMBMNo) <> "" Then
        CmdEdit.Visible = False
        CmdEdit.Enabled = False
        CmdAddRec.Enabled = False
        CmdUpdate.Visible = True
        CmdUpdate.Enabled = True
        Call EntriesEnabled(True)
        FrameMBMCLN.Enabled = False
        A7.Enabled = False    '601 Dental
        A30.Enabled = False   '604 Vision
        'If Trim(txtDentStatus) = "C" Then ' only for Combine Lmt (Dental)
            A7.Enabled = True    '601 Dental
            A30.Enabled = True   '604 Vision
        'End If
        EditFlag = True
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CallLimit(tmpMBMTYpe As String, tmpAnnLmt As String, tmpVisitLmt As String, tmpSpecLmt As String, tmpLimitStatus As String)

    tmpAnnLmt = ""
    tmpAnnLmt = Format(GetAnnualLimit(Mid(txtMBMNo, 3, 1), strClnPln, "N", txtEffDate, tmpMBMTYpe, "DC", tmpLimitStatus), "#,##0.00")
    
    'tmpVisitLmt = ""
    'tmpVisitLmt = Format(GetVisitLimit("", strClnPln, "N", txtEffDate, "", tmpLimitStatus, "DC"), "#,##0.00")
   
    tmpSpecLmt = ""
    tmpSpecLmt = Format(GetSpecLimit(Mid(txtMBMNo, 3, 1), strClnPln, "N", txtEffDate, tmpMBMTYpe, "DC"), "#,##0.00")
    
End Sub

Public Sub CmdGo_Click()
On Error Resume Next
Dim ZZ As String
Dim QQ As Integer
Dim tmpMBMTYpe As String
Dim PictureFileName As String
Dim PatientIC As String
Dim MBMSql As String
Dim MBMRec As New ADODB.Recordset
Dim TotalSpecLimit As Double
Dim tmpRecCnt As Integer
Dim MBMNo As String
Dim CovID As String
Dim sArr() As String
Dim dblMeternity As Double
Dim dblDental As Double
Dim OtherDed As Double
Dim thyroid As Double
Dim SpecLmtRec As New ADODB.Recordset
Dim tmpLimitStatus As String
Dim tmpAnnLmt As String
Dim tmpVisitLmt As String
Dim tmpSpecLmt As String
Dim tmpSpecAnnLmt As String
Dim strAppend As String
Dim tmpMaternityInd As String
Dim tmpSpecLimitStatus As String
Dim tmpSpecAccLmt As String
Dim tmpDentalInd As String

    lblCaseNumber.Visible = False
    lblCAS.Visible = False
    lblCAS = ""
    CmdSave.Enabled = False
    TotalSpecLimit = ""
    If lstInvoice.ListItems.Count > 0 Then
        For QQ = 1 To lstInvoice.ListItems.Count
            If Trim(lstInvoice.ListItems(QQ).SubItems(1)) <> "" Then
                lstInvoice.ListItems.Clear
                Exit For
            End If
        Next QQ
    End If
    txtMBMNo = ChkStr(txtMBMNo)
    txtPatientIC = ChkStr(txtPatientIC)
    If Trim(txtPatientIC) = "" And Trim(txtMBMNo) = "" Then
         MsgBox "Please Key Member No Or IC Number ! ", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    If Trim(txtMBMNo) = "" Then
        If Trim(txtPatientIC) <> "" Then
            frmSearchMBMIC.Source = 1
            frmSearchMBMIC.Show
        End If
    End If
    If Trim(txtMBMNo) <> "" Then
        ZZ = InStr(1, txtMBMNo, "-")
        MembershipNo = Trim(Left(txtMBMNo, IIf(ZZ - 1 <= 0, Len(Trim(txtMBMNo)), ZZ - 1)))
        If ZZ = 0 Then
            CovPersonID = "00"
        Else
            CovPersonID = Trim(Right(txtMBMNo, IIf(ZZ - 1 <= 0, 0, Len(Trim(txtMBMNo)) - ZZ)))
        End If
        tmpRecCnt = 1
        If Len(MembershipNo) = 8 Or Len(MembershipNo) = 9 Then
            frmSearchMBM05.Source = 1
            frmSearchMBM05.Show
            ZZ = InStr(1, txtMBMNo, "-")
            MembershipNo = Trim(Left(txtMBMNo, IIf(ZZ - 1 <= 0, Len(Trim(txtMBMNo)), ZZ - 1)))
            If ZZ = 0 Then
                CovPersonID = "00"
            Else
                CovPersonID = Trim(Right(txtMBMNo, IIf(ZZ - 1 <= 0, 0, Len(Trim(txtMBMNo)) - ZZ)))
            End If
            tmpRecCnt = frmSearchMBM05.SearchRecCnt
        End If
  
        If tmpRecCnt < 2 Then
            Call MemberSuppInfo(MembershipNo, CovPersonID, strBalLimit, strVisitLimit, strAccumAmt, _
                                StrDentStatus, strCoPay, strPayor, strHLTCode, strClnPln, _
                                strClnIns, strEffdate, strExpDate, strPolNo, strMBMDataStatus, _
                                strPOlStatus, strPatientName, strPatientDOB, strPatientIC, _
                                strMBMCLNLTM, strCLNContract, strVisitIND, strAnnualInd, strSpecStatus, _
                                strMEPStatus, strDentLmt, strSpecLmt, strMEPLmt, SuppLmtType, _
                                StrDentLmtStatus, StrSpecLmtStatus, strBNFStatus, strDentExtLmt, _
                                str1stSpecLmt, strGRPMCOStatus, strPLNGrossStatus, strMBMDepartment, _
                                strMCStatus, strMBMMCDaysBal, strGRPPostageFees, strGRPBordxFees, _
                                strPLNMMPerc, Strrelationtype, strMBMCancelDate, strCMBMCancelDate, strCMBMCCLNExpDate)
        End If
        
        ' Write code to get visit limit and values based on plan code
        Dim visitLmt As New ADODB.Recordset
        Dim VisitLimitquery As String
        
        VisitLimitquery = ""
        VisitLimitquery = "Select AnnualLimit.VisitLimit,prepostmatellimit,PDentalStatus,PDentalAnnLimit,SPVisitLimit,MEVisitLimit,PLN.VISLmtInd,PLN.ChildLimitInd,AnnualLimit.ChildDepLimit from  CISDB..AnnualLimit Inner join CISDB..pln ON pln.PLNCode=AnnualLimit.PLNCode where AnnualLimit.PLNCode='" & Trim(strClnPln) & "'"
        visitLmt.Open VisitLimitquery, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        If visitLmt.EOF = False Then
        tmpvisit = IIf(IsNumeric(visitLmt!VisitLimit), visitLmt!VisitLimit, 0)
        TmpSPVisitLimit = IIf(IsNumeric(visitLmt!SPVisitLimit), visitLmt!SPVisitLimit, 0)
        Tmpprepostmatellimit = IIf(IsNumeric(visitLmt!prepostmatellimit), visitLmt!prepostmatellimit, 0)
        TmpMEVisitLimit = IIf(IsNumeric(visitLmt!MEVisitLimit), visitLmt!MEVisitLimit, 0)
        VisitInd = IIf(Len(visitLmt!VISLmtInd), visitLmt!VISLmtInd, "")
        PDentalStatus = IIf(Len(visitLmt!PDentalStatus), visitLmt!PDentalStatus, "")
        tmpPDentalAnnLimit = IIf(IsNumeric(visitLmt!PDentalAnnLimit), visitLmt!PDentalAnnLimit, 0)
        tmpChildLimitInd = IIf(Len(visitLmt!ChildLimitInd), visitLmt!ChildLimitInd, "N")
        tmpChildDepLimit = IIf(IsNumeric(visitLmt!ChildDepLimit), visitLmt!ChildDepLimit, 0)
        End If
        visitLmt.Close
        Set visitLmt = Nothing
 
        If Trim(MembershipNo) <> "" Then
            EditFlag = False
            CmdAddRec.Enabled = True
            CmdMBMInfo.Enabled = True
            tmpMBMTYpe = "P"   ' P - Principal
            If CovPersonID <> "00" Then
                tmpMBMTYpe = "S" ' S - Supplementary
            End If
            Call ClearEntries
            CmdEdit.Visible = True
            Call EntriesEnabled(True)
            txtTabMBMNo = MembershipNo
            txtTabCoverID = CovPersonID
            TxtPayor = strPayor
            txtPlanCode = strClnPln
            TxtInsType = strClnIns
            txtPatientName = strPatientName
            txtEffDate = strEffdate
            txtExpDate = strExpDate
            TxtPolNo = strPolNo
            txtPolStatus = strPOlStatus
            txtMBMStatus = strMBMDataStatus
            txtdob = strPatientDOB
            txtPatientIC = strPatientIC
            txtPatientName = strPatientName
            txtDept = strMBMDepartment
            txtMCDaysBal = strMBMMCDaysBal
            txtAccumAmt = Format(strAccumAmt, "#,##0.00")
            
            
                ASOInd = ""
                tmpSpecLimitStatus = ""
                SpecLmtRec.Open "SELECT Top 1 PLNMaternityInd,PDentalStatus,ASOInd FROM PLN WHERE PLNCode='" & Trim(strClnPln) & "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                If SpecLmtRec.EOF = False Then
                    tmpMaternityInd = IIf(Len(SpecLmtRec!PLNMaternityInd), SpecLmtRec!PLNMaternityInd, "")
                    tmpDentalInd = IIf(Len(SpecLmtRec!PDentalStatus), SpecLmtRec!PDentalStatus, "")
                    ASOInd = IIf(Len(SpecLmtRec!ASOInd), SpecLmtRec!ASOInd, "N")
                End If
                SpecLmtRec.Close
                Set SpecLmtRec = Nothing
                
            
            If ASOInd = "Y" Then
            
                Dim ASOqry As String
                Dim ASORS As New ADODB.Recordset
                Dim sMBMNumber() As String
                Dim MBMNumber As String
                Dim coverid As String
                Dim ASOlimit As Double
                Dim IPLimit As Double
                Dim OPLimit As Double
                
                Dim IPUtilized As Double
                Dim OPUtilized As Double
                
                Dim IsCombined As String
                If (InStr(1, txtMBMNo, "-")) > 0 Then
                    sMBMNumber = Split(txtMBMNo, "-")
                    MBMNumber = sMBMNumber(0)
                    coverid = sMBMNumber(1)
                Else
                    MBMNumber = txtMBMNo
                    coverid = "00"
                End If
                
                ASOqry = ""
                ASOqry = "Select MBMASOLimit,MBMIPLimit,MBMOPLimit,MBMIPUtilized,MBMOPUtilized ,IsIPOPCombined  from Tbl_ASOMemberDetails  where MBMNumber= '" + MBMNumber + "' and MBMCoverID='" + coverid + "'"
                ASORS.Open ASOqry, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                Do Until ASORS.EOF
                     ASOlimit = IIf(Len(ASORS!MBMASOLimit), ASORS!MBMASOLimit, 0)
                     IPLimit = IIf(Len(ASORS!MBMIPLimit), ASORS!MBMIPLimit, 0)
                     OPLimit = IIf(Len(ASORS!MBMOPLimit), ASORS!MBMOPLimit, 0)
                     IPUtilized = IIf(Len(ASORS!MBMIPUtilized), ASORS!MBMIPUtilized, 0)
                     OPUtilized = IIf(Len(ASORS!MBMOPUtilized), ASORS!MBMOPUtilized, 0)
                     IsCombined = IIf(Len(ASORS!IsIPOPCombined), ASORS!IsIPOPCombined, "N")
                ASORS.MoveNext
                Loop
                ASORS.Close
                Set ASORS = Nothing
                
                If IsCombined = "Y" Then
                    'txtAnnLmt = ASOlimit - (IPUtilized + OPUtilized)
                    txtAnnLmt = ASOlimit
                    UtilisedAmt = IPUtilized + OPUtilized
                Else
                    'txtAnnLmt = OPLimit - OPUtilized
                    txtAnnLmt = OPLimit
                    UtilisedAmt = OPUtilized
                End If
                txtAccumAmt = Format(UtilisedAmt, "#,##0.00")
            Else
                txtAnnLmt = GetAnnualLimit(strClnIns, strClnPln, strHLTCode, CDate(strEffdate), tmpMBMTYpe, strPayor, SuppLmtType)
            End If
            
            
            Dim TopupRec As New ADODB.Recordset
            tmpTopupAmt = 0
            SPtmpTopupAmt = 0
            TopupRec.Open "SELECT sum(coalesce(TopupAmt,0)) as sumOfTopupAmt , Sum(coalesce(SP_TopupAmt,0)) AS SP_TopupAmt from MBMCISTopupHistory where MBMNumber='" & MembershipNo & "' and MBMCovID='" & CovPersonID & "' group by MBMNumber", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            With TopupRec
                If TopupRec.EOF = False Then
                    tmpTopupAmt = IIf(IsNumeric(TopupRec!sumOfTopupAmt), TopupRec!sumOfTopupAmt, 0)
                    SPtmpTopupAmt = IIf(IsNumeric(TopupRec!SP_TopupAmt), TopupRec!SP_TopupAmt, 0)
                End If
            End With
            TopupRec.Close
            Set TopupRec = Nothing
            
            
            TxtVisitLimit = Format(GetVisitLimit(strClnIns, strClnPln, strHLTCode, CDate(strEffdate), tmpMBMTYpe, strPayor), "#,##0.00")
            If IsNumeric(txtAnnLmt) Then
                TxtBalLimit = Format(CDbl(txtAnnLmt) + CDbl(tmpTopupAmt) - CDbl(txtAccumAmt), "#,##0.00")
            End If
            If IsNumeric(TxtVisitLimit) Then
                TxtVisitLimit = Format(strVisitLimit, "#,##0.00")
            End If
            
         'get the SP and GP limit  'Trinath Code
            Dim GetGPSPLmt As New ADODB.Recordset
            Dim GPSPQry As String
            GPSPQry = ""
            If CovPersonID = "00" Then
                GPSPQry = "Select PLN.GRPCompany, PLN.PrinGPLimit as GPLimit,PLN.PrinSPLimit as SPLimit  from HISDB..MBMCLNOthers inner join CISDB..PLN on PLN.PLNCode = MBMCLNOthers.PLNType where HISDB..MBMCLNOthers.MBMNumber ='" + MembershipNo + "'"
            Else
                GPSPQry = "Select PLN.GRPCompany, PLN.DepGPLimit as GPLimit,PLN.DepSPLimit as SPLimit  from HISDB..MBMCoveredPersonsCLN  inner join CISDB..PLN on PLN.PLNCode = MBMCoveredPersonsCLN.MBMCCLNPLNCode    where HISDB..MBMCoveredPersonsCLN.MBMCNumber ='" + MembershipNo + "' and HISDB..MBMCoveredPersonsCLN.MBMCCoverID ='" + CovPersonID + "'"
            End If
            GetGPSPLmt.Open GPSPQry, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            If GetGPSPLmt.EOF = False Then
                txtGPLimit = IIf(IsNumeric(GetGPSPLmt!GPLimit), GetGPSPLmt!GPLimit, 0)
                txtSPLimit = IIf(IsNumeric(GetGPSPLmt!SPLimit), GetGPSPLmt!SPLimit, 0)
                tmpGrpCompany = IIf(Len(GetGPSPLmt!GRPCompany), GetGPSPLmt!GRPCompany, "")
            End If
            GetGPSPLmt.Close
            Set GetGPSPLmt = Nothing
            
           '------Calculate Limt Start-------
            
            tmpLimitStatus = SuppLmtType
            
            txtAccumAmt = 0
            Call CallLimit(tmpMBMTYpe, tmpAnnLmt, tmpVisitLmt, tmpSpecLmt, tmpLimitStatus)
            txtSpecLmt = tmpSpecLmt
            tmpSpecAnnLmt = tmpSpecLmt
            
            tmpSpecLimitStatus = ""
            SpecLmtRec.Open "SELECT SuppLimitStatus FROM SpecialistLimit WHERE PLNCode='" & Trim(strClnPln) & "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            If SpecLmtRec.EOF = False Then
                tmpSpecLimitStatus = IIf(Len(SpecLmtRec!SuppLimitStatus), SpecLmtRec!SuppLimitStatus, "A")
            End If
            SpecLmtRec.Close
            Set SpecLmtRec = Nothing
            
            If Trim(tmpSpecAnnLmt) = "N/A" Then
            MembershipNo = txtMBMNo
             If Len(MembershipNo) Then
                    If Len(MembershipNo) Then
                       If InStr(1, MembershipNo, "-") Then
                          sArr = Split(MembershipNo, "-")
                          MBMNo = sArr(0)
                          CovID = sArr(1)
                       Else
                          MBMNo = MembershipNo
                          CovID = "00"
                       End If
                    End If
                    
                strAppend = ""
                If tmpLimitStatus = "A" Then
                    strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "'"
                ElseIf tmpLimitStatus = "B" Then
                    If CovID = "00" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID = '00' "
                    Else
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID <> '00' "
                    End If
                ElseIf tmpLimitStatus = "C" Or tmpLimitStatus = "" Then
                    strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                ElseIf tmpLimitStatus = "F" Then
                    strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                ElseIf tmpLimitStatus = "E" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                ElseIf tmpLimitStatus = "D" Then
                    If CovID = "00" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID = '00' "
                    Else
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID <> '00' "
                    End If
                End If
             End If
                
               ' ASOInd = ""
               ' tmpSpecLimitStatus = ""
               ' SpecLmtRec.Open "SELECT Top 1 PLNMaternityInd,PDentalStatus,ASOInd FROM PLN WHERE PLNCode='" & Trim(strClnPln) & "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
               ' If SpecLmtRec.EOF = False Then
               '     tmpMaternityInd = IIf(Len(SpecLmtRec!PLNMaternityInd), SpecLmtRec!PLNMaternityInd, "")
               '     tmpDentalInd = IIf(Len(SpecLmtRec!PDentalStatus), SpecLmtRec!PDentalStatus, "")
               '     ASOInd = IIf(Len(SpecLmtRec!ASOInd), SpecLmtRec!ASOInd, "N")
               ' End If
               ' SpecLmtRec.Close
               ' Set SpecLmtRec = Nothing
                
                Dim cmd As New ADODB.Command
                Dim Prm1 As New ADODB.Parameter
                Dim rst2 As New ADODB.Recordset
                
                If Trim(strAppend) = "" Then
                   ' tmpMBMFFS = 0
                    'tmpMBMCLMAmt = 0
                Else
                    Set cmd.ActiveConnection = MediConnCISDB
                    cmd.CommandText = "CIS_FFS2013"
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandTimeout = 300
                    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                    cmd.Parameters.Append Prm1
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd.Execute
                    
                    Do While Not rst2.EOF
                        'tmpMBMFFS = CDbl(tmpMBMFFS) + IIf(IsNumeric(rst2!FFS), rst2!FFS, 0)
                        
                        dblMeternity = 0
                        dblDental = 0
                        OtherDed = 0
                        thyroid = 0
                
                         If IsNumeric(rst2!CLMMaternity) Then
                             dblMeternity = CDbl(rst2!CLMMaternity)
                         Else
                             dblMeternity = 0
                         End If
                           
                         If IsNumeric(rst2!CLMDental) Then
                             dblDental = CDbl(rst2!CLMDental)
                         Else
                             dblDental = 0
                         End If
                          
                        If dblDental > 0 Then
                         dblDental = dblDental
                        End If
                        
                         If IsNumeric(rst2!CLMOthersNoDed) Then
                             OtherDed = CDbl(rst2!CLMOthersNoDed)
                         Else
                             OtherDed = 0
                         End If
                           
                         If IsNumeric(rst2!CLMTyphiadJabNoDed) Then
                             thyroid = CDbl(rst2!CLMTyphiadJabNoDed)
                         Else
                             thyroid = 0
                         End If
                         
                        If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                            txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid + dblDental), 0)
                        ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                            txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid), 0)
                        ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                          txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid + dblDental), 0)
                        ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                          txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid), 0)
                        End If
                        
                        If ASOInd = "Y" Then
                            txtAccumAmt = UtilisedAmt
                        End If
                        rst2.MoveNext
                    Loop
                    
                    rst2.Close
                    Set rst2 = Nothing
                    Set cmd = Nothing
                    Set Prm1 = Nothing
                End If
            Else
                If Len(MembershipNo) Then
                
                    If Len(MembershipNo) Then
                       If InStr(1, MembershipNo, "-") Then
                          sArr = Split(MembershipNo, "-")
                          MBMNo = sArr(0)
                          CovID = sArr(1)
                       Else
                          MBMNo = MembershipNo
                          'CovID = "00"
                          CovID = Trim(CovPersonID)
                       End If
                    End If

                    strAppend = ""
                    If tmpLimitStatus = "A" Then
                        strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(MBMNo) & "'"
                    ElseIf tmpLimitStatus = "B" Then
                        If CovID = "00" Then
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID ='00' "
                        Else
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID <>'00' "
                        End If
                    ElseIf tmpLimitStatus = "C" Or tmpLimitStatus = "" Then
                        strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                    ElseIf tmpLimitStatus = "F" Then
                        strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                    ElseIf tmpLimitStatus = "E" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                    ElseIf tmpLimitStatus = "D" Then
                        If CovID = "00" Then
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID = '00' "
                        Else
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID <> '00' "
                        End If
                    End If
                    
                End If
                
                If Trim(strAppend) = "" Then
                    'tmpMBMFFS = 0
                    'tmpMBMCLMAmt = 0
                Else
                    Set cmd.ActiveConnection = MediConnCISDB
                    cmd.CommandText = "CIS_FFS2013"
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandTimeout = 300
                    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                    cmd.Parameters.Append Prm1
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd.Execute
                    Dim DblValue As Double
                
                    Do While Not rst2.EOF
                    
                    
                     dblMeternity = 0
                dblDental = 0
                OtherDed = 0
                thyroid = 0
                
                If IsNumeric(rst2!CLMMaternity) Then
                    dblMeternity = CDbl(rst2!CLMMaternity)
                Else
                    dblMeternity = 0
                End If
                  
                If IsNumeric(rst2!CLMDental) Then
                    dblDental = CDbl(rst2!CLMDental)
                Else
                    dblDental = 0
                End If
                                
                If IsNumeric(rst2!CLMOthersNoDed) Then
                    OtherDed = CDbl(rst2!CLMOthersNoDed)
                Else
                    OtherDed = 0
                End If
                
                If IsNumeric(rst2!CLMTyphiadJabNoDed) Then
                    thyroid = CDbl(rst2!CLMTyphiadJabNoDed)
                Else
                    thyroid = 0
                End If
                    
                If rst2!PLNGPSPShared = "Y" Then
                    If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                        txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid + dblDental), 0)
                    ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                        txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid), 0)
                    ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                        txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid + dblDental), 0)
                    ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                        txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid), 0)
                    End If
                    If ASOInd = "Y" Then
                       txtAccumAmt = UtilisedAmt
                    End If
                Else
                    If rst2!CLNType <> "S" Then
                            If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                                txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid + dblDental), 0)
                            ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                                txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid), 0)
                            ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                                txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid + dblDental), 0)
                            ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                                txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid), 0)
                            End If
                            If ASOInd = "Y" Then
                               txtAccumAmt = UtilisedAmt
                            End If
                    ElseIf rst2!CLNType = "S" Then
                            If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                                TotalSpecLimit = CDbl(TotalSpecLimit) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid + dblDental), 0)
                            ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                                TotalSpecLimit = CDbl(TotalSpecLimit) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid), 0)
                            ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                                TotalSpecLimit = CDbl(TotalSpecLimit) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid + dblDental), 0)
                            ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                                TotalSpecLimit = CDbl(TotalSpecLimit) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid), 0)
                            End If
                            If ASOInd = "Y" Then
                               txtAccumAmt = UtilisedAmt
                            End If
                    End If
                End If
              rst2.MoveNext
                    Loop
                    rst2.Close
                    Set rst2 = Nothing
                    Set cmd = Nothing
                    Set Prm1 = Nothing
                End If
                
                
                tmpSpecAccLmt = Format(TotalSpecLimit, "#,###.00")
                If IsNumeric(tmpSpecLmt) Then
                    tmpSpecAnnLmt = Format(tmpSpecAnnLmt, "#,##0.00")
                    txtSpecLmt = tmpSpecAnnLmt
                     If IsNumeric(tmpSpecAccLmt) Then
                        txtSpecBal = CDbl(tmpSpecAnnLmt) + CDbl(SPtmpTopupAmt) - CDbl(tmpSpecAccLmt)
                    End If
                End If
                
                
               ' ------Limit calculations ------
            End If
            
            If tmpLimitStatus = "F" And tmpMBMTYpe = "S" And (Strrelationtype = "D" Or Strrelationtype = "S") And tmpChildLimitInd = "Y" Then
                txtAnnLmt = tmpChildDepLimit
            End If
            
            
            If IsNumeric(txtAnnLmt) Then
                TxtBalLimit = Format(CDbl(txtAnnLmt) + CDbl(tmpTopupAmt) - CDbl(txtAccumAmt), "#,##0.00")
            Else
            End If
                        
            '------Calculate Limt End-------
           
            txtLTM = strMBMCLNLTM
            txtDentStatus = StrDentStatus
            A7.Enabled = False    '601 Dental
            A30.Enabled = False   '604 Vision
            A7.Enabled = True    '601 Dental
            A30.Enabled = True   '604 Vision
            A29 = Trim(Format(strCoPay, "#,##0.00"))
            lblActCoPay = Trim(Format(strCoPay, "#,##0.00"))
            
            If Trim(strMCStatus) = "Y" Then
                txtMCDaysBal = strMBMMCDaysBal
                If Trim(strMBMMCDaysBal) = 0 Then
                    MsgBox "MC Days Exceeded! ", vbCritical
                End If
            End If
        End If
        
      
        
        'If strPOlStatus <> "A" Then
         '   MsgBox "Membership has been terminated, no transaction allowed!", vbCritical
          '  CmdAddRec.Enabled = False
        'End If
    End If
End Sub

Private Sub CmdMBMInfo_Click()
    FrmINVMBMinfo.Source = 1
    FrmINVMBMinfo.Show vbModal
End Sub

Private Sub cmdPrint_Click()
    If lstInvoice.ListItems.Count > 0 Then
        Call ExportToExcelError(lstInvoice)
    End If
End Sub

Private Sub CmdRemove_Click()
Dim RecRemove
    
    RecRemove = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecRemove = vbYes Then
        Call RemoveRecFrListV
        Call ClearEntries
        Call ClearMBMInfo
    End If
End Sub

Public Sub cmdSave_Click()
On Err GoTo ErrorSave
Dim startTrans As Boolean
Dim RecSaved
Dim ValidateLimit As Boolean
ValidateLimit = False


    frmWorksheet.INVCLNForm = 1
    If lstInvoice.ListItems.Count <> 0 Then
    
    
    Dim typhoidjab As Double
    Dim nolmtdedothers As Double
    
    typhoidjab = IIf(IsNumeric(Trim(A32.Text)), A32.Text, 0)
    nolmtdedothers = IIf(IsNumeric(Trim(A33.Text)), A33.Text, 0)
    
        If IsNumeric(TxtBalLimit) And lblClnType = "G" Then
            If ((CDbl(TxtBalLimit) - CDbl(txtCLNBillAmt)) + (typhoidjab + nolmtdedothers)) >= 0 Then
                ValidateLimit = True
            Else
                ValidateLimit = False
            End If
        ElseIf IsNumeric(txtSpecBal) And lblClnType = "S" Then
            If ((CDbl(txtSpecBal) - CDbl(txtCLNBillAmt)) + (typhoidjab + nolmtdedothers)) >= 0 Then
                ValidateLimit = True
            Else
                ValidateLimit = False
            End If
        Else
                ValidateLimit = True
        End If
    
            If ValidateLimit Then
                
                If Trim(txtAnnLmt) = Trim("FULL") Then
                        RecSaved = MsgBox("Are you sure ? ", vbYesNo + vbExclamation)
                ElseIf ASOInd = "Y" And txtAnnLmt <= 0 Then
                        RecSaved = MsgBox("ASO Limit already Exceed. Still you want to Proceed... ", vbYesNo + vbExclamation)
                Else
                        RecSaved = MsgBox("Are you sure ? ", vbYesNo + vbExclamation)
                End If
            
                'RecSaved = MsgBox("Are you sure ? ", vbYesNo + vbExclamation)
                If RecSaved = vbYes Then
                    Call A29_Change
                    MediConnCISDB.beginTrans
                    MediConnCISWS.beginTrans
                    medixmasterconn.beginTrans
                    startTrans = True
                    For AA = 1 To lstInvoice.ListItems.Count
                        NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(66))
                        Call SaveCase
                        Call SaveInvoice
                        Call saveActual
                        Call saveActualOthers
                        Call SaveINVClaim
                        Call SaveINVCLNApp '426-454
                        Call SaveINVCLNContract '211-239
                        Call SaveINVActual '168-196
                        Call SaveINVCLNExcess '383-411
                        Call SaveINVMBMApp '254-282
                        Call SaveINVMBMExcess '340 - 368
                        Call SaveINVMBMLmt '297-325
                        Call SaveCharges
                        Call UpdateMBMBalance
                        Call UpdateMBMMCDays
                        Call SaveDup_ori
                        MontlyStatus = ""
                                      
                    Next AA
                            
                    MediConnCISDB.commitTrans
                    MediConnCISWS.commitTrans
                    medixmasterconn.commitTrans
                    startTrans = False
                    
                    If ASOInd = "Y" Then
                        Call UpdateASOLimit(NewCaseNo)
                    End If
                    
                    CmdClearAll.Enabled = True
                    Call EntriesEnabled(False)
                    EditFlag = False
                    CmdSave.Enabled = False
                    lstInvoice.ListItems.Clear
                    Screen.MousePointer = vbDefault
                    MsgBox "Record Saved ! ", vbOKOnly
                Else
                    lstInvoice.ListItems.Clear
                End If
            Else
                MsgBox ("No Balance Limit to Register Case. Please Check!!!")
            End If
    End If
    Exit Sub

ErrorSave:
    Screen.MousePointer = vbDefault
    If startTrans = True Then
        MediConnCISDB.RollbackTrans
        MediConnCISWS.RollbackTrans
        medixmasterconn.RollbackTrans
    End If
    MsgBox Err.Description
End Sub
Public Sub UpdateASOLimit(CasNumber As String)
    Dim SqlQry As String
    Dim UpdateRS As New ADODB.Recordset
    SqlQry = ""
    SqlQry = "Exec  sp_ASOClaimsUtilization '" + CasNumber + "','NULL','OP'"
    UpdateRS.Open SqlQry, medixmasterconn, adOpenKeyset, adLockOptimistic
End Sub
Private Sub SaveDup_ori()
Dim DupSeqNo As String
DupSeqNo = ""
Dim countdupcase As Integer
countdupcase = 0
Dim CISCasRec As New ADODB.Recordset
Dim CISCasSql As String
Dim tmpMBMNum As String
Dim tmpMBMArr() As String
tmpMBMNum = Trim(lstInvoice.ListItems(AA).SubItems(2))
Dim Pos As Integer
Pos = 0
Pos = InStr(tmpMBMNum, "-")
If Pos > 0 Then
    tmpMBMArr() = Split(tmpMBMNum, "-")
    tmpMBMNum = tmpMBMArr(0)
Else
    'tmpMBMArr(0) = (tmpMBMNum)
    tmpMBMNum = tmpMBMNum
End If


tmpMBMNumber = tmpMBMNum
tmpCoverID = lstInvoice.ListItems(AA).SubItems(3)
tmpREgisterBy = Logon
tmpClinicCode = lstInvoice.ListItems(AA).SubItems(12)
tmpDateVisit = Format(lstInvoice.ListItems(AA).SubItems(9), "dd/mm/yyyy")
CISCasSql = ""
CISCasSql = "SELECT CoverID,DateVisit,ClinicCode,InvoiceStmtAmt,MBMNumber FROM ciscas where MBMNumber = '" & Trim(tmpMBMNumber) & _
             "' AND COVERID ='" & tmpCoverID & "' AND DateVisit = '" & Format(tmpDateVisit, "mm/dd/yyyy") & "' " & _
             " AND ClinicCode ='" & Trim(tmpClinicCode) & "'"

CISCasRec.Open CISCasSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
If CISCasRec.EOF = False Then
    countdupcase = CISCasRec.RecordCount
End If

CISCasRec.Close
Set CISCasRec = Nothing

If countdupcase > 1 Then
    CISCasSql = ""
    CISCasSql = "SELECT  DupSeqNo FROM CISDB.dbo.CISCas WHERE MBMnumber = '" & Trim(MembershipNo) & _
                "' AND COVERID ='" & tmpCoverID & "' AND DateVisit = '" & Format(tmpDateVisit, "mm/dd/yyyy") & "' " & _
                " AND ClinicCode ='" & Trim(tmpClinicCode) & "' GROUP BY DupSeqNo order by DupSeqNo desc "
    CISCasRec.Open CISCasSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly

    If CISCasRec.EOF = False Then
        If IsNull(CISCasRec!DupSeqNo) Then
            DupSeqNo = ""
        Else
            DupSeqNo = Trim(CISCasRec!DupSeqNo)
        End If
    End If

    CISCasRec.Close
    Set CISCasRec = Nothing
    
    If Len(Trim(DupSeqNo)) > 0 Then
        CISCasSql = ""
        CISCasSql = "UPDATE [CISDB].[dbo].[CISCas] SET DupSeqNo = '" & DupSeqNo & _
                    "' , DupUpdateUser ='" & tmpREgisterBy & "' , DupUpdateDate = '" & Format(DateTime.Now, "mm/dd/yyyy") & "' " & _
                    " , DupAppInd ='N' WHERE MBMnumber = '" & Trim(MembershipNo) & _
                    "' AND COVERID ='" & tmpCoverID & "' AND DateVisit = '" & Format(tmpDateVisit, "mm/dd/yyyy") & "' " & _
                    " AND ClinicCode ='" & Trim(tmpClinicCode) & "' "
        MediConnCISDB.Execute CISCasSql
    Else
        CISCasSql = ""
        CISCasSql = " SELECT Code, DupNo FROM CISDB.dbo.DupClaimSeq WHERE [Code] = 'DUP'"
        CISCasRec.Open CISCasSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        If CISCasRec.RecordCount <= 0 Then
            With CISCasRec
                .AddNew
                CISCasRec!Code = "DUP"
                CISCasRec!DupNo = "000001"
                .update
                DupSeqNo = CISCasRec!Code + CISCasRec!DupNo
            End With
        Else
            CISCasRec!DupNo = Format(CISCasRec!DupNo + 1, "000000")
            DupSeqNo = CISCasRec!Code + CISCasRec!DupNo
            CISCasRec.update
        End If
               
        CISCasRec.Close
        Set CISCasRec = Nothing
        CISCasSql = ""
        CISCasSql = "UPDATE [CISDB].[dbo].[CISCas] SET DupSeqNo = '" & DupSeqNo & _
                    "' , DupUpdateUser ='" & tmpREgisterBy & "' , DupUpdateDate = '" & Format(DateTime.Now, "mm/dd/yyyy") & "' " & _
                    " , DupAppInd ='N' WHERE MBMnumber = '" & Trim(MembershipNo) & _
                    "' AND COVERID ='" & tmpCoverID & "' AND DateVisit = '" & Format(tmpDateVisit, "mm/dd/yyyy") & "' " & _
                    " AND ClinicCode ='" & Trim(tmpClinicCode) & "' "
        MediConnCISDB.Execute CISCasSql
    End If
End If
End Sub


Private Sub UpdateMBMMCDays()
Dim MBM As New ADODB.Recordset
Dim StrSql As String
Dim TempBalMCDays As Double
Dim TempMCDays As Double
Dim TempFinalMCDays As Double
Dim tmpMBMNum As String
Dim tmpMBMArr() As String
tmpMBMNum = Trim(lstInvoice.ListItems(AA).SubItems(2))
Dim Pos As Integer
Pos = 0
Pos = InStr(tmpMBMNum, "-")
If Pos > 0 Then
    tmpMBMArr() = Split(tmpMBMNum, "-")
    tmpMBMNum = tmpMBMArr(0)
Else
    'tmpMBMArr(0) = (tmpMBMNum)
    tmpMBMNum = tmpMBMNum
End If
'tmpMBMNum = tmpMBMArr(0)

    StrSql = "Select MBMNUmber ,MBMCLNMCDaysBal from MBMCLNOthers " & _
            " where MBMNUmber ='" & tmpMBMNum & "'"
    MBM.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    TempBalMCDays = IIf(IsNumeric(MBM!MBMCLNMCDaysBal), MBM!MBMCLNMCDaysBal, 0)
    TempMCDays = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(11)), lstInvoice.ListItems(AA).SubItems(11), 0)

    If CDbl(TempBalMCDays) > CDbl(TempMCDays) Then
        TempFinalMCDays = CDbl(TempBalMCDays) - CDbl(TempMCDays)
    Else
        TempFinalMCDays = 0
    End If
    MBM!MBMCLNMCDaysBal = TempFinalMCDays
    MBM.update
    MBM.Close
    Set MBM = Nothing
End Sub

Private Sub cmdSearch_Click()
    bDataStatus = True
    frmSearchMBM.Source = 1
    frmSearchMBM.Show vbModal
End Sub

Private Sub cmdSearchICD1_Click()
    frmSearchICD.Source = 1
    frmSearchICD.Show vbModal
End Sub

Private Sub cmdSearchICD2_Click()
    frmSearchICD.Source = 2
    frmSearchICD.Show vbModal
End Sub

Private Sub cmdSearchICD3_Click()
    frmSearchICD.Source = 3
    frmSearchICD.Show vbModal
End Sub

Private Sub cmdSearchICD4_Click()
    frmSearchICD.Source = 4
    frmSearchICD.Show vbModal
End Sub

Private Sub cmdSearchICD5_Click()
Dim TreatRec As New ADODB.Recordset
Dim TreatSql As String

    frmSearchICD.Source = 5
    frmSearchICD.Show vbModal
    
    TxtDiagnosisOth.Enabled = False
    TxtDiagDesc5.Visible = True
    If Trim(TxtDiagnosis5) <> "" Then
       ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

            TreatSql = "Select ReasonOtherDiag from TreatmentReason WHERE ReasonCode='" & Trim(TxtDiagnosis5) & "'"
            TreatRec.Open TreatSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            Do While Not TreatRec.EOF
                ReasonOtherIND = IIf(Len(Trim(TreatRec!ReasonOtherDiag)), Trim(TreatRec!ReasonOtherDiag), "N")
                If ReasonOtherIND = "Y" Then
                    TxtDiagDesc5.Visible = False
                    TxtDiagnosisOth.Enabled = True
                    TxtDiagnosisOth.SetFocus
                Else
                    TxtDiagnosisOth.Enabled = True
                    TxtDiagnosisOth = ""
                End If
            TreatRec.MoveNext
            Loop
            TreatRec.Close
            Set TreatRec = Nothing
       ' MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
    End If
End Sub

Private Sub CmdTotInv_Click()
Dim LstCnt As Integer
    txtAllBillAmt = 0
    txtAllNoOfBill = 0
    If lstInvoice.ListItems.Count > 0 Then
        txtAllNoOfBill.Visible = True
        txtAllBillAmt.Visible = True
        txtAllNoOfBill = lstInvoice.ListItems.Count
        For LstCnt = 1 To lstInvoice.ListItems.Count
            txtAllBillAmt = txtAllBillAmt + CDbl(lstInvoice.ListItems(LstCnt).SubItems(13))
        Next LstCnt
    End If
End Sub

Private Sub CmdUpdate_Click()
Dim isVisit As Boolean
isVisit = False
tmpLmt = ""
'1. Check visit status if Visit status =y then check case type (GP,SP,ME)
 If VisitInd = "Y" Or PDentalStatus = "Y" Then
       If txtDiagnosis1 = "211" Or txtDiagnosis1 = "210" Or txtDiagnosis2 = "211" Or txtDiagnosis2 = "210" Or txtDiagnosis3 = "211" Or txtDiagnosis3 = "210" Or txtDiagnosis4 = "211" Or txtDiagnosis4 = "210" Or TxtDiagnosis5 = "211" Or TxtDiagnosis5 = "210" Then
            tmpLmt = Tmpprepostmatellimit
            If lblDispBillAmt <= Tmpprepostmatellimit Then
            isVisit = True
            Else
            Tmpprepostmatellimit = 0
            isVisit = True
            End If
       ElseIf Len(A8) Then
       tmpLmt = TmpMEVisitLimit
        If lblDispBillAmt <= TmpMEVisitLimit Then
            isVisit = True
        End If
       ElseIf txtDiagnosis1 = "303" Or txtDiagnosis2 = "303" Or txtDiagnosis3 = "303" Or txtDiagnosis4 = "303" Or TxtDiagnosis5 = "303" Then
       tmpLmt = tmpPDentalAnnLimit
       If lblDispBillAmt <= tmpPDentalAnnLimit Then
          isVisit = True
          Else
          tmpPDentalAnnLimit = 0
          isVisit = True
       End If
      ElseIf Mid(lblClnType, 1, 1) = "G" Then
      tmpLmt = tmpvisit
        If lblDispBillAmt <= tmpvisit Then
            isVisit = True
            Else
            tmpvisit = 0
            isVisit = True
        End If
      ElseIf Mid(lblClnType, 1, 1) = "S" Then
      tmpLmt = TmpSPVisitLimit
        If lblDispBillAmt <= TmpSPVisitLimit Then
            isVisit = True
            Else
            TmpSPVisitLimit = 0
            isVisit = True
        End If
     End If
    Else
        isVisit = True
    End If
    If isVisit Then
    frmWorksheet.InvoiceClick = 1
    frmWorksheet.INVCLNForm = 1
    Call ChkData
     Else
        MsgBox ("Please check claim amount more then visist limit " + tmpLmt)
    End If
End Sub

Private Sub cmdVerify_Click()
Dim ChkSql As String
Dim ChKItemRec As New ADODB.Recordset
Dim TotBillAmt As Double
Dim HvSave As Boolean
    HvSave = False
    TotBillAmt = 0
    lstErr.ListItems.Clear
    If lstInvoice.ListItems.Count = 0 Then
        MsgBox "No Record to be Verified ! ", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        
        For CC = 1 To lstInvoice.ListItems.Count
            TotBillAmt = TotBillAmt + lstInvoice.ListItems(CC).SubItems(13)
            VPlan = lstInvoice.ListItems(CC).SubItems(64)
            
            VEffDate = lstInvoice.ListItems(CC).SubItems(78)
            VExpDate = lstInvoice.ListItems(CC).SubItems(79)
            If Len(Trim(lstInvoice.ListItems(CC).SubItems(1))) Then
                HvSave = True
            End If

            If Len(Trim(lstInvoice.ListItems(CC).SubItems(62))) Then
                VDiagCode = Trim(lstInvoice.ListItems(CC).SubItems(62))
                Call VerifyNonPayableDiag
            End If
            If Len(Trim(lstInvoice.ListItems(CC).SubItems(63))) Then
                VDiagCode = Trim(lstInvoice.ListItems(CC).SubItems(63))
                Call VerifyNonPayableDiag
            End If
            If Len(Trim(lstInvoice.ListItems(CC).SubItems(64))) Then
                VDiagCode = Trim(lstInvoice.ListItems(CC).SubItems(64))
                Call VerifyNonPayableDiag
            End If
            If Len(Trim(lstInvoice.ListItems(CC).SubItems(65))) Then
                VDiagCode = Trim(lstInvoice.ListItems(CC).SubItems(65))
                Call VerifyNonPayableDiag
            End If
            If Len(Trim(lstInvoice.ListItems(CC).SubItems(93))) Then
                VDiagCode = Trim(lstInvoice.ListItems(CC).SubItems(93))
                Call VerifyNonPayableDiag
            End If
            If Trim(lstInvoice.ListItems(CC).SubItems(6)) = "" Then
                Call ErrRpt
                lstErrRec.SubItems(9) = "Invoice No Can't be null ! "
            End If
            If IsDate(lstInvoice.ListItems(CC).SubItems(8)) Then
                If CDate(lstInvoice.ListItems(CC).SubItems(8)) < CDate(VEffDate) Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "Invoice Date MUST BE GREATER than Clinical Effective date ! "
                End If
                If CDate(lstInvoice.ListItems(CC).SubItems(8)) > CDate(VExpDate) + 14 Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "Invoice date IS GREATER THAN Clinical Expiry date + 14 days ! "
                End If
    '           If IsDate(lstInvoice.ListItems(CC).SubItems(8)) Then
                If (CDate(Date) - CDate(lstInvoice.ListItems(CC).SubItems(8))) > 45 Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "45 Days LATE ! "
                End If
            End If
            If IsDate(lstInvoice.ListItems(CC).SubItems(9)) Then
                If CDate(lstInvoice.ListItems(CC).SubItems(9)) < CDate(VEffDate) Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "Visit date MUST BE AFTER the Effective date ! "
                End If
                If CDate(lstInvoice.ListItems(CC).SubItems(9)) > CDate(VExpDate) + 14 Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "Visit date IS GREATER THAN Expiry date + 14 days ! "
                End If
            ElseIf lstInvoice.ListItems(CC).SubItems(9) <> "__/__/____" Then
            
            End If
            If IsNumeric(lstInvoice.ListItems(CC).SubItems(70)) Then ' Bill Amt > Balance Limit
                If CDbl(lstInvoice.ListItems(CC).SubItems(13)) > CDbl(lstInvoice.ListItems(CC).SubItems(70)) Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "Bill Amt > Balance Limit ! "
                End If
            End If
            
            If IsNumeric(lstInvoice.ListItems(CC).SubItems(71)) Then 'Bill Amt > Visit Limit
                If CDbl(lstInvoice.ListItems(CC).SubItems(13)) > CDbl(lstInvoice.ListItems(CC).SubItems(71)) Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "Bill Amt > Visit Limit ! "
                End If
            End If
            If Trim(lstInvoice.SelectedItem.SubItems(72)) <> "" Then
                ChkSql = "SELECT Catcode, Amount FROM CheckItems WHERE CatCode = 'CSMD'" & _
                        " and CatName ='" & lstInvoice.ListItems(CC).SubItems(72) & "'"
                ChKItemRec.Open ChkSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                Do While Not ChKItemRec.EOF
                    If IsNumeric(ChKItemRec!Amount) Then
                        If CDbl(lstInvoice.ListItems(CC).SubItems(13)) > CDbl(ChKItemRec!Amount) Then
                            Call ErrRpt
                            lstErrRec.SubItems(9) = "Bill Amount > Consultation Medication - " & lstInvoice.SelectedItem.SubItems(72) & "Card"
                        End If
                    End If
                'End If
                ChKItemRec.MoveNext
                Loop
                ChKItemRec.Close
                Set ChKItemRec = Nothing
            End If
        Next CC
        TxtTotInvoiceAmt = IIf(IsNumeric(TxtTotInvoiceAmt), TxtTotInvoiceAmt, 0)
        If CInt(TotBillAmt) <> CInt(TxtTotInvoiceAmt) Then
            Set lstErrRec = lstErr.ListItems.Add(, , " ")
            lstErrRec.SubItems(1) = " "
            lstErrRec.SubItems(2) = " "
            lstErrRec.SubItems(3) = " "
            lstErrRec.SubItems(4) = " "
            lstErrRec.SubItems(5) = " "
            lstErrRec.SubItems(6) = " "
            lstErrRec.SubItems(7) = "Invoice Amt = " & CDbl(TxtTotInvoiceAmt)
            lstErrRec.SubItems(8) = "Total Bill Amt = " & CDbl(TotBillAmt)
            lstErrRec.SubItems(9) = "Invoice Amt is different from Total Bill Amt"
            'CmdVerify.Visible = True
            'CmdSave.Visible = False
            'CmdSave.Enabled = False
        Else
            'CmdVerify.Visible = False
            'CmdSave.Visible = True
            'CmdSave.Enabled = True
        End If
        If HvSave = True Then
            Set lstErrRec = lstErr.ListItems.Add(, , " ")
            lstErrRec.SubItems(9) = "Record(s) has been Saved"
            'CmdVerify.Visible = True
            'CmdSave.Visible = False
            'CmdSave.Enabled = False
        End If
        If lstErr.ListItems.Count > 0 Then
            Call ExportToExcelError(lstErr)
        Else
            MsgBox "Verifed Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
        End If
        CmdSave.Enabled = True
    End If
End Sub

Private Sub CmdView_Click()
    If CmdView.Caption = "&View" Then
        lstInvoice.Top = 240
        lstInvoice.Height = 6000
        CmdView.Caption = "&Back"
    Else
        CmdView.Caption = "&View"
        lstInvoice.Top = 6650
        lstInvoice.Height = 1300
    End If

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
    Call ActionLogUpdate(12, "frmInvoice")
    bDataStatus = True
    bExit = False
    EditFlag = False
    Call ClearEntries
    Call EntriesEnabled(False)
    Call ComboListing
    CmdClearAll.Enabled = True
    txtLGNo.Enabled = True
    'CmdSave.Enabled = False
    'CmdSave.Visible = False
    'IntIndex = 0
'sl    If CheckAccess(117) = True Then
        strPayTo = "C"
'sl    ElseIf CheckAccess(118) = True Then
'sl        strPayTo = "M"
'sl    End If
End Sub

Private Sub ClearEntries()
   On Error Resume Next
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Name <> "txtMBMNo" And ctl.Name <> "txtInvNo" And ctl.Name <> "TxtClinicCode" _
                And ctl.Name <> "TxtTotInvoiceAmt" And ctl.Name <> "txtClinicName" Then
               ctl.Text = ""
            End If
        ElseIf TypeOf ctl Is ComboBox Then
        
'        ElseIf TypeOf ctl Is ListView Then
'            ctl.ListItems.Clear
        ElseIf TypeOf ctl Is MaskEdBox Then
            If ctl.Name <> "txtInvDate" And ctl.Name <> "txtInvRecDate" Then
                ctl.mask = ""
                ctl.Text = ""
                ctl.mask = "##/##/####"
            End If
            If ctl.Name = "txtVisitTime" Then
                ctl.mask = ""
                ctl.Text = ""
                ctl.mask = "##:##"
            End If
        End If
    Next ctl
    Z1 = 0
    Z2 = 0
    Z3 = 0
    Z4 = 0
    Z5 = 0
    Z6 = 0
    Z7 = 0
    Z8 = 0
    lblDispBillAmt = ""
    lblDispCLNBillAmt = ""
    lblBillAmt = ""
    lblCLNBillAmt = ""
    lblFFS = ""
    lblPay2CLN = ""
    lblActCoPay = ""
    A29 = ""
    ReasonOtherIND = ""
    TxtDiagnosisOth.Enabled = False
    TxtDiagDesc5.Visible = True
    CmdUpdate.Visible = False
    CmdEdit.Visible = True
    CmdEdit.Enabled = True
    cboFFS = ""
    cboClaimability = "A - Accepted"
    cboMedType1(0).Clear
    cboMedType1(1).Clear
    cboMedType1(2).Clear
    cboMedType1(3).Clear
    cboMedType1(4).Clear
    cboMedType1(5).Clear
    cboMedType1(6).Clear
    cboMedType1(0) = ""
    cboMedType1(1) = ""
    cboMedType1(2) = ""
    cboMedType1(3) = ""
    cboMedType1(4) = ""
    cboMedType1(5) = ""
    cboMedType1(6) = ""
    

End Sub


Private Sub ClearMBMInfo()
    strMBMno = ""
    strPayor = ""
    strClnPln = ""
    strClnIns = ""
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
    strAccumAmt = ""
    strCLNContract = ""
    strMBMCLNLTM = ""
    strCoPay = ""
    MembershipNo = ""
    CASENumber = ""
    txtMBMNo = ""
    txtTabMBMNo = ""
    txtPatientName = ""
    TxtPayor = ""
    txtPlanCode = ""
    TxtInsType = ""
    txtPatientName = ""
    txtEffDate.mask = "##/##/####"
    txtExpDate.mask = "##/##/####"
    TxtPolNo = ""
    txtPolStatus = ""
    txtMBMStatus = ""
    txtdob = ""
    txtPatientIC = ""
    txtPatientName = ""
    TxtBalLimit = ""
    TxtVisitLimit = ""
    'TxtCardType = ""
    txtLTM = ""
    txtDept = ""
    txtMCDaysBal = ""
End Sub
Private Sub ComboListing()
'Dim ClinicRec As New ADODB.Recordset
        
'    txtClinicName.Clear
'    ClinicRec.Open "SELECT CLNCode, CLNName FROM CLN WHERE CLNPanel = 'Y'", medixmasterconn, adOpenKeyset, adLockOptimistic
'    Do Until ClinicRec.EOF
'        txtClinicName.AddItem ClinicRec!CLNCode & " - " & ClinicRec!CLNName
'        ClinicRec.MoveNext
'    Loop
'    ClinicRec.Close
'    Set ClinicRec = Nothing

Dim InvRec As New ADODB.Recordset
Dim InvMth As New ADODB.Recordset

    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

        CboInvRecMode.Clear
        InvRec.Open "Select RecModeCode, RecModeDesc From InvRecMode ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do Until InvRec.EOF
            CboInvRecMode.AddItem InvRec!RecModeCode & " - " & Trim(InvRec!RecModeDesc)
            InvRec.MoveNext
        Loop
        InvRec.Close
        Set InvRec = Nothing
        
        
        InvMth.Open "Select INVProcMth, INVProcYR From INVProcMonth ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        cboProcMth.Clear
        Do Until InvMth.EOF
            cboProcMth.AddItem InvMth!INVProcMth & "/" & InvMth!INVProcYR
            InvMth.MoveNext
        Loop
        InvMth.Close
        Set InvMth = Nothing
        
        cboFFS.Clear
        cboFFS.AddItem "Y" & " - " & "Yes"
        cboFFS.AddItem "N" & " - " & "No"
        
        cboClaimability.Clear
        cboClaimability.AddItem "A" & " - " & "Accepted"
        cboClaimability.AddItem "M" & " - " & "Memo"
        cboClaimability.AddItem "R" & " - " & "Rejected"
    
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
    
End Sub

Private Sub lblDispBillAmt_Change()
Dim a As Double
Dim B As Double

    a = IIf(IsNumeric(lblDispBillAmt), lblDispBillAmt, 0)
    B = IIf(IsNumeric(lblActCoPay), lblActCoPay, 0)
    If CDbl(a) <> "0" Then
        lblBillAmt = Format(CDbl(a) - CDbl(B), "#,##0.00")
    End If
End Sub


Private Sub lblDispCLNBillAmt_Change()
Dim a As Double
Dim B As Double

    a = IIf(IsNumeric(lblDispCLNBillAmt), lblDispCLNBillAmt, 0)
    B = IIf(IsNumeric(A29), A29, 0)
    If CDbl(a) <> "0" Then
        lblCLNBillAmt = CDbl(a) - CDbl(B)
    End If
End Sub


Private Sub lstInvoice_Click()
Dim tmpMBMTYpe As String

 CmdRemove.Enabled = True
 '   If EditFlag = True Then
        If lstInvoice.ListItems.Count Then
            CmdUpdate.Enabled = True
            
            txtMBMNo = ""
            txtMBMNo = lstInvoice.SelectedItem.SubItems(2) & "-" & lstInvoice.SelectedItem.SubItems(3)
            MembershipNo = lstInvoice.SelectedItem.SubItems(2)
            CovPersonID = lstInvoice.SelectedItem.SubItems(3)
            Call MemberSuppInfo(MembershipNo, CovPersonID, strBalLimit, strVisitLimit, strAccumAmt, _
                            StrDentStatus, strCoPay, strPayor, strHLTCode, strClnPln, _
                            strClnIns, strEffdate, strExpDate, strPolNo, strMBMDataStatus, _
                            strPOlStatus, strPatientName, strPatientDOB, strPatientIC, _
                            strMBMCLNLTM, strCLNContract, strVisitIND, strAnnualInd, strSpecStatus, _
                            strMEPStatus, strDentLmt, strSpecLmt, strMEPLmt, SuppLmtType, _
                            StrDentLmtStatus, StrSpecLmtStatus, strBNFStatus, strDentExtLmt, _
                            str1stSpecLmt, strGRPMCOStatus, strPLNGrossStatus, strMBMDepartment, strMCStatus, strMBMMCDaysBal, _
                            strGRPPostageFees, strGRPBordxFees, strPLNMMPerc, Strrelationtype, _
                            strMBMCancelDate, strCMBMCancelDate, strCMBMCCLNExpDate)

            txtTabMBMNo = lstInvoice.SelectedItem.SubItems(2)
            txtTabCoverID = lstInvoice.SelectedItem.SubItems(3)
            TxtPayor = strPayor
            
            txtPlanCode = lstInvoice.SelectedItem.SubItems(67)
            TxtInsType = lstInvoice.SelectedItem.SubItems(68)
            txtPatientName = lstInvoice.SelectedItem.SubItems(4)
            txtEffDate = lstInvoice.SelectedItem.SubItems(78)
            txtExpDate = lstInvoice.SelectedItem.SubItems(79)
            TxtPolNo = lstInvoice.SelectedItem.SubItems(76)
            txtPolStatus = strPOlStatus
            txtMBMStatus = strMBMDataStatus
            txtdob = strPatientDOB
            txtPatientIC = lstInvoice.SelectedItem.SubItems(77)
            txtPatientName = lstInvoice.SelectedItem.SubItems(4)
            TxtBalLimit = lstInvoice.SelectedItem.SubItems(70)
            TxtVisitLimit = lstInvoice.SelectedItem.SubItems(71)
            If ASOInd = "Y" Then
               txtAccumAmt = UtilisedAmt
            Else
                txtAccumAmt = strAccumAmt
            End If
            tmpMBMTYpe = "P"
            If Trim(CovPersonID) <> "00" Then
                tmpMBMTYpe = "S"
            End If
            
            If ASOInd = "Y" Then
            
            Dim ASOqry As String
            Dim ASORS As New ADODB.Recordset
            Dim sMBMNumber() As String
            Dim MBMNumber As String
            Dim coverid As String
            Dim ASOlimit As Double
            Dim IPLimit As Double
            Dim OPLimit As Double
            Dim IsCombined As String
            
                If (InStr(1, txtMBMNo, "-")) > 0 Then
                    sMBMNumber = Split(txtMBMNo, "-")
                    MBMNumber = sMBMNumber(0)
                    coverid = sMBMNumber(1)
                Else
                    MBMNumber = txtMBMNo
                    coverid = "00"
                End If
                
                ASOqry = ""
                ASOqry = "Select MBMASOLimit,MBMIPLimit,MBMOPLimit,MBMIPUtilized,MBMOPUtilized ,IsIPOPCombined  from Tbl_ASOMemberDetails  where MBMNumber= '" + MBMNumber + "' and MBMCoverID='" + coverid + "'"
                ASORS.Open ASOqry, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                Do Until ASORS.EOF
                     ASOlimit = IIf(Len(ASORS!MBMASOLimit), ASORS!MBMASOLimit, 0)
                     IPLimit = IIf(Len(ASORS!MBMIPLimit), ASORS!MBMIPLimit, 0)
                     OPLimit = IIf(Len(ASORS!MBMOPLimit), ASORS!MBMOPLimit, 0)
                     IsCombined = IIf(Len(ASORS!IsIPOPCombined), ASORS!IsIPOPCombined, "N")
                ASORS.MoveNext
                Loop
                ASORS.Close
                Set ASORS = Nothing
                
                If IsCombined = "Y" Then
                    txtAnnLmt = ASOlimit
                Else
                    txtAnnLmt = OPLimit
                End If
            Else
                txtAnnLmt = GetAnnualLimit(strClnIns, strClnPln, strHLTCode, CDate(strEffdate), tmpMBMTYpe, strPayor, SuppLmtType)
            End If
            
            
           ' TxtCardType = lstInvoice.SelectedItem.SubItems(72)
            txtBillNo = lstInvoice.SelectedItem.SubItems(7)
            txtInvDate = lstInvoice.SelectedItem.SubItems(8)
            txtVisitDate = lstInvoice.SelectedItem.SubItems(9)
            txtVisitTime = lstInvoice.SelectedItem.SubItems(10)
            txtMC = lstInvoice.SelectedItem.SubItems(11)
            txtClinicName = lstInvoice.SelectedItem.SubItems(12)
            strPayTo = lstInvoice.SelectedItem.SubItems(80)
            strPayee = lstInvoice.SelectedItem.SubItems(81)
            lblDispBillAmt = Format(lstInvoice.SelectedItem.SubItems(13), "#,##0.00")
            lblDispCLNBillAmt = Format(lstInvoice.SelectedItem.SubItems(13), "#,##0.00")
            A1 = lstInvoice.SelectedItem.SubItems(15)
            A2 = lstInvoice.SelectedItem.SubItems(16)
            A3 = lstInvoice.SelectedItem.SubItems(17)
            A4 = lstInvoice.SelectedItem.SubItems(18)
            A5 = lstInvoice.SelectedItem.SubItems(19)
            A6 = lstInvoice.SelectedItem.SubItems(20)
            A7 = lstInvoice.SelectedItem.SubItems(21)
            A8 = lstInvoice.SelectedItem.SubItems(22)
            A9 = lstInvoice.SelectedItem.SubItems(23)
            A10 = lstInvoice.SelectedItem.SubItems(24)
            A11 = lstInvoice.SelectedItem.SubItems(25)
            A12 = lstInvoice.SelectedItem.SubItems(26)
            A13 = lstInvoice.SelectedItem.SubItems(27)
            A14 = lstInvoice.SelectedItem.SubItems(28)
            A15 = lstInvoice.SelectedItem.SubItems(29)
            A16 = lstInvoice.SelectedItem.SubItems(30)
            A17 = lstInvoice.SelectedItem.SubItems(31)
            A18 = lstInvoice.SelectedItem.SubItems(32)
            A19 = lstInvoice.SelectedItem.SubItems(33)
            A20 = lstInvoice.SelectedItem.SubItems(34)
            A21 = lstInvoice.SelectedItem.SubItems(35)
            A22 = lstInvoice.SelectedItem.SubItems(36)
            A23 = lstInvoice.SelectedItem.SubItems(37)
            A24 = lstInvoice.SelectedItem.SubItems(38)
            A25 = lstInvoice.SelectedItem.SubItems(39)
            A26 = lstInvoice.SelectedItem.SubItems(40)
            A27 = lstInvoice.SelectedItem.SubItems(41)
            A28 = lstInvoice.SelectedItem.SubItems(42)
            lblActCoPay = lstInvoice.SelectedItem.SubItems(43)
            A30 = lstInvoice.SelectedItem.SubItems(101)
            A31 = lstInvoice.SelectedItem.SubItems(114)
            A32 = lstInvoice.SelectedItem.SubItems(474)
            A33 = lstInvoice.SelectedItem.SubItems(475)
            txtPOthDesc = lstInvoice.SelectedItem.SubItems(44)
            txtLOthDesc = lstInvoice.SelectedItem.SubItems(45)
            txtIOthDesc = lstInvoice.SelectedItem.SubItems(46)
            
            cboMedType1(0) = lstInvoice.SelectedItem.SubItems(47)
            cboMedType1(1) = lstInvoice.SelectedItem.SubItems(48)
            cboMedType1(2) = lstInvoice.SelectedItem.SubItems(49)
            cboMedType1(3) = lstInvoice.SelectedItem.SubItems(50)
            cboMedType1(4) = lstInvoice.SelectedItem.SubItems(51)
            txtMedDay1(0) = lstInvoice.SelectedItem.SubItems(52)
            txtMedDay1(1) = lstInvoice.SelectedItem.SubItems(53)
            txtMedDay1(2) = lstInvoice.SelectedItem.SubItems(54)
            txtMedDay1(3) = lstInvoice.SelectedItem.SubItems(55)
            txtMedDay1(4) = lstInvoice.SelectedItem.SubItems(56)
            txtMedQty1(0) = lstInvoice.SelectedItem.SubItems(57)
            txtMedQty1(1) = lstInvoice.SelectedItem.SubItems(58)
            txtMedQty1(2) = lstInvoice.SelectedItem.SubItems(59)
            txtMedQty1(3) = lstInvoice.SelectedItem.SubItems(60)
            txtMedQty1(4) = lstInvoice.SelectedItem.SubItems(61)
            txtDiagnosis1 = lstInvoice.SelectedItem.SubItems(62)
            txtDiagnosis2 = lstInvoice.SelectedItem.SubItems(63)
            txtDiagnosis3 = lstInvoice.SelectedItem.SubItems(64)
            txtDiagnosis4 = lstInvoice.SelectedItem.SubItems(65)
            TxtPayor = lstInvoice.SelectedItem.SubItems(66)
            
            cboMedType1(5) = lstInvoice.SelectedItem.SubItems(102)
            cboMedType1(6) = lstInvoice.SelectedItem.SubItems(103)
            txtMedDay1(5) = lstInvoice.SelectedItem.SubItems(104)
            txtMedDay1(6) = lstInvoice.SelectedItem.SubItems(105)
            txtMedQty1(5) = lstInvoice.SelectedItem.SubItems(106)
            txtMedQty1(6) = lstInvoice.SelectedItem.SubItems(107)
            txtInvRecDate = lstInvoice.SelectedItem.SubItems(110)
            CboInvRecMode = lstInvoice.SelectedItem.SubItems(111)
 
            SuppLmtType = lstInvoice.SelectedItem.SubItems(82)
            B1 = lstInvoice.SelectedItem.SubItems(83)
            B2 = lstInvoice.SelectedItem.SubItems(84)
            B3 = lstInvoice.SelectedItem.SubItems(85)
            B4 = lstInvoice.SelectedItem.SubItems(86)
            B5 = lstInvoice.SelectedItem.SubItems(87)
            B6 = lstInvoice.SelectedItem.SubItems(88)
            B7 = lstInvoice.SelectedItem.SubItems(89)
            B8 = lstInvoice.SelectedItem.SubItems(90)
            Z1 = lstInvoice.SelectedItem.SubItems(83)
            Z2 = lstInvoice.SelectedItem.SubItems(84)
            Z3 = lstInvoice.SelectedItem.SubItems(85)
            Z4 = lstInvoice.SelectedItem.SubItems(86)
            Z5 = lstInvoice.SelectedItem.SubItems(87)
            Z6 = lstInvoice.SelectedItem.SubItems(88)
            Z7 = lstInvoice.SelectedItem.SubItems(89)
            Z8 = lstInvoice.SelectedItem.SubItems(90)
            C1 = lstInvoice.SelectedItem.SubItems(91)
            C2 = lstInvoice.SelectedItem.SubItems(92)
            TxtDiagDesc5.Visible = True
            TxtDiagnosis5 = lstInvoice.SelectedItem.SubItems(93)
            ReasonOtherIND = lstInvoice.SelectedItem.SubItems(94)
            If ReasonOtherIND = "Y" Then
                TxtDiagDesc5.Visible = False
                TxtDiagnosisOth.Enabled = True
            Else
                TxtDiagnosisOth = ""
            End If
            TxtDiagnosisOth = lstInvoice.SelectedItem.SubItems(95)
            TxtDiagDesc1 = lstInvoice.SelectedItem.SubItems(96)
            TxtDiagDesc2 = lstInvoice.SelectedItem.SubItems(97)
            TxtDiagDesc3 = lstInvoice.SelectedItem.SubItems(98)
            TxtDiagDesc4 = lstInvoice.SelectedItem.SubItems(99)
            TxtDiagDesc5 = lstInvoice.SelectedItem.SubItems(100)
            TxtRemark = lstInvoice.SelectedItem.SubItems(112)
            txtMBMAdjAnnLmt = lstInvoice.SelectedItem.SubItems(113)
            ' --- RESERVED until (150) - last use 114
            lblPay2CLN = Format(lstInvoice.SelectedItem.SubItems(158), "#,##0.00")
            lblFFS = Format(lstInvoice.SelectedItem.SubItems(468), "#,##0.00")
            cboFFS = lstInvoice.SelectedItem.SubItems(469)
            A29 = Format(lstInvoice.SelectedItem.SubItems(471), "#,##0.00")
            txtCLNBillAmt = IIf(IsNumeric(lstInvoice.SelectedItem.SubItems(472)), Format(lstInvoice.SelectedItem.SubItems(472), "#,##0.00"), 0)
            cboClaimability = lstInvoice.SelectedItem.SubItems(473)
            txtMCDaysBal = IIf(IsNumeric(strMBMMCDaysBal), strMBMMCDaysBal, 0)
            CmdAddRec.Enabled = False
        End If
'    End If
End Sub


Private Sub lstInvoice_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        'SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstInvoice, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstInvoice, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
    Case 11
        SortListView lstInvoice, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lstInvoice, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    Case 13
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 19
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(19)
    Case 20
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(21)
    Case 22
        SortListView lstInvoice, ColumnHeader.Index, ldtString, m_blnDirection(22)
    End Select

End Sub




Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub


Private Sub TxtClinicCode_GotFocus()
Dim CLNRec As New ADODB.Recordset
Dim CLNSql As String
Dim UpCln
    Call EntriesEnabled(False)
    TxtClinicCode = ChkStr(TxtClinicCode)
    cboFFS = ""
    If Trim(TxtClinicCode) <> "" Then
        txtClinicName = ""
        lblClnType = "G"
        
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
'        CLNSql = "SELECT CLNCode, CLNName FROM CLN WHERE (CLNPanel = 'C' OR CLNPanel = 'Y')" & _
                " AND (CLNCode ='" & Trim(TxtClinicCode) & "')"
        CLNSql = "SELECT CLNPanel, CLNCode, CLNName, CLNType,CLNFFSStatus FROM CLN WHERE CLNCode ='" & Trim(TxtClinicCode) & "'"

        CLNRec.Open CLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If CLNRec.RecordCount <= 0 Then
            MsgBox "No such Clinic Code ! ", vbOKOnly + vbCritical, DialogBoxTitle
            TxtClinicCode.SetFocus
        ElseIf Trim(CLNRec!CLNPanel) = "C" Or Trim(CLNRec!CLNPanel) = "Y" Then
            txtClinicName = Trim(CLNRec!CLNName)
            lblClnType = IIf(Len(Trim(CLNRec!CLNType)), Trim(CLNRec!CLNType), "G")
            If Trim(txtMBMNo) <> "" Then
                Call EntriesEnabled(True)
            End If
            cboFFS = IIf(Len(Trim(CLNRec!CLNFFSStatus)), Trim(CLNRec!CLNFFSStatus), "")
        Else
        '    UpCln = MsgBox("Non-Panel Clinic !, Press 'Yes' to Proceed, 'No' to Cancel", vbYesNo + vbExclamation)
        '    If UpCln = vbYes Then
        '        If CheckAccess(137) = True Then 'Plan
        '            If FormLoadedByName("frmMaintenanceEntity") = True Then
        '                frmMaintenanceEntity.SetFocus
        '            Else
        '                frmMaintenanceEntity.Show
        '            End If
        '        Else
        '           MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
        '        End If
        '    ElseIf UpCln = vbNo Then
        '        TxtClinicCode.SetFocus
        '    End If
        End If
        CLNRec.Close
        Set CLNRec = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
    End If
End Sub

Private Sub TxtClinicCode_LostFocus()
Dim CLNRec As New ADODB.Recordset
Dim CLNSql As String
Dim UpCln
    TxtClinicCode = ChkStr(TxtClinicCode)
    cboFFS = ""
    If Trim(TxtClinicCode) <> "" Then
        txtClinicName = ""
        lblClnType = "G"
        
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
'        CLNSql = "SELECT CLNCode, CLNName FROM CLN WHERE (CLNPanel = 'C' OR CLNPanel = 'Y')" & _
                " AND (CLNCode ='" & Trim(TxtClinicCode) & "')"
        CLNSql = "SELECT CLNPanel, CLNCode, CLNName, CLNType,CLNFFSStatus FROM CLN WHERE CLNCode ='" & Trim(TxtClinicCode) & "'"

        CLNRec.Open CLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If CLNRec.RecordCount <= 0 Then
            MsgBox "No such Clinic Code ! ", vbOKOnly + vbCritical, DialogBoxTitle
            TxtClinicCode.SetFocus
        ElseIf Trim(CLNRec!CLNPanel) = "C" Or Trim(CLNRec!CLNPanel) = "Y" Then
            txtClinicName = Trim(CLNRec!CLNName)
            lblClnType = IIf(Len(Trim(CLNRec!CLNType)), Trim(CLNRec!CLNType), "G")
            If Trim(txtMBMNo) <> "" Then
                Call EntriesEnabled(True)
            End If
            cboFFS = IIf(Len(Trim(CLNRec!CLNFFSStatus)), Trim(CLNRec!CLNFFSStatus), "")
        Else
        '    UpCln = MsgBox("Non-Panel Clinic !, Press 'Yes' to Proceed, 'No' to Cancel", vbYesNo + vbExclamation)
        '    If UpCln = vbYes Then
        '        If CheckAccess(137) = True Then 'Plan
        '            If FormLoadedByName("frmMaintenanceEntity") = True Then
        '                frmMaintenanceEntity.SetFocus
        '            Else
        '                frmMaintenanceEntity.Show
        '            End If
        '        Else
        '           MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
        '        End If
        '    ElseIf UpCln = vbNo Then
        '        TxtClinicCode.SetFocus
        '    End If
        End If
        CLNRec.Close
        Set CLNRec = Nothing
    End If
    
End Sub

Private Sub txtBillNo_LostFocus()
    txtBillNo = ChkStr(txtBillNo)
End Sub



Private Sub txtCLNBillAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtDiagnosis1_LostFocus()
Dim tmpDesc As String
    
    tmpDesc = ""
    If Trim(txtDiagnosis1) <> "" Then
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            Call VerifyDiagnosis(txtDiagnosis1, tmpDesc)
            If tmpDesc <> "" Then
                TxtDiagDesc1 = tmpDesc
            End If
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
    End If
End Sub

Private Sub txtDiagnosis2_LostFocus()
Dim tmpDesc As String
    
    tmpDesc = ""
    If Trim(txtDiagnosis2) <> "" Then
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            Call VerifyDiagnosis(txtDiagnosis2, tmpDesc)
            If tmpDesc <> "" Then
                TxtDiagDesc2 = tmpDesc
            End If
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
    End If
End Sub

Private Sub txtDiagnosis3_LostFocus()
Dim tmpDesc As String
    
    tmpDesc = ""
    If Trim(txtDiagnosis3) <> "" Then
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            Call VerifyDiagnosis(txtDiagnosis3, tmpDesc)
            If tmpDesc <> "" Then
                TxtDiagDesc3 = tmpDesc
            End If
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
    End If
End Sub

Private Sub txtDiagnosis4_LostFocus()
Dim tmpDesc As String
    
    tmpDesc = ""
    If Trim(txtDiagnosis4) <> "" Then
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        Call VerifyDiagnosis(txtDiagnosis4, tmpDesc)
        If tmpDesc <> "" Then
            TxtDiagDesc4 = tmpDesc
        End If
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
    End If
End Sub

Private Sub TxtDiagnosis5_LostFocus()
Dim tmpDesc As String
Dim TreatRec As New ADODB.Recordset
Dim TreatSql As String

    TxtDiagnosisOth.Enabled = False
    TxtDiagDesc5.Visible = True
    tmpDesc = ""
    If Trim(TxtDiagnosis5) = "" Then
        TxtDiagDesc5 = ""
        TxtDiagnosisOth = ""
    Else
        
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        Call VerifyDiagnosis(TxtDiagnosis5, tmpDesc)
        If tmpDesc <> "" Then
            TxtDiagDesc5 = tmpDesc
        End If
        TreatSql = "Select ReasonOtherDiag from TreatmentReason WHERE ReasonCode='" & Trim(TxtDiagnosis5) & "'"
        TreatRec.Open TreatSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not TreatRec.EOF
            ReasonOtherIND = IIf(Len(Trim(TreatRec!ReasonOtherDiag)), Trim(TreatRec!ReasonOtherDiag), "N")
            If ReasonOtherIND = "Y" Then
                TxtDiagDesc5.Visible = False
                TxtDiagnosisOth.Enabled = True
                TxtDiagnosisOth.SetFocus
            Else
                TxtDiagnosisOth.Enabled = True
                TxtDiagnosisOth = ""
            End If
        TreatRec.MoveNext
        Loop
        TreatRec.Close
        Set TreatRec = Nothing
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
    End If
    
End Sub

Private Sub TxtDiagnosisOth_LostFocus()
    TxtDiagnosisOth = ChkStr(TxtDiagnosisOth)
End Sub

Private Sub txtInvRecDate_LostFocus()
    If txtInvRecDate <> "__/__/____" Then
        If IsDate(txtInvRecDate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtInvRecDate.SetFocus
        ElseIf CDate(txtInvRecDate) > CDate(Date) Then
            MsgBox "Invoice Date cannot greater than system date ! ", vbCritical
            txtInvRecDate.SetFocus
        End If
    End If
End Sub

Private Sub txtIOthDesc_LostFocus()
    txtIOthDesc = ChkStr(txtIOthDesc)
End Sub

Private Sub txtLOthDesc_LostFocus()
    txtLOthDesc = ChkStr(txtLOthDesc)
End Sub

Private Sub txtMBMAdjAnnLmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMBMAdjAnnLmt_LostFocus()
Dim tmpAdjAmt As String
Dim tmpBillAmt As String
Dim tmpBalLmt As String
Dim tmpNonGP As Double

    tmpAdjAmt = CDbl(IIf(IsNumeric(txtMBMAdjAnnLmt), txtMBMAdjAnnLmt, 0))
    If tmpAdjAmt <= 0 Then
        If IsNumeric(TxtBalLimit) Then
            tmpBillAmt = IIf(IsNumeric(lblDispBillAmt), lblDispBillAmt, 0)
            tmpNonGP = IIf(IsNumeric(A31), A31, 0)
            If CDbl(tmpBillAmt - tmpNonGP) > CDbl(TxtBalLimit) Then
                tmpAdjAmt = CDbl(CDbl(tmpBillAmt) - tmpNonGP - CDbl(TxtBalLimit))
                txtMBMAdjAnnLmt = tmpAdjAmt
            End If
        End If
    End If
End Sub

Private Sub txtMBMNo_GotFocus()
    cmdGo.Default = True
End Sub

Private Sub txtMBMNo_LostFocus()
    cmdGo.Default = False
    If Trim(txtMBMNo) = "" Then
        txtPatientIC = ""
        txtPatientName = ""
    End If
End Sub

Private Sub EntriesEnabled(EStatus As Boolean)
    FrameInv.Enabled = EStatus
    FrameDet.Enabled = EStatus
    FrameDiag.Enabled = EStatus
    FrameMed.Enabled = EStatus
    If EStatus = True Then
        B1.Enabled = True
        B2.Enabled = True
        B3.Enabled = True
        B4.Enabled = True
        B5.Enabled = False
        B6.Enabled = False
        B7.Enabled = False
        B8.Enabled = False
        If Trim(lblClnType) = "S" Then
            B1.Enabled = False
            B2.Enabled = False
            B3.Enabled = False
            B4.Enabled = False
            B5.Enabled = True
            B6.Enabled = True
            B7.Enabled = True
            B8.Enabled = True
        End If
    End If
'    FrameInvoice.Enabled = EStatus
End Sub

Private Sub txtMC_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub ChkData()
Dim CISCasRec As New ADODB.Recordset
Dim CISCasSql, ChkDupKey As String
Dim XX, QQ, BB As Integer
Dim AddNewInv As Boolean
Dim DupInvoice
Dim ArrChk() As String
Dim TempCLNBillAmt As Double
Dim tmpDulInvStatus As Boolean

    Call getTotal
    
    If Trim(txtMBMNo) = "" Then
        MsgBox "Please keyin Member No ! ", vbCritical
        txtMBMNo.SetFocus
    ElseIf txtVisitDate = "__/__/____" Or txtVisitDate = "" Then
        MsgBox "Please Enter Visit Date! ", vbCritical
        txtVisitDate.SetFocus
    ElseIf IsDate(txtVisitDate) = False Then
        MsgBox "Invalid Visit date format!", vbCritical
        txtVisitDate.SetFocus
    ElseIf CDate(txtVisitDate) > CDate(Date) Then
        MsgBox "Visit Date cannot greater than system date ! ", vbCritical
        txtVisitDate.SetFocus
    ElseIf IsDate(txtEffDate) = False Then
        MsgBox "Invalid Effective date format!", vbCritical
        txtVisitDate.SetFocus
    ElseIf IsDate(txtExpDate) = False Then
        MsgBox "Invalid Expiry date format!", vbCritical
        txtVisitDate.SetFocus
    ElseIf CDate(txtVisitDate) < CDate(txtEffDate) Then
        MsgBox "Visit Date cannot less than Membership Effective date ! ", vbCritical
        txtVisitDate.SetFocus
    ElseIf CDate(txtVisitDate) > CDate(txtExpDate) Then
        MsgBox "Visit Date cannot greater than Membership Expiry date ! ", vbCritical
        txtVisitDate.SetFocus
    ElseIf Trim(TxtClinicCode) = "" Then
        MsgBox "Please keyin the Clinic code ! ", vbCritical
        TxtClinicCode.SetFocus
    ElseIf Trim(txtInvNo) = "" Then
        MsgBox "Please Enter Invoice No ! ", vbCritical
        txtInvNo.SetFocus
    ElseIf txtInvDate = "__/__/____" Or Trim(txtInvDate) = "" Then
        MsgBox "Please Enter Invoice Date! ", vbCritical
        txtInvDate.SetFocus
    ElseIf txtInvRecDate = "__/__/____" Or Trim(txtInvRecDate) = "" Then
        MsgBox "Please Enter Invoice Receipt Date ! ", vbCritical
        txtInvRecDate.SetFocus
'    ElseIf Trim(CboInvRecMode.Text) = "" Then
'        MsgBox "Please Enter Invoice Receipt Mode ! ", vbCritical
'        CboInvRecMode.SetFocus
    ElseIf Trim(TxtTotInvoiceAmt) = "" Then
        MsgBox "Please Keyin Invoice Amount ! ", vbCritical
        TxtTotInvoiceAmt.SetFocus
    ElseIf IntInvAmt = 0 Then
        MsgBox "Total Invoice Amount cannot be Zero ! ", vbCritical
        B1.Enabled = True
        B1.SetFocus
    ElseIf Trim(ReasonOtherIND) = "Y" And Trim(TxtDiagnosisOth) = "" Then
        MsgBox "Please Keyin Others Diagnosis ! ", vbCritical
        TxtDiagnosisOth.Enabled = True
        TxtDiagnosisOth.SetFocus
    ElseIf Trim(txtCLNBillAmt) = "" Then
        MsgBox "Please enter Bill Amount! ", vbCritical
        txtCLNBillAmt.SetFocus
    ElseIf txtCLNBillAmt <> CDbl(lblCLNBillAmt) + CDbl(IIf(IsNumeric(A29), A29, 0)) Then
        MsgBox "Please check Bill Amount! ", vbCritical
        txtCLNBillAmt.SetFocus
    ElseIf Trim(cboClaimability) = "" Then
        MsgBox "Please Enter Claimability ", vbCritical
    Else
        tmpDulInvStatus = False
        Erase ArrChk
        ReDim ArrChk(0)
        QQ = 0
        CISCasSql = ""
        'edit by m.shan 23/02/2010 remove InvoiceStmtAmt checking from original statement.
        CISCasSql = "SELECT CoverID,DateVisit,ClinicCode,InvoiceStmtAmt,MBMNumber FROM ciscas where MBMNumber = '" & Trim(MembershipNo) & _
                    "' AND COVERID ='" & CovPersonID & "' AND DateVisit = '" & Format(txtVisitDate, "mm/dd/yyyy") & "' " & _
                    " AND ClinicCode ='" & Trim(TxtClinicCode) & "'"
        CISCasRec.Open CISCasSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        If CISCasRec.EOF = False Then
            Do While Not CISCasRec.EOF
                QQ = QQ + 1
                ReDim Preserve ArrChk(QQ)
                ArrChk(QQ) = Trim(CISCasRec!MBMNumber & "~" & CISCasRec!coverid & "#" & CISCasRec!DateVisit & "*" & CISCasRec!ClinicCode & "*" & CISCasRec!InvoiceStmtAmt)
                CISCasRec.MoveNext
            Loop
            'If CISCasRec.RecordCount > 0 Then
                DupInvoice = MsgBox("Invoice has been entered !, Press 'Yes' to Proceed, 'No' to Cancel", vbYesNo + vbExclamation)
                If DupInvoice = vbYes Then
                    If EditFlag = True Then
                        Call RemoveRecFrListV
                    End If
                    frmWorksheet.InvoiceClick = 1
                    Call AssignData2ListView
                    Load frmWorksheet
                    Unload frmWorksheet
                    EditFlag = False
                    cmdVerify_Click
                    FrameMBMCLN.Enabled = True
                    txtMBMNo.SetFocus
                ElseIf DupInvoice = vbNo Then
                    Call EntriesEnabled(False)
                End If
            'End If
            CISCasRec.Close
            Set CISCasRec = Nothing
            Exit Sub
        End If
        
        For BB = 1 To lstInvoice.ListItems.Count
            ReDim Preserve ArrChk(QQ + BB)
            ArrChk(QQ + BB) = Trim(lstInvoice.ListItems(BB).SubItems(2) & "~" & lstInvoice.ListItems(BB).SubItems(3) & "#" & lstInvoice.ListItems(BB).SubItems(9) & "*" & lstInvoice.ListItems(BB).SubItems(14) & "*" & lstInvoice.ListItems(BB).SubItems(162))
        Next BB
        ChkDupKey = MembershipNo & "~" & CovPersonID & "#" & txtVisitDate & "*" & TxtClinicCode & "*" & lblDispBillAmt
        AddNewInv = True
        For XX = 0 To UBound(ArrChk)
            If EditFlag = False Then
            If Trim(ChkDupKey) = ArrChk(XX) Then
                AddNewInv = False
                If Trim(txtBillNo) = "" Then
                    DupInvoice = MsgBox("Invoice has been entered !, Press 'Yes' to Proceed, 'No' to Cancel", vbYesNo + vbExclamation)
                    If DupInvoice = vbYes Then
                        If EditFlag = True Then
                            Call RemoveRecFrListV
                        End If
                        frmWorksheet.InvoiceClick = 1
                        Call AssignData2ListView
                        Load frmWorksheet
                        Unload frmWorksheet
                        EditFlag = False
                        cmdVerify_Click
                        FrameMBMCLN.Enabled = True
                        txtMBMNo.SetFocus
                    ElseIf DupInvoice = vbNo Then
                        Call EntriesEnabled(False)
                    End If
                Else
                    MsgBox "Record Duplicated, Same Invoice has been entered ", vbCritical, "Error"
                End If
            End If
            End If
        Next XX
        If AddNewInv = True Then
            If EditFlag = True Then
                Call RemoveRecFrListV
            End If
            Call AssignData2ListView
            Load frmWorksheet
            Unload frmWorksheet
            EditFlag = False
            FrameMBMCLN.Enabled = True
            txtMBMNo.SetFocus
            cmdVerify_Click
        End If
        CmdMBMInfo.Enabled = False
        FrameMBMCLN.Enabled = True
    End If
End Sub

Private Function getTotal() As Double
Dim ctl  As Control
Dim tmpBillAmt As Double
Dim tmpBalLmt As Double
Dim tmpAdjAmt As Double
Dim tmpNonGP As Double

    A1 = 0
    A2 = 0
    A3 = 0
    If Z2 > 0 Then
        A3 = C1
    Else
        If Z3 > 0 And Z4 > 0 Then
            A2 = C1
        Else
            A1 = C1
        End If
    End If
    
    A4 = 0
    A5 = 0
    A6 = 0
    If Z6 > 0 Then
        A6 = C2
    Else
        If Z7 > 0 And Z8 > 0 Then
            A5 = C2
        Else
            A4 = C2
        End If
    End If
    IntInvAmt = 0
    For Each ctl In Me.Controls
        If Mid(ctl.Name, 1, 1) = "A" And Len(ctl.Name) <= 7 Then
            If IsNumeric(ctl.Text) = True Then
                IntInvAmt = IntInvAmt + CDbl(ctl.Text)
            End If
        End If
    Next
    IntInvAmt = Format(CDbl(IntInvAmt) - CDbl(IIf(IsNumeric(A29), A29, 0)), "#,##0.00")
    getTotal = CDbl(IntInvAmt)
    
    txtMBMAdjAnnLmt.Enabled = False
    If IsNumeric(TxtBalLimit) Then
        tmpBillAmt = IIf(IsNumeric(IntInvAmt), IntInvAmt, 0)
        tmpNonGP = IIf(IsNumeric(A31), A31, 0)
        If (CDbl(tmpBillAmt) - CDbl(tmpNonGP)) > CDbl(TxtBalLimit) Then
            txtMBMAdjAnnLmt.Enabled = True
            tmpAdjAmt = CDbl(tmpBillAmt) - CDbl(tmpNonGP) - CDbl(TxtBalLimit)
            txtMBMAdjAnnLmt = Format(tmpAdjAmt, "###0.00")
        Else
            txtMBMAdjAnnLmt = ""
        End If
    End If
End Function

Private Sub AssignData2ListView()
Dim strPayee As String

    'CmdSave.Enabled = True
    Set lstRec = lstInvoice.ListItems.Add(, , "")
    'frmInvoice
    If strPayTo = "C" Then
'        strPayee = Trim(Left(txtClinicName, IIf(InStr(1, txtClinicName, "-") - 1 < 0, Len(txtClinicName), InStr(1, txtClinicName, "-") - 1)))
        strPayee = TxtClinicCode
    Else
        strPayee = "MEMBER"
    End If
    
    Call getTotal
    lstRec.SubItems(2) = MembershipNo
    lstRec.SubItems(3) = CovPersonID
    lstRec.SubItems(4) = txtPatientName
    lstRec.SubItems(6) = IIf(Len(Trim(txtInvNo)), Trim(txtInvNo), "")
    lstRec.SubItems(7) = IIf(Len(Trim(txtBillNo)), txtBillNo, "")
    lstRec.SubItems(8) = IIf(IsDate(txtInvDate), txtInvDate, "__/__/____")
    lstRec.SubItems(9) = IIf(IsDate(txtVisitDate), txtVisitDate, "__/__/____")
    lstRec.SubItems(10) = IIf(isTime(txtVisitTime), txtVisitTime, "__:__")
    lstRec.SubItems(11) = IIf(IsNumeric(txtMC), txtMC, 0)
    lstRec.SubItems(12) = Trim(TxtClinicCode)
    lstRec.SubItems(13) = IntInvAmt
    lstRec.SubItems(14) = Logon
    lstRec.SubItems(15) = IIf(IsNumeric(A1), A1, 0)
    lstRec.SubItems(16) = IIf(IsNumeric(A2), A2, 0)
    lstRec.SubItems(17) = IIf(IsNumeric(A3), A3, 0)
    lstRec.SubItems(18) = IIf(IsNumeric(A4), A4, 0)
    lstRec.SubItems(19) = IIf(IsNumeric(A5), A5, 0)
    lstRec.SubItems(20) = IIf(IsNumeric(A6), A6, 0)
    lstRec.SubItems(21) = IIf(IsNumeric(A7), A7, 0)
    lstRec.SubItems(22) = IIf(IsNumeric(A8), A8, 0)
    lstRec.SubItems(23) = IIf(IsNumeric(A9), A9, 0)
    lstRec.SubItems(24) = IIf(IsNumeric(A10), A10, 0)
    lstRec.SubItems(25) = IIf(IsNumeric(A11), A11, 0)
    lstRec.SubItems(26) = IIf(IsNumeric(A12), A12, 0)
    lstRec.SubItems(27) = IIf(IsNumeric(A13), A13, 0)
    lstRec.SubItems(28) = IIf(IsNumeric(A14), A14, 0)
    lstRec.SubItems(29) = IIf(IsNumeric(A15), A15, 0)
    lstRec.SubItems(30) = IIf(IsNumeric(A16), A16, 0)
    lstRec.SubItems(31) = IIf(IsNumeric(A17), A17, 0)
    lstRec.SubItems(32) = IIf(IsNumeric(A18), A18, 0)
    lstRec.SubItems(33) = IIf(IsNumeric(A19), A19, 0)
    lstRec.SubItems(34) = IIf(IsNumeric(A20), A20, 0)
    lstRec.SubItems(35) = IIf(IsNumeric(A21), A21, 0)
    lstRec.SubItems(36) = IIf(IsNumeric(A22), A22, 0)
    lstRec.SubItems(37) = IIf(IsNumeric(A23), A23, 0)
    lstRec.SubItems(38) = IIf(IsNumeric(A24), A24, 0)
    lstRec.SubItems(39) = IIf(IsNumeric(A25), A25, 0)
    lstRec.SubItems(40) = IIf(IsNumeric(A26), A26, 0)
    lstRec.SubItems(41) = IIf(IsNumeric(A27), A27, 0)
    lstRec.SubItems(42) = IIf(IsNumeric(A28), A28, 0)
    lstRec.SubItems(43) = IIf(IsNumeric(lblActCoPay), Format(lblActCoPay, "#,##0.00"), 0)
    lstRec.SubItems(44) = IIf(Len(Trim(txtPOthDesc)), Trim(txtPOthDesc), "")
    lstRec.SubItems(45) = IIf(Len(Trim(txtLOthDesc)), Trim(txtLOthDesc), "")
    lstRec.SubItems(46) = IIf(Len(Trim(txtIOthDesc)), Trim(txtIOthDesc), "")
    lstRec.SubItems(47) = IIf(Len(Trim(cboMedType1(0))), Trim(cboMedType1(0)), "")
    lstRec.SubItems(48) = IIf(Len(Trim(cboMedType1(1))), Trim(cboMedType1(1)), "")
    lstRec.SubItems(49) = IIf(Len(Trim(cboMedType1(2))), Trim(cboMedType1(2)), "")
    lstRec.SubItems(50) = IIf(Len(Trim(cboMedType1(3))), Trim(cboMedType1(3)), "")
    lstRec.SubItems(51) = IIf(Len(Trim(cboMedType1(4))), Trim(cboMedType1(4)), "")
    lstRec.SubItems(52) = IIf(Len(Trim(txtMedDay1(0))), Trim(txtMedDay1(0)), "")
    lstRec.SubItems(53) = IIf(Len(Trim(txtMedDay1(1))), Trim(txtMedDay1(1)), "")
    lstRec.SubItems(54) = IIf(Len(Trim(txtMedDay1(2))), Trim(txtMedDay1(2)), "")
    lstRec.SubItems(55) = IIf(Len(Trim(txtMedDay1(3))), Trim(txtMedDay1(3)), "")
    lstRec.SubItems(56) = IIf(Len(Trim(txtMedDay1(4))), Trim(txtMedDay1(4)), "")
    lstRec.SubItems(57) = IIf(Len(Trim(txtMedQty1(0))), Trim(txtMedQty1(0)), "")
    lstRec.SubItems(58) = IIf(Len(Trim(txtMedQty1(1))), Trim(txtMedQty1(1)), "")
    lstRec.SubItems(59) = IIf(Len(Trim(txtMedQty1(2))), Trim(txtMedQty1(2)), "")
    lstRec.SubItems(60) = IIf(Len(Trim(txtMedQty1(3))), Trim(txtMedQty1(3)), "")
    lstRec.SubItems(61) = IIf(Len(Trim(txtMedQty1(4))), Trim(txtMedQty1(4)), "")
    lstRec.SubItems(62) = IIf(Len(Trim(txtDiagnosis1)), Trim(txtDiagnosis1), "")
    lstRec.SubItems(63) = IIf(Len(Trim(txtDiagnosis2)), Trim(txtDiagnosis2), "")
    lstRec.SubItems(64) = IIf(Len(Trim(txtDiagnosis3)), Trim(txtDiagnosis3), "")
    lstRec.SubItems(65) = IIf(Len(Trim(txtDiagnosis4)), Trim(txtDiagnosis4), "")
    lstRec.SubItems(66) = TxtPayor
    lstRec.SubItems(67) = txtPlanCode
    lstRec.SubItems(68) = TxtInsType
    lstRec.SubItems(69) = strHLTCode
    lstRec.SubItems(70) = TxtBalLimit
    lstRec.SubItems(71) = TxtVisitLimit
    'lstRec.SubItems(72) = TxtCardType
    lstRec.SubItems(73) = strVisitIND
    lstRec.SubItems(74) = strAnnualInd
    lstRec.SubItems(75) = TxtCoPay
    lstRec.SubItems(76) = TxtPolNo
    lstRec.SubItems(77) = strPatientIC
    lstRec.SubItems(78) = txtEffDate
    lstRec.SubItems(79) = txtExpDate
    lstRec.SubItems(80) = strPayTo
    lstRec.SubItems(81) = strPayee
    lstRec.SubItems(82) = SuppLmtType
    lstRec.SubItems(83) = IIf(IsNumeric(B1), B1, 0)
    lstRec.SubItems(84) = IIf(IsNumeric(B2), B2, 0)
    lstRec.SubItems(85) = IIf(IsNumeric(B3), B3, 0)
    lstRec.SubItems(86) = IIf(IsNumeric(B4), B4, 0)
    lstRec.SubItems(87) = IIf(IsNumeric(B5), B5, 0)
    lstRec.SubItems(88) = IIf(IsNumeric(B6), B6, 0)
    lstRec.SubItems(89) = IIf(IsNumeric(B7), B7, 0)
    lstRec.SubItems(90) = IIf(IsNumeric(B8), B8, 0)
    lstRec.SubItems(91) = IIf(IsNumeric(C1), C1, 0)
    lstRec.SubItems(92) = IIf(IsNumeric(C2), C2, 0)
    lstRec.SubItems(93) = IIf(Len(Trim(TxtDiagnosis5)), Trim(TxtDiagnosis5), "")
    lstRec.SubItems(94) = ReasonOtherIND
    If Trim(ReasonOtherIND) = "Y" Then
        lstRec.SubItems(95) = Trim(TxtDiagnosisOth)
    Else
        lstRec.SubItems(95) = ""
    End If
    lstRec.SubItems(96) = IIf(Len(Trim(TxtDiagDesc1)), Trim(TxtDiagDesc1), "")
    lstRec.SubItems(97) = IIf(Len(Trim(TxtDiagDesc2)), Trim(TxtDiagDesc2), "")
    lstRec.SubItems(98) = IIf(Len(Trim(TxtDiagDesc3)), Trim(TxtDiagDesc3), "")
    lstRec.SubItems(99) = IIf(Len(Trim(TxtDiagDesc4)), Trim(TxtDiagDesc4), "")
    lstRec.SubItems(100) = IIf(Len(Trim(TxtDiagDesc5)), Trim(TxtDiagDesc5), "")
    lstRec.SubItems(101) = IIf(IsNumeric(A30), A30, 0)
    
    lstRec.SubItems(102) = IIf(Len(Trim(cboMedType1(5))), Trim(cboMedType1(5)), "")
    lstRec.SubItems(103) = IIf(Len(Trim(cboMedType1(6))), Trim(cboMedType1(6)), "")
    lstRec.SubItems(104) = IIf(Len(Trim(txtMedDay1(5))), Trim(txtMedDay1(5)), "")
    lstRec.SubItems(105) = IIf(Len(Trim(txtMedDay1(6))), Trim(txtMedDay1(6)), "")
    lstRec.SubItems(106) = IIf(Len(Trim(txtMedQty1(5))), Trim(txtMedQty1(5)), "")
    lstRec.SubItems(107) = IIf(Len(Trim(txtMedQty1(6))), Trim(txtMedQty1(6)), "")
    lstRec.SubItems(110) = txtInvRecDate
    lstRec.SubItems(111) = Trim(Left(CboInvRecMode, IIf(InStr(1, CboInvRecMode, "-") - 1 < 0, Len(CboInvRecMode), InStr(1, CboInvRecMode, "-") - 1)))
    lstRec.SubItems(112) = IIf(Len(Trim(TxtRemark)), Trim(TxtRemark), "")
    lstRec.SubItems(113) = IIf(IsNumeric(txtMBMAdjAnnLmt), txtMBMAdjAnnLmt, 0)
    lstRec.SubItems(114) = IIf(IsNumeric(A31), A31, 0)
    lstRec.SubItems(115) = IIf(IsNumeric(txtExcess), txtExcess, 0)
    lstRec.SubItems(474) = IIf(IsNumeric(A32), A32, 0)
    lstRec.SubItems(475) = IIf(IsNumeric(A33), A33, 0)
    lstRec.SubItems(469) = IIf(Len(cboFFS), Mid(cboFFS, 1, 1), "")
    lstRec.SubItems(470) = IIf(Len(cboProcMth), cboProcMth, "")
    lstRec.SubItems(471) = IIf(IsNumeric(A29), Format(A29, "#,##0.00"), 0)
    lstRec.SubItems(472) = IIf(IsNumeric(txtCLNBillAmt), Format(txtCLNBillAmt, "#,##0.00"), 0)
    lstRec.SubItems(473) = Mid(cboClaimability, 1, 1)
    lstRec.SubItems(478) = IIf(IsNumeric(txtLmtExceeded), txtLmtExceeded, 0)
    ' -- --- reserved for invoice until (150) - last use 114
End Sub
Private Sub txtMC_LostFocus()
    If Trim(strMCStatus) = "Y" Then
        If Len(Trim(txtMBMNo)) And IsNumeric(txtMC) Then
            If CDbl(strMBMMCDaysBal) <= CDbl(txtMC) Then
                MsgBox "MC Days Exceeded " & (CDbl(strMBMMCDaysBal) - CDbl(txtMC)) * -1 & " Day(s)!", vbCritical
            End If
        End If
    End If
End Sub

Private Sub txtMedDay1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMedQty1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPOthDesc_LostFocus()
    txtPOthDesc = ChkStr(txtPOthDesc)
End Sub

Private Sub txtInvDate_LostFocus()
    If txtInvDate <> "__/__/____" Then
        If IsDate(txtInvDate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtInvDate.SetFocus
        ElseIf CDate(txtInvDate) > CDate(Date) Then
            MsgBox "Invoice Date cannot greater than system date ! ", vbCritical
            txtInvDate.SetFocus
        End If
    End If
End Sub

Private Sub txtInvNo_LostFocus()
    txtInvNo = ChkStr(txtInvNo)
End Sub

Private Sub TxtRemark_lostfocus()
    TxtRemark = ChkStr(TxtRemark)
End Sub

Private Sub TxtTotInvoiceAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub



Private Sub txtTyphoid_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtVisitDate_LostFocus()
    If txtVisitDate <> "__/__/____" Then
        CmdAddRec.Enabled = True
        If IsDate(txtVisitDate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtVisitDate.SetFocus
        ElseIf CDate(txtVisitDate) > CDate(Date) Then
            MsgBox "Visit Date cannot greater than system date ! ", vbCritical
            txtVisitDate.SetFocus
        ElseIf CovPersonID = "00" And strPOlStatus <> "A" And Len(strMBMCancelDate) <> "0" Then
            If (CDate(txtVisitDate) > CDate(strMBMCancelDate)) Then
                MsgBox "Membership has been terminated before the entered visit date.", vbCritical
                CmdAddRec.Enabled = False
                txtVisitDate.SetFocus
            End If
        ElseIf CovPersonID = "00" And strPOlStatus <> "A" And CDate(txtVisitDate) > CDate(txtExpDate) Then
            MsgBox "Visit date entered is not in range of Policy Effective Date and Expiry Date", vbCritical
            CmdAddRec.Enabled = False
            txtVisitDate.SetFocus
        ElseIf CovPersonID <> "00" And strPOlStatus <> "A" And Len(strCMBMCancelDate) <> "0" Then
            If (CDate(txtVisitDate) > CDate(strCMBMCancelDate)) Then
                MsgBox "Membership has been terminated before the entered visit date.", vbCritical
                CmdAddRec.Enabled = False
                txtVisitDate.SetFocus
            End If
        ElseIf CovPersonID <> "00" And strPOlStatus <> "A" And Len(strCMBMCCLNExpDate) <> "0" Then
            If CDate(txtVisitDate) > CDate(strCMBMCCLNExpDate) Then
                MsgBox "Visit date entered is not in range of Policy Effective Date and Expiry Date", vbCritical
                CmdAddRec.Enabled = False
                txtVisitDate.SetFocus
            End If
        ElseIf strPOlStatus = "A" And CovPersonID = "00" And CDate(txtVisitDate) > CDate(txtExpDate) Then
            MsgBox "Visit date entered is not in range of Policy Effective Date and Expiry Date", vbCritical
            CmdAddRec.Enabled = False
        ElseIf strPOlStatus = "A" And CovPersonID <> "00" And Len(strCMBMCCLNExpDate) <> "0" Then
            If CDate(txtVisitDate) > CDate(strCMBMCCLNExpDate) Then
                MsgBox "Visit date entered is not in range of Policy Effective Date and Expiry Date", vbCritical
                CmdAddRec.Enabled = False
            End If
        End If
    End If
End Sub
Private Sub txtVisitTime_LostFocus()
    If txtVisitTime <> "__:__" Then
        If isTime(txtVisitTime) = False Then
            MsgBox "Invalid Time Format! ", vbCritical
            txtVisitTime.SetFocus
        End If
    End If
End Sub
Private Sub RemoveRecFrListV()
On Error Resume Next
Dim ForLoop As Integer
Dim ListCount As Integer
Dim StrSql As String
Dim DEL As New ADODB.Recordset

    ListCount = 0
    If Trim(lstInvoice.SelectedItem.SubItems(2)) = "" Then
        MsgBox "Please Choose a Record", vbOKOnly + vbCritical, DialogBoxTitle
        lstInvoice.SetFocus
    Else
            ListCount = lstInvoice.ListItems.Count
            For ForLoop = 1 To ListCount
                If lstInvoice.ListItems(ForLoop).Selected = True Then
                      lstInvoice.ListItems.Remove (ForLoop)
                End If
            Next ForLoop

    End If
End Sub

Private Sub VerifyNonPayableDiag()
Dim Diag As New ADODB.Recordset
Dim DiagSql As String

'    If VDiagIND = "0" Then
        DiagSql = "SELECT * FROM DiagNotPayable WHERE PLNCode='" & VPlan & "' AND DiagCode='" & VDiagCode & "'"
        Diag.Open DiagSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
'    ElseIf VDiagIND = "1" Then
'        DiagSql = "SELECT * FROM GenDiagNotPay WHERE PLNCode='" & VPlan & "' AND DiagCode='" & VDiagCode & "'"
'        Diag.Open DiagSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
'    End If
    Do While Not Diag.EOF
        Call ErrRpt
        lstErrRec.SubItems(9) = "Diagnosis '" & VDiagCode & "' is not payable ! "
    
    Diag.MoveNext
    Loop
    Diag.Close
    Set Diag = Nothing
End Sub


Private Sub ErrRpt()
        
    Set lstErrRec = lstErr.ListItems.Add(, , lstInvoice.ListItems(CC).SubItems(2))
        lstErrRec.SubItems(1) = lstInvoice.ListItems(CC).SubItems(3)
        lstErrRec.SubItems(2) = lstInvoice.ListItems(CC).SubItems(4)
        lstErrRec.SubItems(3) = lstInvoice.ListItems(CC).SubItems(76)
        lstErrRec.SubItems(4) = lstInvoice.ListItems(CC).SubItems(77)
        lstErrRec.SubItems(5) = lstInvoice.ListItems(CC).SubItems(78)
        lstErrRec.SubItems(6) = lstInvoice.ListItems(CC).SubItems(79)
        lstErrRec.SubItems(7) = lstInvoice.ListItems(CC).SubItems(70)
        lstErrRec.SubItems(8) = lstInvoice.ListItems(CC).SubItems(71)
End Sub

Private Function GenerateCISCaseNo(PayorCode)
    Dim CISCAQ As New ADODB.Recordset
    Dim StrSql As String
    Dim CISSeqNo As String
    Dim CAQCharacter As String
    Dim Prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    
    Dim cmd As New ADODB.Command
    
    'CISCAQ.Close
    'Set CISCAQ = Nothing
    
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "WebSearchCASNo_New2013test"
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    'Set Prm1 = cmd.CreateParameter("PayCode", adVarChar, adParamInput, 2000, "DC")
    Set Prm1 = cmd.CreateParameter("PayCode", adVarChar, adParamInput, 2000, "EA")
    cmd.Parameters.Append Prm1
  
    cmd.Parameters.Append cmd.CreateParameter("CaseNo1", adVarChar, adParamOutput, 2000)
    cmd.Execute
    
    'CISSeqNo = Format(cmd("CaseNo1"), "000000000")
    CISSeqNo = Format(cmd("CaseNo1"), "00000000")
    GenerateCISCaseNo = "EA" & CISSeqNo
    'GenerateCISCaseNo = PayorCode & "K" & CISSeqNo

End Function

Private Function GenerateOPSeq(OPVar)
Dim OPRec As New ADODB.Recordset
Dim StrSql As String
Dim SeqNo As String
    
    StrSql = "Select OPCode,OPBatchSeq from OPBatch Where OPCode = '" & OPVar & "'"
    OPRec.Open StrSql, MediConnCISDB, adOpenStatic, adLockOptimistic
    
    If OPRec.RecordCount <= 0 Then
        With OPRec
            .AddNew
            !OPCode = OPVar
            !OPBatchSeq = "00001"
            .update
            SeqNo = "00001"
        End With
    Else
        OPRec!OPBatchSeq = Format(OPRec!OPBatchSeq + 1, "00000")
        SeqNo = OPRec!OPBatchSeq
        OPRec.update
    End If

    GenerateOPSeq = OPVar & SeqNo
    OPRec.Close
    Set OPRec = Nothing
End Function

Private Sub SaveCase()
Dim CISSql As String
Dim tmpDiag1, tmpDiag2, tmpDiag3, tmpDiag4, tmpDiag5 As String
Dim NoOfDiag As Integer
Dim NoofMedi As Integer
Dim MCStat As Integer
Dim tmpRBAmt, tmpRBCoPay, tmpLmtExceed
Dim tmpINVMonth, tmpINVYear As String

    CISSql = ""
    MCStat = 0
    NoOfDiag = 0
    NoofMedi = 0
    tmpRBAmt = 0
    tmpRBCoPay = 0
    tmpLmtExceed = 0
    
    Dim tmpMBMNum As String
    Dim tmpMBMArr() As String
    tmpMBMNum = Trim(lstInvoice.ListItems(AA).SubItems(2))
    Dim Pos As Integer
    Pos = 0
    Pos = InStr(tmpMBMNum, "-")
    If Pos > 0 Then
        tmpMBMArr() = Split(tmpMBMNum, "-")
        tmpMBMNum = tmpMBMArr(0)
    Else
        tmpMBMNum = (tmpMBMNum)
    End If
    
    lstInvoice.ListItems(AA).Text = IIf(Len(OPBatchNo), OPBatchNo, "")
    lstInvoice.ListItems(AA).SubItems(1) = NewCaseNo
    
        tmpCASnumber = NewCaseNo
        lblCaseNumber.Visible = True
        lblCAS.Visible = True
        lblCAS = tmpCASnumber
        tmpMBMNumber = Trim(tmpMBMNum)
        tmpCoverID = lstInvoice.ListItems(AA).SubItems(3)
        tmpCASStatus = "O"
        tmpClaimability = lstInvoice.ListItems(AA).SubItems(473)
        tmpCaseType = "C" ' C - Clinical Case   D - Dental Case
        TmpPending = "N"
        tmpRegisterDate = Format(Date, "yyyy/mm/dd")
        tmpREgisterBy = Logon
        tmpLastUpdateDate = Format(Date, "yyyy/mm/dd")
        tmpLastUpdateBy = Logon

        tmpClinicCode = lstInvoice.ListItems(AA).SubItems(12)
        tmpInvoiceNo = Trim(lstInvoice.ListItems(AA).SubItems(6))
        tmpInvoiceDate = Format(lstInvoice.ListItems(AA).SubItems(8), "yyyy/mm/dd")
        If tmpInvoiceDate <> "" Then
            tmpINVMonth = Format(Month(tmpInvoiceDate), "00")
        Else
            tmpINVMonth = ""
        End If
        
        If tmpInvoiceDate <> "" Then
            tmpINVYear = Year(tmpInvoiceDate)
        Else
            tmpINVYear = ""
        End If
        tmpInvoiceStmtAmt = IIf(IsNumeric(TxtTotInvoiceAmt), TxtTotInvoiceAmt, 0)
        tmpDateVisit = Format(lstInvoice.ListItems(AA).SubItems(9), "yyyy/mm/dd")
        tmpMCDays = lstInvoice.ListItems(AA).SubItems(11)
        tmpOPBatchNo = OPBatchNo
        If isTime(lstInvoice.ListItems(AA).SubItems(10)) Then
            tmpTimeVisit = Format(lstInvoice.ListItems(AA).SubItems(10), "hh:mm")
        End If
        
        tmpMedInjType1 = lstInvoice.ListItems(AA).SubItems(47)
        tmpMedInjType2 = lstInvoice.ListItems(AA).SubItems(48)
        tmpMedInjType3 = lstInvoice.ListItems(AA).SubItems(49)
        tmpMedInjType4 = lstInvoice.ListItems(AA).SubItems(50)
        tmpMedInjType5 = lstInvoice.ListItems(AA).SubItems(51)
        tmpMedInjDays1 = lstInvoice.ListItems(AA).SubItems(52)
        tmpMedInjDays2 = lstInvoice.ListItems(AA).SubItems(53)
        tmpMedInjDays3 = lstInvoice.ListItems(AA).SubItems(54)
        tmpMedInjDays4 = lstInvoice.ListItems(AA).SubItems(55)
        tmpMedInjDays5 = lstInvoice.ListItems(AA).SubItems(56)
        tmpMedInjQty1 = lstInvoice.ListItems(AA).SubItems(57)
        tmpMedInjQty2 = lstInvoice.ListItems(AA).SubItems(58)
        tmpMedInjQty3 = lstInvoice.ListItems(AA).SubItems(59)
        tmpMedInjQty4 = lstInvoice.ListItems(AA).SubItems(60)
        tmpMedInjQty5 = lstInvoice.ListItems(AA).SubItems(61)
        
        tmpMedInjType6 = lstInvoice.ListItems(AA).SubItems(102)
        tmpMedInjType7 = lstInvoice.ListItems(AA).SubItems(103)
        tmpMedInjDays6 = lstInvoice.ListItems(AA).SubItems(104)
        tmpMedInjDays7 = lstInvoice.ListItems(AA).SubItems(105)
        tmpMedInjQty6 = lstInvoice.ListItems(AA).SubItems(106)
        tmpMedInjQty7 = lstInvoice.ListItems(AA).SubItems(107)
        tmpInvoiceReceiptDate = Format(lstInvoice.ListItems(AA).SubItems(110), "yyyy/mm/dd")
        tmpInvoiceReceiptMode = lstInvoice.ListItems(AA).SubItems(111)
        tmpCasRemark = lstInvoice.ListItems(AA).SubItems(112)
        
        tmpCASTreatReason1 = lstInvoice.ListItems(AA).SubItems(62)
        tmpCASTreatReason2 = lstInvoice.ListItems(AA).SubItems(63)
        tmpCASTreatReason3 = lstInvoice.ListItems(AA).SubItems(64)
        tmpCASTreatReason4 = lstInvoice.ListItems(AA).SubItems(65)
        tmpCASTreatReason5 = lstInvoice.ListItems(AA).SubItems(93)
        tmpCASDiagnosis5 = lstInvoice.ListItems(AA).SubItems(95)
        
        If CDbl(lstInvoice.ListItems(AA).SubItems(11)) > 0 Then
            MCStat = 1
        End If
        tmpMCStatus = MCStat
' ----- Number of Diagnosis calculation
        If Trim(lstInvoice.ListItems(AA).SubItems(62)) <> "" Then
            NoOfDiag = NoOfDiag + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(63)) <> "" Then
            NoOfDiag = NoOfDiag + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(64)) <> "" Then
            NoOfDiag = NoOfDiag + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(65)) <> "" Then
            NoOfDiag = NoOfDiag + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(93)) <> "" Then
            NoOfDiag = NoOfDiag + 1
        End If
        tmpNoOfDiag = NoOfDiag
        
' ----- Number of Medication calculation
        If Trim(lstInvoice.ListItems(AA).SubItems(47)) <> "" Then
            NoofMedi = NoofMedi + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(48)) <> "" Then
            NoofMedi = NoofMedi + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(49)) <> "" Then
            NoofMedi = NoofMedi + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(50)) <> "" Then
            NoofMedi = NoofMedi + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(51)) <> "" Then
            NoofMedi = NoofMedi + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(102)) <> "" Then
            NoofMedi = NoofMedi + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(103)) <> "" Then
            NoofMedi = NoofMedi + 1
        End If
        tmpNoOfMedication = NoofMedi
        
        CISSql = "EXEC Save_Case_2013 '" & tmpCASnumber & "' ,'" & tmpMBMNumber & "', '" & tmpCoverID & "', " & _
             "'" & tmpCASStatus & "', '" & tmpClaimability & "', '" & tmpCaseType & "', '" & TmpPending & "', " & _
             "'" & tmpRegisterDate & "', '" & tmpREgisterBy & "','" & tmpLastUpdateDate & "', '" & tmpLastUpdateBy & "', " & _
             "'" & tmpClinicCode & "', '" & tmpInvoiceNo & "','" & tmpInvoiceDate & "', " & tmpInvoiceStmtAmt & ", " & _
             "'" & tmpDateVisit & "', " & tmpMCDays & ", '" & tmpOPBatchNo & "','" & tmpTimeVisit & "', '" & tmpMedInjType1 & "', " & _
             "'" & tmpMedInjType2 & "', '" & tmpMedInjType3 & "', '" & tmpMedInjType4 & "', '" & tmpMedInjType5 & "', " & _
             "'" & tmpMedInjDays1 & "' , '" & tmpMedInjDays2 & "', '" & tmpMedInjDays3 & "', '" & tmpMedInjDays4 & "'," & _
             "'" & tmpMedInjDays5 & "', '" & tmpMedInjQty1 & "', '" & tmpMedInjQty2 & "', '" & tmpMedInjQty3 & "'," & _
             "'" & tmpMedInjQty4 & "', '" & tmpMedInjQty5 & "', " & _
             "'" & tmpMedInjType6 & "', '" & tmpMedInjType7 & "', '" & tmpMedInjDays6 & "', " & _
             "'" & tmpMedInjDays7 & "', '" & tmpMedInjQty6 & "', '" & tmpMedInjQty7 & "','" & tmpInvoiceReceiptDate & "', " & _
             "'" & tmpInvoiceReceiptMode & "', '" & tmpCasRemark & "', '" & tmpCASTreatReason1 & "', '" & tmpCASTreatReason2 & "'," & _
             "'" & tmpCASTreatReason3 & "', '" & tmpCASTreatReason4 & "', '" & tmpCASTreatReason5 & "'," & _
             "'" & tmpCASDiagnosis5 & "', '" & tmpMCStatus & "', " & tmpNoOfDiag & "," & tmpNoOfMedication & ", '" & tmpINVMonth & "', '" & tmpINVYear & "'"
        
    MediConnCISDB.Execute CISSql
    
    'CISSql = "Select CASNumber, MBMNumber, CoverID, CASStatus, Claimability, " & _
    '         "Pending, CaseType, REgisterDate, REgisterBy, LastUpdateDate, LastUpdateBy, ClinicCode, " & _
    '         "InvoiceNo, InvoiceDate, InvoiceStmtAmt, DateVisit, MCDays, OPBatchNo, TimeVisit, " & _
    '         "MedInjType1, MedInjType2, MedInjType3, MedInjType4, MedInjType5, MedInjType6, MedInjType7, " & _
    '         "MedInjDays1, MedInjDays2, MedInjDays3, MedInjDays4, MedInjDays5, MedInjDays6, MedInjDays7, " & _
    '         "MedInjQty1, MedInjQty2, MedInjQty3, MedInjQty4, MedInjQty5, MedInjQty6, MedInjQty7, " & _
    '         "InvoiceReceiptDate, InvoiceReceiptMode, CASTreatReason1, CASTreatReason2, CASTreatReason3, " & _
    '         "CASTreatReason4, CASTreatReason5, CASDiagnosis5, CasRemark, MCStatus, NoOfDiag, NoOfMedication " & _
    '         "FROM CISCAS "

    'CISCAS.Open CISSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    'CISCAS.AddNew
    
    'With CISCAS
        '!CasNumber = NewCaseNo
        '!MBMNumber = Trim(lstInvoice.ListItems(AA).SubItems(2))
        '!CoverID = lstInvoice.ListItems(AA).SubItems(3)
        '!CASStatus = "O"
        '!Claimability = lstInvoice.ListItems(AA).SubItems(473)
        '!CaseType = "C" ' C - Clinical Case   D - Dental Case
        '!Pending = "N"
        '!RegisterDate = Date
        '!REgisterBy = Logon
        '!LastUpdateDate = Date
        '!LastUpdateBy = Logon

        '!ClinicCode = lstInvoice.ListItems(AA).SubItems(12)
        '!InvoiceNo = Trim(lstInvoice.ListItems(AA).SubItems(6))
        '!InvoiceDate = lstInvoice.ListItems(AA).SubItems(8)
        '!InvoiceStmtAmt = IIf(IsNumeric(TxtTotInvoiceAmt), TxtTotInvoiceAmt, 0)
        '!DateVisit = lstInvoice.ListItems(AA).SubItems (9)
        '!MCDays = lstInvoice.ListItems(AA).SubItems(11)
        '!OPBatchNo = OPBatchNo
        'If isTime(lstInvoice.ListItems(AA).SubItems(10)) Then
        '    !TimeVisit = lstInvoice.ListItems(AA).SubItems(10)
        'End If
        
        '!MedInjType1 = lstInvoice.ListItems(AA).SubItems(47)
        '!MedInjType2 = lstInvoice.ListItems(AA).SubItems(48)
        '!MedInjType3 = lstInvoice.ListItems(AA).SubItems(49)
        '!MedInjType4 = lstInvoice.ListItems(AA).SubItems(50)
        '!MedInjType5 = lstInvoice.ListItems(AA).SubItems(51)
        '!MedInjDays1 = lstInvoice.ListItems(AA).SubItems(52)
        '!MedInjDays2 = lstInvoice.ListItems(AA).SubItems(53)
        '!MedInjDays3 = lstInvoice.ListItems(AA).SubItems(54)
        '!MedInjDays4 = lstInvoice.ListItems(AA).SubItems(55)
        '!MedInjDays5 = lstInvoice.ListItems(AA).SubItems(56)
        '!MedInjQty1 = lstInvoice.ListItems(AA).SubItems(57)
        '!MedInjQty2 = lstInvoice.ListItems(AA).SubItems(58)
        '!MedInjQty3 = lstInvoice.ListItems(AA).SubItems(59)
        '!MedInjQty4 = lstInvoice.ListItems(AA).SubItems(60)
        '!MedInjQty5 = lstInvoice.ListItems(AA).SubItems(61)
        
        '!MedInjType6 = lstInvoice.ListItems(AA).SubItems(102)
        '!MedInjType7 = lstInvoice.ListItems(AA).SubItems(103)
        '!MedInjDays6 = lstInvoice.ListItems(AA).SubItems(104)
        '!MedInjDays7 = lstInvoice.ListItems(AA).SubItems(105)
        '!MedInjQty6 = lstInvoice.ListItems(AA).SubItems(106)
        '!MedInjQty7 = lstInvoice.ListItems(AA).SubItems(107)
        '!InvoiceReceiptDate = lstInvoice.ListItems(AA).SubItems(110)
        '!InvoiceReceiptMode = lstInvoice.ListItems(AA).SubItems(111)
        '!CasRemark = lstInvoice.ListItems(AA).SubItems(112)
        
        '!CASTreatReason1 = lstInvoice.ListItems(AA).SubItems(62)
        '!CASTreatReason2 = lstInvoice.ListItems(AA).SubItems(63)
        '!CASTreatReason3 = lstInvoice.ListItems(AA).SubItems(64)
        '!CASTreatReason4 = lstInvoice.ListItems(AA).SubItems(65)
        '!CASTreatReason5 = lstInvoice.ListItems(AA).SubItems(93)
        '!CASDiagnosis5 = lstInvoice.ListItems(AA).SubItems(95)
        
        'If CDbl(lstInvoice.ListItems(AA).SubItems(11)) > 0 Then
         '   MCStat = 1
        'End If
        '!MCStatus = MCStat
        
        'If Trim(lstInvoice.ListItems(AA).SubItems(62)) <> "" Then
         '   NoOfDiag = NoOfDiag + 1
        'End If
        'If Trim(lstInvoice.ListItems(AA).SubItems(63)) <> "" Then
        '    NoOfDiag = NoOfDiag + 1
        'End If
        'If Trim(lstInvoice.ListItems(AA).SubItems(64)) <> "" Then
        '    NoOfDiag = NoOfDiag + 1
        'End If
        'If Trim(lstInvoice.ListItems(AA).SubItems(65)) <> "" Then
            'NoOfDiag = NoOfDiag + 1
        'End If
        'If Trim(lstInvoice.ListItems(AA).SubItems(93)) <> "" Then
        '    NoOfDiag = NoOfDiag + 1
        'End If
        '!NoOfDiag = NoOfDiag
        
' ----- Number of Medication calculation
        'If Trim(lstInvoice.ListItems(AA).SubItems(47)) <> "" Then
        '    NoofMedi = NoofMedi + 1
        'End If
        'If Trim(lstInvoice.ListItems(AA).SubItems(48)) <> "" Then
        '    NoofMedi = NoofMedi + 1
        'End If
        'If Trim(lstInvoice.ListItems(AA).SubItems(49)) <> "" Then
        '    NoofMedi = NoofMedi + 1
        'End If
        'If Trim(lstInvoice.ListItems(AA).SubItems(50)) <> "" Then
        '    NoofMedi = NoofMedi + 1
        'End If
        'If Trim(lstInvoice.ListItems(AA).SubItems(51)) <> "" Then
        '    NoofMedi = NoofMedi + 1
        'End If
        'If Trim(lstInvoice.ListItems(AA).SubItems(102)) <> "" Then
        '    NoofMedi = NoofMedi + 1
        'End If
        'If Trim(lstInvoice.ListItems(AA).SubItems(103)) <> "" Then
        '    NoofMedi = NoofMedi + 1
        'End If
        '!NoOfMedication = NoofMedi
    'End With
    'CISCAS.Update
    
    'CISCAS.Close
    'Set CISCAS = Nothing
End Sub

Private Sub SaveInvoice()
Dim ClaimInvoice  As New ADODB.Recordset
Dim SaveInvoice As String
        SaveInvoice = ""
      
        tmpCASnumber = NewCaseNo
        tmpINVEnteredDate = Format(Date, "yyyy/mm/dd")
        tmpINVEnteredBy = Logon
        tmpINVLastUpdateBy = Logon
        tmpINVLastUpdateDate = Format(Date, "yyyy/mm/dd")
        tmpBillNo = lstInvoice.ListItems(AA).SubItems(7)
        tmpBillDate = Format(lstInvoice.ListItems(AA).SubItems(9), "yyyy/mm/dd")
        tmpINVAmount = Format(lstInvoice.ListItems(AA).SubItems(13), "#,##0.00")
        tmpINVPayTo = lstInvoice.ListItems(AA).SubItems(80)
        tmpINVPayee = lstInvoice.ListItems(AA).SubItems(81)
        tmpINVBillAmt = Format(IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(472)), lstInvoice.ListItems(AA).SubItems(472), 0), "#,##0.00")
        SaveInvoice = ""
    SaveInvoice = "EXEC Save_Invoice '" & tmpCASnumber & "', '" & tmpINVEnteredDate & "', '" & tmpINVEnteredBy & "','" & tmpINVLastUpdateBy & "', '" & _
                  tmpINVLastUpdateDate & "', '" & tmpBillNo & "', '" & tmpBillDate & "'," & CDbl(tmpINVAmount) & ",'" & _
                  tmpINVPayTo & "', '" & tmpINVPayee & "'," & CDbl(tmpINVBillAmt) & ""
    
    MediConnCISWS.Execute SaveInvoice
End Sub

Private Sub saveActual()
     
    'Dim ClaimActual As New ADODB.Recordset
    'Dim StrSql As String
    Dim SaveActualSql As String
    Dim abcefg As Double
        'tmpInvINdex = 0
        SaveActualSql = ""
        'tmpInvINdex = IntIndex
        tmpCaseNo = Trim(NewCaseNo)
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(15)) Then
            tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(15), "#,##0.00")
        Else
            tmpCLMGPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(16)) Then
            tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(16), "#,##0.00")
        Else
            tmpCLMGPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(17)) Then
            tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(17), "#,##0.00")
        Else
            tmpCLMGPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(18)) Then
            tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(18), "#,##0.00")
        Else
            tmpCLMSPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(19)) Then
            tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(19), "#,##0.00")
        Else
            tmpCLMSPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(20)) Then
            tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(20), "#,##0.00")
        Else
            tmpCLMSPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(21)) Then
            tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(21), "#,##0.00")
        Else
            tmpCLMDental = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(22)) Then
            tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(22), "#,##0.00")
        Else
            tmpCLMMaternity = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(23)) Then
            tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(23), "#,##0.00")
        Else
            tmpCLMPRVCheckup = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(24)) Then
            tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(24), "#,##0.00")
        Else
            tmpCLMXRay = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(25)) Then
            tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(25), "#,##0.00")
        Else
            tmpCLMUltraSound = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(26)) Then
            tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(26), "#,##0.00")
        Else
            tmpCLMECG = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(27)) Then
            tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(27), "#,##0.00")
        Else
            tmpCLMDressing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(28)) Then
            tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(28), "#,##0.00")
        Else
            tmpCLMNebulizer = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(29)) Then
            tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(29), "#,##0.00")
        Else
            tmpCLMID = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(30)) Then
            tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(30), "#,##0.00")
        Else
            tmpCLMTSSTO = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(31)) Then
            tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(31), "#,##0.00")
        Else
            tmpCLMMinorSurgery = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(32)) Then
            tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(32), "#,##0.00")
        Else
            tmpCLMPOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(33)) Then
            tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(33), "#,##0.00")
        Else
            tmpCLMBloodDrawing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(34)) Then
            tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(34), "#,##0.00")
        Else
            tmpCLMBloodGlucose = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(35)) Then
            tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(35), "#,##0.00")
        Else
            tmpCLMUrine = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(36)) Then
            tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(36), "#,##0.00")
        Else
            tmpCLMPapSmear = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(37)) Then
            tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(37), "#,##0.00")
        Else
            tmpCLMLOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(38)) Then
            tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(38), "#,##0.00")
        Else
            tmpCLMDTPPolio = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(39)) Then
            tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(39), "#,##0.00")
        Else
            tmpCLMDTPPolioHIB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(40)) Then
            tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(40), "#,##0.00")
        Else
            tmpCLMHepatitisB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(41)) Then
            tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(41), "#,##0.00")
        Else
            tmpCLMMMR = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(42)) Then
            tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(42), "#,##0.00")
        Else
            tmpCLMIOthers = 0
        End If
        tmpCLMPOthersDesc = lstInvoice.ListItems(AA).SubItems(44)
        tmpCLMLOthersDesc = lstInvoice.ListItems(AA).SubItems(45)
        tmpCLMIOthersDesc = lstInvoice.ListItems(AA).SubItems(46)
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(101)) Then
            tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(101), "#,##0.00")
        Else
            tmpCLMVision = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(474)) Then
            tmpCLMTyphiad = lstInvoice.ListItems(AA).SubItems(474)
        Else
            tmpCLMTyphiad = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(475)) Then
            tmpCLMOtherNoBalDed = lstInvoice.ListItems(AA).SubItems(475)
        Else
            tmpCLMOtherNoBalDed = 0
        End If
        'SaveActualSql = "EXEC Save_ClaimActual '" & tmpInvINdex & "', '" & tmpCaseNo & "', " & tmpCLMGPFeeBasicMed & ", " & tmpCLMGPFeeLTMed & ", " & _

        SaveActualSql = "EXEC Save_ClaimActual '" & tmpCaseNo & "', " & tmpCLMGPFeeBasicMed & ", " & tmpCLMGPFeeLTMed & ", " & _
                        "" & tmpCLMGPFeeMedInj & ", " & tmpCLMSPFeeBasicMed & ", " & tmpCLMSPFeeLTMed & ", " & tmpCLMSPFeeMedInj & ", " & _
                        "" & tmpCLMDental & ", " & tmpCLMMaternity & "," & tmpCLMPRVCheckup & ", " & tmpCLMXRay & ", " & _
                        "" & tmpCLMUltraSound & ", " & tmpCLMECG & ", " & tmpCLMDressing & ", " & tmpCLMNebulizer & ", " & _
                        "" & tmpCLMID & ", " & tmpCLMTSSTO & ", " & tmpCLMMinorSurgery & ", " & _
                        "" & tmpCLMPOthers & ", " & tmpCLMBloodDrawing & ", " & tmpCLMBloodGlucose & ", " & tmpCLMUrine & ", " & _
                        "" & tmpCLMPapSmear & ", " & tmpCLMLOthers & ", " & tmpCLMDTPPolio & ", " & tmpCLMDTPPolioHIB & ", " & _
                        "" & tmpCLMHepatitisB & ", " & tmpCLMMMR & ", " & tmpCLMIOthers & ", " & _
                        "'" & tmpCLMPOthersDesc & "', '" & tmpCLMLOthersDesc & "', '" & tmpCLMIOthersDesc & "', " & tmpCLMVision & ", " & _
                        "" & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
    'Debug.Print SaveActualSql
    'ClaimActual.Open "Select INVINdex , * from ClaimInvoiceActualNew ", MediConnCISWS, adOpenKeyset, adLockOptimistic
    'ClaimActual.AddNew
    'ClaimActual!InvINdex = IntIndex
    'IntIndex = 9999
    'Debug.Print "IntIndex: " & IntIndex
    'With ClaimActual
        '!CasNumber = NewCaseNo
        '!CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(15)
        '!CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(16)
        '!CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(17)
        '!CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(18)
        '!CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(19)
        '!CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(20)
        '!CLMDental = lstInvoice.ListItems(AA).SubItems(21)
        '!CLMMaternity = lstInvoice.ListItems(AA).SubItems(22)
        '!CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(23)
        '!CLMXRay = lstInvoice.ListItems(AA).SubItems(24)
        '!CLMUltraSound = lstInvoice.ListItems(AA).SubItems(25)
        '!CLMECG = lstInvoice.ListItems(AA).SubItems(26)
        '!CLMDressing = lstInvoice.ListItems(AA).SubItems(27)
        '!CLMNebulizer = lstInvoice.ListItems(AA).SubItems(28)
        '!CLMID = lstInvoice.ListItems(AA).SubItems(29)
        '!CLMTSSTO = lstInvoice.ListItems(AA).SubItems(30)
        '!CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(31)
        '!CLMPOthers = lstInvoice.ListItems(AA).SubItems(32)
        '!CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(33)
        '!CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(34)
        '!CLMUrine = lstInvoice.ListItems(AA).SubItems(35)
        '!CLMPapSmear = lstInvoice.ListItems(AA).SubItems(36)
        '!CLMLOthers = lstInvoice.ListItems(AA).SubItems(37)
        '!CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(38)
        '!CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(39)
       ' !CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(40)
        '!CLMMMR = lstInvoice.ListItems(AA).SubItems(41)
        '!CLMIOthers = lstInvoice.ListItems(AA).SubItems(42)
        '!CLMPOthersDesc = lstInvoice.ListItems(AA).SubItems(44)
        '!CLMLOthersDesc = lstInvoice.ListItems(AA).SubItems(45)
        '!CLMIOthersDesc = lstInvoice.ListItems(AA).SubItems(46)
        '!CLMVision = lstInvoice.ListItems(AA).SubItems(101)
    'End With
    'ClaimActual.Update
    'ClaimActual.Close
    'Set ClaimActual = Nothing
    
    MediConnCISWS.Execute SaveActualSql
    
End Sub
Private Sub saveActualOthers()
    'Dim ClaimActOTH As New ADODB.Recordset
    Dim StrSql As String
    StrSql = ""
    Dim dblExcess As Double
    dblExcess = 0
    'tmpInvINdex = IntIndex
    tmpCASnumber = NewCaseNo
    If IsNumeric(lstInvoice.ListItems(AA).SubItems(83)) Then
        tmpGPFee = Format(lstInvoice.ListItems(AA).SubItems(83), "#,##0.00")
    Else
        tmpGPFee = 0
    End If
    If IsNumeric(lstInvoice.ListItems(AA).SubItems(84)) Then
        tmpGPLTMed = Format(lstInvoice.ListItems(AA).SubItems(84), "#,##0.00")
    Else
        tmpGPLTMed = 0
    End If
    If IsNumeric(lstInvoice.ListItems(AA).SubItems(85)) Then
        tmpGPInj = Format(lstInvoice.ListItems(AA).SubItems(85), "#,##0.00")
    Else
        tmpGPInj = 0
    End If
    If IsNumeric(lstInvoice.ListItems(AA).SubItems(86)) Then
        tmpGPMed = Format(lstInvoice.ListItems(AA).SubItems(86), "#,##0.00")
    Else
        tmpGPMed = 0
    End If
    If IsNumeric(lstInvoice.ListItems(AA).SubItems(87)) Then
        tmpSPFee = Format(lstInvoice.ListItems(AA).SubItems(87), "#,##0.00")
    Else
        tmpSPFee = 0
    End If
    If IsNumeric(lstInvoice.ListItems(AA).SubItems(88)) Then
        tmpSPLTMed = Format(lstInvoice.ListItems(AA).SubItems(88), "#,##0.00")
    Else
        tmpSPLTMed = 0
    End If
    If IsNumeric(lstInvoice.ListItems(AA).SubItems(89)) Then
        tmpSPInj = Format(lstInvoice.ListItems(AA).SubItems(89), "#,##0.00")
    Else
        tmpSPInj = 0
    End If
    If IsNumeric(lstInvoice.ListItems(AA).SubItems(90)) Then
        tmpSPMed = Format(lstInvoice.ListItems(AA).SubItems(90), "#,##0.00")
    Else
        tmpSPMed = 0
    End If
    'If IsNumeric(lstInvoice.ListItems(AA).SubItems(115)) Then
    '    dblExcess = Format(lstInvoice.ListItems(AA).SubItems(115), "#,##0.00")
    'Else
    '    dblExcess = 0
    'End If
    
    StrSql = ""
    StrSql = "Exec SaveActualOthers '" & tmpCASnumber & "'," & CDbl(tmpGPFee) & "," & CDbl(tmpGPLTMed) & "," & CDbl(tmpGPInj) & "," & _
             CDbl(tmpGPMed) & "," & CDbl(tmpSPFee) & "," & CDbl(tmpSPLTMed) & "," & CDbl(tmpSPInj) & "," & CDbl(tmpSPMed) & ""
    MediConnCISWS.Execute StrSql
End Sub

Private Sub SaveINVActual()
'Dim CISAct As New ADODB.Recordset
Dim SaveINVActualSql As String
    
        SaveINVActualSql = ""
        
        tmpCLMNo = NewCaseNo
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(168)) Then
            tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(168), "#,##0.00")
        Else
            tmpCLMGPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(169)) Then
            tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(169), "#,##0.00")
        Else
            tmpCLMGPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(170)) Then
            tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(170), "#,##0.00")
        Else
            tmpCLMGPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(171)) Then
            tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(171), "#,##0.00")
        Else
            tmpCLMSPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(172)) Then
            tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(172), "#,##0.00")
        Else
            tmpCLMSPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(173)) Then
            tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(173), "#,##0.00")
        Else
            tmpCLMSPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(174)) Then
            tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(174), "#,##0.00")
        Else
            tmpCLMXRay = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(175)) Then
            tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(175), "#,##0.00")
        Else
            tmpCLMUltraSound = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(176)) Then
            tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(176), "#,##0.00")
        Else
            tmpCLMECG = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(177)) Then
            tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(177), "#,##0.00")
        Else
            tmpCLMDressing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(178)) Then
            tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(178), "#,##0.00")
        Else
            tmpCLMNebulizer = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(179)) Then
            tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(179), "#,##0.00")
        Else
            tmpCLMID = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(180)) Then
            tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(180), "#,##0.00")
        Else
            tmpCLMTSSTO = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(181)) Then
            tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(181), "#,##0.00")
        Else
            tmpCLMMinorSurgery = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(182)) Then
            tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(182), "#,##0.00")
        Else
            tmpCLMPOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(183)) Then
            tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(183), "#,##0.00")
        Else
            tmpCLMBloodDrawing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(184)) Then
            tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(184), "#,##0.00")
        Else
            tmpCLMBloodGlucose = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(185)) Then
            tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(185), "#,##0.00")
        Else
            tmpCLMUrine = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(186)) Then
            tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(186), "#,##0.00")
        Else
            tmpCLMPapSmear = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(187)) Then
            tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(187), "#,##0.00")
        Else
            tmpCLMLOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(188)) Then
            tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(188), "#,##0.00")
        Else
            tmpCLMDTPPolio = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(189)) Then
            tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(189), "#,##0.00")
        Else
            tmpCLMDTPPolioHIB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(190)) Then
            tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(190), "#,##0.00")
        Else
            tmpCLMHepatitisB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(191)) Then
            tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(191), "#,##0.00")
        Else
            tmpCLMMMR = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(192)) Then
            tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(192), "#,##0.00")
        Else
            tmpCLMIOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(193)) Then
            tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(193), "#,##0.00")
        Else
            tmpCLMDental = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(194)) Then
            tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(194), "#,##0.00")
        Else
            tmpCLMMaternity = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(195)) Then
            tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(195), "#,##0.00")
        Else
            tmpCLMPRVCheckup = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(196)) Then
            tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(196), "#,##0.00")
        Else
            tmpCLMVision = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(197)) Then
            tmpCLMTyphiad = lstInvoice.ListItems(AA).SubItems(197)
        Else
            tmpCLMTyphiad = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(198)) Then
            tmpCLMOtherNoBalDed = lstInvoice.ListItems(AA).SubItems(198)
        Else
            tmpCLMOtherNoBalDed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(474)) Then
            tmpCLMTyphiad = lstInvoice.ListItems(AA).SubItems(474)
        Else
            tmpCLMTyphiad = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(475)) Then
            tmpCLMOtherNoBalDed = lstInvoice.ListItems(AA).SubItems(475)
        Else
            tmpCLMOtherNoBalDed = 0
        End If


        SaveINVActualSql = "Exec SaveINVActual '" & tmpCLMNo & "'," & tmpCLMGPFeeBasicMed & "," & tmpCLMGPFeeMedInj & "," & _
                            tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & _
                            tmpCLMXRay & "," & tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & tmpCLMNebulizer & "," & _
                            tmpCLMID & "," & tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & _
                            tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & tmpCLMLOthers & "," & _
                            tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & _
                            tmpCLMDental & "," & tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & _
                            tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
         Debug.Print SaveINVActualSql
         
        MediConnCISWS.Execute SaveINVActualSql
'    CISSql = "SELECT * FROM CLMActualNew "
'    CISAct.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
'    CISAct.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))

 '   With CISAct
        '!CLMNo = NewCaseNo
        '!CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(168)
        '!CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(169)
        '!CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(170)
        '!CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(171)
        '!CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(172)
        '!CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(173)
        '!CLMXRay = lstInvoice.ListItems(AA).SubItems(174)
        '!CLMUltraSound = lstInvoice.ListItems(AA).SubItems(175)
        '!CLMECG = lstInvoice.ListItems(AA).SubItems(176)
        '!CLMDressing = lstInvoice.ListItems(AA).SubItems(177)
        '!CLMNebulizer = lstInvoice.ListItems(AA).SubItems(178)
        '!CLMID = lstInvoice.ListItems(AA).SubItems(179)
        '!CLMTSSTO = lstInvoice.ListItems(AA).SubItems(180)
        '!CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(181)
        '!CLMPOthers = lstInvoice.ListItems(AA).SubItems(182)
        '!CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(183)
        '!CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(184)
        '!CLMUrine = lstInvoice.ListItems(AA).SubItems(185)
        '!CLMPapSmear = lstInvoice.ListItems(AA).SubItems(186)
        '!CLMLOthers = lstInvoice.ListItems(AA).SubItems(187)
        '!CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(188)
        '!CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(189)
        '!CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(190)
        '!CLMMMR = lstInvoice.ListItems(AA).SubItems(191)
        '!CLMIOthers = lstInvoice.ListItems(AA).SubItems(192)
        '!CLMDental = lstInvoice.ListItems(AA).SubItems(193)
        '!CLMMaternity = lstInvoice.ListItems(AA).SubItems(194)
        '!CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(195)
        '!CLMVision = lstInvoice.ListItems(AA).SubItems(196)
  '  End' With
  '  CISAct.Update
    
  '  CISAct.Close
  '  Set CISAct = Nothing
End Sub

Private Sub SaveINVCLNContract()
'Dim CISCLNLmt As New ADODB.Recordset
Dim SaveINVCLNContractSql As String
        
        SaveINVCLNContractSql = ""
        
        tmpCASnumber = NewCaseNo
        
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(211)) Then
            tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(211), "#,##0.00")
        Else
            tmpCLMGPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(212)) Then
            tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(212), "#,##0.00")
        Else
            tmpCLMGPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(213)) Then
            tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(213), "#,##0.00")
        Else
            tmpCLMGPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(214)) Then
            tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(214), "#,##0.00")
        Else
            tmpCLMSPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(215)) Then
            tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(215), "#,##0.00")
        Else
            tmpCLMSPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(216)) Then
            tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(216), "#,##0.00")
        Else
            tmpCLMSPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(217)) Then
            tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(217), "#,##0.00")
        Else
            tmpCLMXRay = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(218)) Then
            tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(218), "#,##0.00")
        Else
            tmpCLMUltraSound = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(219)) Then
            tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(219), "#,##0.00")
        Else
            tmpCLMECG = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(220)) Then
            tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(220), "#,##0.00")
        Else
            tmpCLMDressing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(221)) Then
            tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(221), "#,##0.00")
        Else
            tmpCLMNebulizer = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(222)) Then
            tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(222), "#,##0.00")
        Else
            tmpCLMID = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(223)) Then
            tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(223), "#,##0.00")
        Else
            tmpCLMTSSTO = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(224)) Then
            tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(224), "#,##0.00")
        Else
            tmpCLMMinorSurgery = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(225)) Then
            tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(225), "#,##0.00")
        Else
            tmpCLMPOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(226)) Then
            tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(226), "#,##0.00")
        Else
            tmpCLMBloodDrawing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(227)) Then
            tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(227), "#,##0.00")
        Else
            tmpCLMBloodGlucose = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(228)) Then
            tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(228), "#,##0.00")
        Else
            tmpCLMUrine = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(229)) Then
            tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(229), "#,##0.00")
        Else
            tmpCLMPapSmear = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(230)) Then
            tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(230), "#,##0.00")
        Else
            tmpCLMLOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(231)) Then
            tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(231), "#,##0.00")
        Else
            tmpCLMDTPPolio = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(232)) Then
            tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(232), "#,##0.00")
        Else
            tmpCLMDTPPolioHIB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(233)) Then
            tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(233), "#,##0.00")
        Else
            tmpCLMHepatitisB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(234)) Then
            tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(234), "#,##0.00")
        Else
            tmpCLMMMR = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(235)) Then
            tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(235), "#,##0.00")
        Else
            tmpCLMIOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(236)) Then
            tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(236), "#,##0.00")
        Else
            tmpCLMDental = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(237)) Then
            tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(237), "#,##0.00")
        Else
            tmpCLMMaternity = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(238)) Then
            tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(238), "#,##0.00")
        Else
            tmpCLMPRVCheckup = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(239)) Then
            tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(239), "#,##0.00")
        Else
            tmpCLMVision = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(240)) Then
            tmpCLMTyphiad = Format(lstInvoice.ListItems(AA).SubItems(240), "#,##0.00")
        Else
            tmpCLMTyphiad = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(241)) Then
            tmpCLMOtherNoBalDed = Format(lstInvoice.ListItems(AA).SubItems(241), "#,##0.00")
        Else
            tmpCLMOtherNoBalDed = 0
        End If
    
        SaveINVCLNContractSql = "Exec SaveINVCLNContract '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
        tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & tmpCLMSPFeeLTMed & "," & _
        tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & _
        tmpCLMNebulizer & "," & tmpCLMID & "," & tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & _
        tmpCLMBloodDrawing & "," & tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & _
        tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & _
        tmpCLMMMR & "," & tmpCLMIOthers & "," & tmpCLMDental & "," & tmpCLMMaternity & "," & _
        tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
    
    MediConnCISWS.Execute SaveINVCLNContractSql
    'CISSql = "SELECT * FROM CLMClinicLmtNew "
    'CISCLNLmt.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    'CISCLNLmt.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))
    
    'With CISCLNLmt
    '    !CasNumber = NewCaseNo
    '    !CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(211)
    '    !CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(212)
    '    !CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(213)
    '    !CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(214)
    '    !CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(215)
    '    !CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(216)
    '    !CLMXRay = lstInvoice.ListItems(AA).SubItems(217)
    '    !CLMUltraSound = lstInvoice.ListItems(AA).SubItems(218)
    '    !CLMECG = lstInvoice.ListItems(AA).SubItems(219)
    '    !CLMDressing = lstInvoice.ListItems(AA).SubItems(220)
    '    !CLMNebulizer = lstInvoice.ListItems(AA).SubItems(221)
    '   !CLMID = lstInvoice.ListItems(AA).SubItems(222)
    '    !CLMTSSTO = lstInvoice.ListItems(AA).SubItems(223)
    '    !CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(224)
    '    !CLMPOthers = lstInvoice.ListItems(AA).SubItems(225)
    '    !CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(226)
    '    !CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(227)
     '   !CLMUrine = lstInvoice.ListItems(AA).SubItems(228)
    '    !CLMPapSmear = lstInvoice.ListItems(AA).SubItems(229)
    '    !CLMLOthers = lstInvoice.ListItems(AA).SubItems(230)
    '    !CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(231)
    '    !CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(232)
    '    !CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(233)
    '    !CLMMMR = lstInvoice.ListItems(AA).SubItems(234)
    '    !CLMIOthers = lstInvoice.ListItems(AA).SubItems(235)
    '    !CLMDental = lstInvoice.ListItems(AA).SubItems(236)
    '    !CLMMaternity = lstInvoice.ListItems(AA).SubItems(237)
    '    !CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(238)
    '    !CLMVision = lstInvoice.ListItems(AA).SubItems(239)
    'End With
    'CISCLNLmt.Update
    
    'CISCLNLmt.Close
    'Set CISCLNLmt = Nothing
End Sub

Private Sub SaveINVMBMApp()
'Dim CISAPP As New ADODB.Recordset
Dim SaveINVMBMAppSql As String
    
        SaveINVMBMAppSql = ""
        tmpCASnumber = NewCaseNo
        
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(254)) Then
            tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(254), "#,##0.00")
        Else
            tmpCLMGPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(255)) Then
            tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(255), "#,##0.00")
        Else
            tmpCLMGPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(256)) Then
            tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(256), "#,##0.00")
        Else
            tmpCLMGPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(257)) Then
            tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(257), "#,##0.00")
        Else
            tmpCLMSPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(258)) Then
            tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(258), "#,##0.00")
        Else
            tmpCLMSPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(259)) Then
            tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(259), "#,##0.00")
        Else
            tmpCLMSPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(260)) Then
            tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(260), "#,##0.00")
        Else
            tmpCLMXRay = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(261)) Then
            tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(261), "#,##0.00")
        Else
            tmpCLMUltraSound = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(262)) Then
            tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(262), "#,##0.00")
        Else
            tmpCLMECG = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(263)) Then
            tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(263), "#,##0.00")
        Else
            tmpCLMDressing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(264)) Then
            tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(264), "#,##0.00")
        Else
            tmpCLMNebulizer = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(265)) Then
            tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(265), "#,##0.00")
        Else
            tmpCLMID = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(266)) Then
            tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(266), "#,##0.00")
        Else
            tmpCLMTSSTO = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(267)) Then
            tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(267), "#,##0.00")
        Else
            tmpCLMMinorSurgery = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(268)) Then
            tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(268), "#,##0.00")
        Else
            tmpCLMPOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(269)) Then
            tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(269), "#,##0.00")
        Else
            tmpCLMBloodDrawing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(270)) Then
            tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(270), "#,##0.00")
        Else
            tmpCLMBloodGlucose = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(271)) Then
            tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(271), "#,##0.00")
        Else
            tmpCLMUrine = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(272)) Then
            tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(272), "#,##0.00")
        Else
            tmpCLMPapSmear = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(273)) Then
            tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(273), "#,##0.00")
        Else
            tmpCLMLOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(274)) Then
            tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(274), "#,##0.00")
        Else
            tmpCLMDTPPolio = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(275)) Then
            tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(275), "#,##0.00")
        Else
            tmpCLMDTPPolioHIB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(276)) Then
            tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(276), "#,##0.00")
        Else
            tmpCLMHepatitisB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(277)) Then
            tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(277), "#,##0.00")
        Else
            tmpCLMMMR = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(278)) Then
            tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(278), "#,##0.00")
        Else
            tmpCLMIOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(279)) Then
            tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(279), "#,##0.00")
        Else
            tmpCLMDental = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(280)) Then
            tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(280), "#,##0.00")
        Else
            tmpCLMMaternity = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(281)) Then
            tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(281), "#,##0.00")
        Else
            tmpCLMPRVCheckup = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(282)) Then
            tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(282), "#,##0.00")
        Else
            tmpCLMVision = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(283)) Then
            tmpCLMTyphiad = Format(lstInvoice.ListItems(AA).SubItems(283), "#,##0.00")
        Else
            tmpCLMTyphiad = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(284)) Then
            tmpCLMOtherNoBalDed = Format(lstInvoice.ListItems(AA).SubItems(284), "#,##0.00")
        Else
            tmpCLMOtherNoBalDed = 0
        End If

        SaveINVMBMAppSql = "Exec SaveINVMBMApp '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
                            tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & _
                            tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & _
                            tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & tmpCLMNebulizer & "," & _
                            tmpCLMID & "," & tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & _
                            tmpCLMBloodDrawing & "," & tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & _
                            tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & _
                            tmpCLMMMR & "," & tmpCLMIOthers & "," & tmpCLMDental & "," & tmpCLMMaternity & "," & _
                            tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
        MediConnCISWS.Execute SaveINVMBMAppSql
    'CISSql = "SELECT * FROM CLMAppNew "
    'CISAPP.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    'CISAPP.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))

    'With CISAPP
        '!CasNumber = NewCaseNo
        '!CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(254)
        '!CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(255)
        '!CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(256)
        '!CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(257)
        '!CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(258)
        '!CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(259)
        '!CLMXRay = lstInvoice.ListItems(AA).SubItems(260)
        '!CLMUltraSound = lstInvoice.ListItems(AA).SubItems(261)
        '!CLMECG = lstInvoice.ListItems(AA).SubItems(262)
        '!CLMDressing = lstInvoice.ListItems(AA).SubItems(263)
        '!CLMNebulizer = lstInvoice.ListItems(AA).SubItems(264)
        '!CLMID = lstInvoice.ListItems(AA).SubItems(265)
        '!CLMTSSTO = lstInvoice.ListItems(AA).SubItems(266)
        '!CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(267)
        '!CLMPOthers = lstInvoice.ListItems(AA).SubItems(268)
        '!CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(269)
        '!CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(270)
        '!CLMUrine = lstInvoice.ListItems(AA).SubItems(271)
        '!CLMPapSmear = lstInvoice.ListItems(AA).SubItems(272)
        '!CLMLOthers = lstInvoice.ListItems(AA).SubItems(273)
        '!CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(274)
        '!CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(275)
        '!CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(276)
        '!CLMMMR = lstInvoice.ListItems(AA).SubItems(277)
        '!CLMIOthers = lstInvoice.ListItems(AA).SubItems(278)
        '!CLMDental = lstInvoice.ListItems(AA).SubItems(279)
        '!CLMMaternity = lstInvoice.ListItems(AA).SubItems(280)
        '!CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(281)
        '!CLMVision = lstInvoice.ListItems(AA).SubItems(282)
    'End With
    'CISAPP.Update
    
    'CISAPP.Close
    'Set CISAPP = Nothing
End Sub

Private Sub SaveINVMBMLmt()
'Dim CISLmt As New ADODB.Recordset
Dim SaveINVMBMLmtSql As String
    
        SaveINVMBMLmtSql = ""
        tmpCASnumber = NewCaseNo
        
        
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(297)) Then
            tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(297), "#,##0.00")
        Else
            tmpCLMGPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(298)) Then
            tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(298), "#,##0.00")
        Else
            tmpCLMGPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(299)) Then
            tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(299), "#,##0.00")
        Else
            tmpCLMGPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(300)) Then
            tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(300), "#,##0.00")
        Else
            tmpCLMSPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(301)) Then
            tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(301), "#,##0.00")
        Else
            tmpCLMSPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(302)) Then
            tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(302), "#,##0.00")
        Else
            tmpCLMSPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(303)) Then
            tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(303), "#,##0.00")
        Else
            tmpCLMXRay = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(304)) Then
            tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(304), "#,##0.00")
        Else
            tmpCLMUltraSound = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(305)) Then
            tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(305), "#,##0.00")
        Else
            tmpCLMECG = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(306)) Then
            tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(306), "#,##0.00")
        Else
            tmpCLMDressing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(307)) Then
            tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(307), "#,##0.00")
        Else
            tmpCLMNebulizer = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(308)) Then
            tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(308), "#,##0.00")
        Else
            tmpCLMID = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(309)) Then
            tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(309), "#,##0.00")
        Else
            tmpCLMTSSTO = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(310)) Then
            tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(310), "#,##0.00")
        Else
            tmpCLMMinorSurgery = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(311)) Then
            tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(311), "#,##0.00")
        Else
            tmpCLMPOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(312)) Then
            tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(312), "#,##0.00")
        Else
            tmpCLMBloodDrawing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(313)) Then
            tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(313), "#,##0.00")
        Else
            tmpCLMBloodGlucose = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(314)) Then
            tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(314), "#,##0.00")
        Else
            tmpCLMUrine = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(315)) Then
            tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(315), "#,##0.00")
        Else
            tmpCLMPapSmear = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(316)) Then
            tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(316), "#,##0.00")
        Else
            tmpCLMLOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(317)) Then
            tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(317), "#,##0.00")
        Else
            tmpCLMDTPPolio = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(318)) Then
            tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(318), "#,##0.00")
        Else
            tmpCLMDTPPolioHIB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(319)) Then
            tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(319), "#,##0.00")
        Else
            tmpCLMHepatitisB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(320)) Then
            tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(320), "#,##0.00")
        Else
            tmpCLMMMR = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(321)) Then
            tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(321), "#,##0.00")
        Else
            tmpCLMIOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(322)) Then
            tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(322), "#,##0.00")
        Else
            tmpCLMDental = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(323)) Then
            tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(323), "#,##0.00")
        Else
            tmpCLMMaternity = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(324)) Then
            tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(324), "#,##0.00")
        Else
            tmpCLMPRVCheckup = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(325)) Then
            tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(325), "#,##0.00")
        Else
            tmpCLMVision = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(326)) Then
            tmpCLMTyphiad = Format(lstInvoice.ListItems(AA).SubItems(326), "#,##0.00")
        Else
            tmpCLMTyphiad = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(327)) Then
            tmpCLMOtherNoBalDed = Format(lstInvoice.ListItems(AA).SubItems(327), "#,##0.00")
        Else
            tmpCLMOtherNoBalDed = 0
        End If

        SaveINVMBMLmtSql = "Exec SaveINVMBMLmt '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
                            tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & _
                            tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & _
                            tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & _
                            tmpCLMNebulizer & "," & tmpCLMID & "," & tmpCLMTSSTO & "," & _
                            tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & _
                            tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & _
                            tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & _
                            tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & _
                            tmpCLMDental & "," & tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & _
                            tmpCLMVision & ", " & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
        
        MediConnCISWS.Execute SaveINVMBMLmtSql
    'CISSql = "SELECT * FROM CLMLimitNew "
    'CISLmt.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    'CISLmt.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))

    'With CISLmt
        '!CasNumber = NewCaseNo
        '!CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(297)
        '!CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(298)
        '!CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(299)
        '!CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(300)
        '!CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(301)
        '!CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(302)
        '!CLMXRay = lstInvoice.ListItems(AA).SubItems(303)
        '!CLMUltraSound = lstInvoice.ListItems(AA).SubItems(304)
        '!CLMECG = lstInvoice.ListItems(AA).SubItems(305)
        '!CLMDressing = lstInvoice.ListItems(AA).SubItems(306)
        '!CLMNebulizer = lstInvoice.ListItems(AA).SubItems(307)
        '!CLMID = lstInvoice.ListItems(AA).SubItems(308)
        '!CLMTSSTO = lstInvoice.ListItems(AA).SubItems(309)
        '!CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(310)
        '!CLMPOthers = lstInvoice.ListItems(AA).SubItems(311)
        '!CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(312)
        '!CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(313)
        '!CLMUrine = lstInvoice.ListItems(AA).SubItems(314)
        '!CLMPapSmear = lstInvoice.ListItems(AA).SubItems(315)
        '!CLMLOthers = lstInvoice.ListItems(AA).SubItems(316)
        '!CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(317)
        '!CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(318)
        '!CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(319)
        '!CLMMMR = lstInvoice.ListItems(AA).SubItems(320)
        '!CLMIOthers = lstInvoice.ListItems(AA).SubItems(321)
        '!CLMDental = lstInvoice.ListItems(AA).SubItems(322)
        '!CLMMaternity = lstInvoice.ListItems(AA).SubItems(323)
        '!CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(324)
        '!CLMVision = lstInvoice.ListItems(AA).SubItems(325)
    'End With
    'CISLmt.Update
    
    'CISLmt.Close
    'Set CISLmt = Nothing
End Sub

Private Sub SaveINVMBMExcess()
'Dim CISExc As New ADODB.Recordset
Dim SaveINVMBMExcessSql As String

        SaveINVMBMExcessSql = ""
        tmpCASnumber = NewCaseNo
        
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(340)) Then
            tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(340), "#,##0.00")
        Else
            tmpCLMGPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(341)) Then
            tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(341), "#,##0.00")
        Else
            tmpCLMGPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(342)) Then
            tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(342), "#,##0.00")
        Else
            tmpCLMGPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(343)) Then
            tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(343), "#,##0.00")
        Else
            tmpCLMSPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(344)) Then
            tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(344), "#,##0.00")
        Else
            tmpCLMSPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(345)) Then
            tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(345), "#,##0.00")
        Else
            tmpCLMSPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(346)) Then
            tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(346), "#,##0.00")
        Else
            tmpCLMXRay = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(347)) Then
            tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(347), "#,##0.00")
        Else
            tmpCLMUltraSound = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(348)) Then
            tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(348), "#,##0.00")
        Else
            tmpCLMECG = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(349)) Then
            tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(349), "#,##0.00")
        Else
            tmpCLMDressing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(350)) Then
            tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(350), "#,##0.00")
        Else
            tmpCLMNebulizer = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(351)) Then
            tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(351), "#,##0.00")
        Else
            tmpCLMID = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(352)) Then
            tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(352), "#,##0.00")
        Else
            tmpCLMTSSTO = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(353)) Then
            tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(353), "#,##0.00")
        Else
            tmpCLMMinorSurgery = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(354)) Then
            tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(354), "#,##0.00")
        Else
            tmpCLMPOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(355)) Then
            tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(355), "#,##0.00")
        Else
            tmpCLMBloodDrawing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(356)) Then
            tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(356), "#,##0.00")
        Else
            tmpCLMBloodGlucose = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(357)) Then
            tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(357), "#,##0.00")
        Else
            tmpCLMUrine = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(358)) Then
            tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(358), "#,##0.00")
        Else
            tmpCLMPapSmear = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(359)) Then
            tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(359), "#,##0.00")
        Else
            tmpCLMLOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(360)) Then
            tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(360), "#,##0.00")
        Else
            tmpCLMDTPPolio = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(361)) Then
            tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(361), "#,##0.00")
        Else
            tmpCLMDTPPolioHIB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(362)) Then
            tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(362), "#,##0.00")
        Else
            tmpCLMHepatitisB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(363)) Then
            tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(363), "#,##0.00")
        Else
            tmpCLMMMR = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(364)) Then
            tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(364), "#,##0.00")
        Else
            tmpCLMIOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(365)) Then
            tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(365), "#,##0.00")
        Else
            tmpCLMDental = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(366)) Then
            tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(366), "#,##0.00")
        Else
            tmpCLMMaternity = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(367)) Then
            tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(367), "#,##0.00")
        Else
            tmpCLMPRVCheckup = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(368)) Then
            tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(368), "#,##0.00")
        Else
            tmpCLMVision = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(369)) Then
            tmpCLMTyphiad = Format(lstInvoice.ListItems(AA).SubItems(369), "#,##0.00")
        Else
            tmpCLMTyphiad = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(370)) Then
            tmpCLMOtherNoBalDed = Format(lstInvoice.ListItems(AA).SubItems(370), "#,##0.00")
        Else
            tmpCLMOtherNoBalDed = 0
        End If


        SaveINVMBMExcessSql = "Exec SaveINVMBMExcess '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
                                tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & _
                                tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & tmpCLMUltraSound & "," & _
                                tmpCLMECG & "," & tmpCLMDressing & "," & tmpCLMNebulizer & "," & tmpCLMID & "," & _
                                tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & _
                                tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & _
                                tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & _
                                tmpCLMDental & "," & tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & _
                                tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
    
    MediConnCISWS.Execute SaveINVMBMExcessSql
   
    
   ' CISSql = "SELECT * FROM CLMMemberNew "
   ' CISExc.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
   ' CISExc.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))
   ' With CISExc
   '     !CasNumber = NewCaseNo
    '    !CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(340)
    '    !CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(341)
    '    !CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(342)
    '    !CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(343)
    '    !CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(344)
    '    !CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(345)
    '    !CLMXRay = lstInvoice.ListItems(AA).SubItems(346)
    '    !CLMUltraSound = lstInvoice.ListItems(AA).SubItems(347)
    '    !CLMECG = lstInvoice.ListItems(AA).SubItems(348)
    '    !CLMDressing = lstInvoice.ListItems(AA).SubItems(349)
    '    !CLMNebulizer = lstInvoice.ListItems(AA).SubItems(350)
    '    !CLMID = lstInvoice.ListItems(AA).SubItems(351)
    '    !CLMTSSTO = lstInvoice.ListItems(AA).SubItems(352)
    '    !CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(353)
    '    !CLMPOthers = lstInvoice.ListItems(AA).SubItems(354)
    '    !CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(355)
    '    !CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(356)
    '    !CLMUrine = lstInvoice.ListItems(AA).SubItems(357)
    '    !CLMPapSmear = lstInvoice.ListItems(AA).SubItems(358)
    '    !CLMLOthers = lstInvoice.ListItems(AA).SubItems(359)
    '    !CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(360)
    '    !CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(361)
    '    !CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(362)
    '    !CLMMMR = lstInvoice.ListItems(AA).SubItems(363)
    '    !CLMIOthers = lstInvoice.ListItems(AA).SubItems(364)
    '    !CLMDental = lstInvoice.ListItems(AA).SubItems(365)
    '    !CLMMaternity = lstInvoice.ListItems(AA).SubItems(366)
    '    !CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(367)
    '    !CLMVision = lstInvoice.ListItems(AA).SubItems(368)
    'End With
    'CISExc.Update
    
   ' CISExc.Close
    'Set CISExc = Nothing
End Sub

Private Sub SaveINVCLNExcess()
'Dim CISExc As New ADODB.Recordset
Dim SaveINVCLNExcessSql As String
        
        SaveINVCLNExcessSql = ""
        
        tmpCASnumber = NewCaseNo
        
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(383)) Then
            tmpCLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(383)
        Else
            tmpCLMGPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(384)) Then
            tmpCLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(384)
        Else
            tmpCLMGPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(385)) Then
            tmpCLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(385)
        Else
            tmpCLMGPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(386)) Then
            tmpCLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(386)
        Else
            tmpCLMSPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(387)) Then
            tmpCLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(387)
        Else
            tmpCLMSPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(388)) Then
            tmpCLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(388)
        Else
            tmpCLMSPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(389)) Then
            tmpCLMXRay = lstInvoice.ListItems(AA).SubItems(389)
        Else
            tmpCLMXRay = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(390)) Then
            tmpCLMUltraSound = lstInvoice.ListItems(AA).SubItems(390)
        Else
            tmpCLMUltraSound = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(391)) Then
            tmpCLMECG = lstInvoice.ListItems(AA).SubItems(391)
        Else
            tmpCLMECG = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(392)) Then
            tmpCLMDressing = lstInvoice.ListItems(AA).SubItems(392)
        Else
            tmpCLMDressing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(393)) Then
            tmpCLMNebulizer = lstInvoice.ListItems(AA).SubItems(393)
        Else
            tmpCLMNebulizer = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(394)) Then
            tmpCLMID = lstInvoice.ListItems(AA).SubItems(394)
        Else
            tmpCLMID = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(395)) Then
            tmpCLMTSSTO = lstInvoice.ListItems(AA).SubItems(395)
        Else
            tmpCLMTSSTO = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(396)) Then
            tmpCLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(396)
        Else
            tmpCLMMinorSurgery = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(397)) Then
            tmpCLMPOthers = lstInvoice.ListItems(AA).SubItems(397)
        Else
            tmpCLMPOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(398)) Then
            tmpCLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(398)
        Else
            tmpCLMBloodDrawing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(399)) Then
            tmpCLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(399)
        Else
            tmpCLMBloodGlucose = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(400)) Then
            tmpCLMUrine = lstInvoice.ListItems(AA).SubItems(400)
        Else
            tmpCLMUrine = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(401)) Then
            tmpCLMPapSmear = lstInvoice.ListItems(AA).SubItems(401)
        Else
            tmpCLMPapSmear = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(402)) Then
            tmpCLMLOthers = lstInvoice.ListItems(AA).SubItems(402)
        Else
            tmpCLMLOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(403)) Then
            tmpCLMDTPPolio = lstInvoice.ListItems(AA).SubItems(403)
        Else
            tmpCLMDTPPolio = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(404)) Then
            tmpCLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(404)
        Else
            tmpCLMDTPPolioHIB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(405)) Then
            tmpCLMHepatitisB = lstInvoice.ListItems(AA).SubItems(405)
        Else
            tmpCLMHepatitisB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(406)) Then
            tmpCLMMMR = lstInvoice.ListItems(AA).SubItems(406)
        Else
            tmpCLMMMR = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(407)) Then
            tmpCLMIOthers = lstInvoice.ListItems(AA).SubItems(407)
        Else
            tmpCLMIOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(408)) Then
            tmpCLMDental = lstInvoice.ListItems(AA).SubItems(408)
        Else
            tmpCLMDental = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(409)) Then
            tmpCLMMaternity = lstInvoice.ListItems(AA).SubItems(409)
        Else
            tmpCLMMaternity = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(410)) Then
            tmpCLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(410)
        Else
            tmpCLMPRVCheckup = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(411)) Then
            tmpCLMVision = lstInvoice.ListItems(AA).SubItems(411)
        Else
            tmpCLMVision = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(412)) Then
            tmpCLMTyphiad = lstInvoice.ListItems(AA).SubItems(412)
        Else
            tmpCLMTyphiad = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(413)) Then
            tmpCLMOtherNoBalDed = lstInvoice.ListItems(AA).SubItems(413)
        Else
            tmpCLMOtherNoBalDed = 0
        End If

    SaveINVCLNExcessSql = "Exec SaveINVCLNExcess '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
                        tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & tmpCLMSPFeeLTMed & "," & _
                        tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & _
                        tmpCLMNebulizer & "," & tmpCLMID & "," & tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & _
                        tmpCLMBloodDrawing & "," & tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & tmpCLMLOthers & "," & _
                        tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & _
                        tmpCLMDental & "," & tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & _
                        tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
    MediConnCISWS.Execute SaveINVCLNExcessSql
    'CISSql = "SELECT * FROM CLMClinicNew "
    'CISExc.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    'CISExc.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))

    'With CISExc
        '!CasNumber = NewCaseNo
        '!CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(383)
        '!CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(384)
        '!CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(385)
        '!CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(386)
        '!CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(387)
        '!CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(388)
        '!CLMXRay = lstInvoice.ListItems(AA).SubItems(389)
        '!CLMUltraSound = lstInvoice.ListItems(AA).SubItems(390)
        '!CLMECG = lstInvoice.ListItems(AA).SubItems(391)
        '!CLMDressing = lstInvoice.ListItems(AA).SubItems(392)
        '!CLMNebulizer = lstInvoice.ListItems(AA).SubItems(393)
        '!CLMID = lstInvoice.ListItems(AA).SubItems(394)
        '!CLMTSSTO = lstInvoice.ListItems(AA).SubItems(395)
        '!CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(396)
        '!CLMPOthers = lstInvoice.ListItems(AA).SubItems(397)
        '!CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(398)
        '!CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(399)
        '!CLMUrine = lstInvoice.ListItems(AA).SubItems(400)
        '!CLMPapSmear = lstInvoice.ListItems(AA).SubItems(401)
        '!CLMLOthers = lstInvoice.ListItems(AA).SubItems(402)
        '!CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(403)
        '!CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(404)
        '!CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(405)
        '!CLMMMR = lstInvoice.ListItems(AA).SubItems(406)
        '!CLMIOthers = lstInvoice.ListItems(AA).SubItems(407)
        '!CLMDental = lstInvoice.ListItems(AA).SubItems(408)
        '!CLMMaternity = lstInvoice.ListItems(AA).SubItems(409)
        '!CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(410)
        '!CLMVision = lstInvoice.ListItems(AA).SubItems(411)
    'End With
    'CISExc.Update
    
    'CISExc.Close
    'Set CISExc = Nothing
End Sub


Private Sub SaveINVCLNApp()
'Dim CISExc As New ADODB.Recordset
Dim SaveINVCLNAppSql As String
        
        SaveINVCLNAppSql = ""
        tmpCASnumber = NewCaseNo
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(426)) Then
            tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(426), "#,##0.00")
        Else
            tmpCLMGPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(427)) Then
            tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(427), "#,##0.00")
        Else
            tmpCLMGPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(428)) Then
            tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(428), "#,##0.00")
        Else
            tmpCLMGPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(429)) Then
            tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(429), "#,##0.00")
        Else
            tmpCLMSPFeeBasicMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(430)) Then
            tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(430), "#,##0.00")
        Else
            tmpCLMSPFeeLTMed = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(431)) Then
            tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(431), "#,##0.00")
        Else
            tmpCLMSPFeeMedInj = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(432)) Then
            tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(432), "#,##0.00")
        Else
            tmpCLMXRay = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(433)) Then
            tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(433), "#,##0.00")
        Else
            tmpCLMUltraSound = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(434)) Then
            tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(434), "#,##0.00")
        Else
            tmpCLMECG = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(435)) Then
            tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(435), "#,##0.00")
        Else
            tmpCLMDressing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(436)) Then
            tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(436), "#,##0.00")
        Else
            tmpCLMNebulizer = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(437)) Then
            tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(437), "#,##0.00")
        Else
            tmpCLMID = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(438)) Then
            tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(438), "#,##0.00")
        Else
            tmpCLMTSSTO = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(439)) Then
            tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(439), "#,##0.00")
        Else
            tmpCLMMinorSurgery = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(440)) Then
            tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(440), "#,##0.00")
        Else
            tmpCLMPOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(441)) Then
            tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(441), "#,##0.00")
        Else
            tmpCLMBloodDrawing = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(442)) Then
            tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(442), "#,##0.00")
        Else
            tmpCLMBloodGlucose = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(443)) Then
            tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(443), "#,##0.00")
        Else
            tmpCLMUrine = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(444)) Then
            tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(444), "#,##0.00")
        Else
            tmpCLMPapSmear = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(445)) Then
            tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(445), "#,##0.00")
        Else
            tmpCLMLOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(446)) Then
            tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(446), "#,##0.00")
        Else
            tmpCLMDTPPolio = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(447)) Then
            tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(447), "#,##0.00")
        Else
            tmpCLMDTPPolioHIB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(448)) Then
            tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(448), "#,##0.00")
        Else
            tmpCLMHepatitisB = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(449)) Then
            tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(449), "#,##0.00")
        Else
            tmpCLMMMR = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(450)) Then
            tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(450), "#,##0.00")
        Else
            tmpCLMIOthers = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(451)) Then
            tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(451), "#,##0.00")
        Else
            tmpCLMDental = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(452)) Then
            tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(452), "#,##0.00")
        Else
            tmpCLMMaternity = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(453)) Then
            tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(453), "#,##0.00")
        Else
            tmpCLMPRVCheckup = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(454)) Then
            tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(454), "#,##0.00")
        Else
            tmpCLMVision = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(455)) Then
            tmpCLMTyphiad = Format(lstInvoice.ListItems(AA).SubItems(455), "#,##0.00")
        Else
            tmpCLMTyphiad = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(456)) Then
            tmpCLMOtherNoBalDed = Format(lstInvoice.ListItems(AA).SubItems(456), "#,##0.00")
        Else
            tmpCLMOtherNoBalDed = 0
        End If

    SaveINVCLNAppSql = "Exec SaveINVCLNApp '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
                        tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & tmpCLMSPFeeLTMed & "," & _
                        tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & tmpCLMUltraSound & "," & tmpCLMECG & "," & _
                        tmpCLMDressing & "," & tmpCLMNebulizer & "," & tmpCLMID & "," & tmpCLMTSSTO & "," & _
                        tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & tmpCLMBloodGlucose & "," & _
                        tmpCLMUrine & "," & tmpCLMPapSmear & "," & tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & _
                        tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & _
                        tmpCLMDental & "," & tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & _
                        tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
                        
    MediConnCISWS.Execute SaveINVCLNAppSql
    'CISSql = "SELECT * FROM CLMClinicAppNew "
    
    'CISExc.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    'CISExc.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))

    'With CISExc
        '!CasNumber = NewCaseNo
        '!CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(426)
        '!CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(427)
        '!CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(428)
        '!CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(429)
        '!CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(430)
        '!CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(431)
        '!CLMXRay = lstInvoice.ListItems(AA).SubItems(432)
        '!CLMUltraSound = lstInvoice.ListItems(AA).SubItems(433)
        '!CLMECG = lstInvoice.ListItems(AA).SubItems(434)
        '!CLMDressing = lstInvoice.ListItems(AA).SubItems(435)
        '!CLMNebulizer = lstInvoice.ListItems(AA).SubItems(436)
        '!CLMID = lstInvoice.ListItems(AA).SubItems(437)
        '!CLMTSSTO = lstInvoice.ListItems(AA).SubItems(438)
        '!CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(439)
        '!CLMPOthers = lstInvoice.ListItems(AA).SubItems(440)
        '!CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(441)
        '!CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(442)
        '!CLMUrine = lstInvoice.ListItems(AA).SubItems(443)
        '!CLMPapSmear = lstInvoice.ListItems(AA).SubItems(444)
        '!CLMLOthers = lstInvoice.ListItems(AA).SubItems(445)
        '!CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(446)
        '!CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(447)
        '!CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(448)
        '!CLMMMR = lstInvoice.ListItems(AA).SubItems(449)
        '!CLMIOthers = lstInvoice.ListItems(AA).SubItems(450)
        '!CLMDental = lstInvoice.ListItems(AA).SubItems(451)
        '!CLMMaternity = lstInvoice.ListItems(AA).SubItems(452)
        '!CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(453)
        '!CLMVision = lstInvoice.ListItems(AA).SubItems(454)
    'End With
    'CISExc.Update
    
    'CISExc.Close
    'Set CISExc = Nothing
    
    
End Sub

Private Sub SaveCharges()
Dim MedType1, MedType2, MedType3, MedType4, MedType5, MedType6, rate1, rate2, rate3, rate4, rate5, rate6, MedType7, rate7 As String
Dim SQLCharges As String
SQLCharges = ""
tmpCASnumber = NewCaseNo

MedType1 = IIf(Len(cboMedType1(0)), cboMedType1(0), "")
rate1 = IIf(Len(txtrate(0)), txtrate(0), "")
MedType2 = IIf(Len(cboMedType1(1)), cboMedType1(1), "")
rate2 = IIf(Len(txtrate(1)), txtrate(1), "")
MedType3 = IIf(Len(cboMedType1(2)), cboMedType1(2), "")
rate3 = IIf(Len(txtrate(2)), txtrate(2), "")
MedType4 = IIf(Len(cboMedType1(3)), cboMedType1(3), "")
rate4 = IIf(Len(txtrate(3)), txtrate(3), "")
MedType5 = IIf(Len(cboMedType1(4)), cboMedType1(4), "")
rate5 = IIf(Len(txtrate(4)), txtrate(4), "")
MedType6 = IIf(Len(cboMedType1(5)), cboMedType1(5), "")
rate6 = IIf(Len(txtrate(5)), txtrate(5), "")
MedType7 = IIf(Len(cboMedType1(6)), cboMedType1(6), "")
rate7 = IIf(Len(txtrate(6)), txtrate(6), "")

SQLCharges = "EXEC Save_Charges '" & tmpCASnumber & "','" & MedType1 & "','" & rate1 & "','" & MedType2 & "','" & rate2 & "','" & MedType3 & "','" & rate3 & "','" & MedType4 & "','" & rate4 & "','" & MedType5 & "','" & rate5 & "','" & MedType6 & "','" & rate6 & "','" & MedType7 & "','" & rate7 & "'"
MediConnCISWS.Execute SQLCharges
End Sub
Private Sub SaveINVClaim()
Dim SaveINVClaimSql As String
Dim tmpCISExc As New ADODB.Recordset
Dim StrSql As String
Dim tmpMBMCostCenter, tmpMBMDeptCode, tmpsubsidiary, tmpGRPCoCode, tmpMBMDepartment As String
Dim tmpMBMNum As String
    Dim tmpMBMArr() As String
    tmpMBMNum = Trim(lstInvoice.ListItems(AA).SubItems(2))
    Dim Pos As Integer
    Pos = 0
    Pos = InStr(tmpMBMNum, "-")
    If Pos > 0 Then
        tmpMBMArr() = Split(tmpMBMNum, "-")
        tmpMBMNum = tmpMBMArr(0)
    Else
        tmpMBMNum = (tmpMBMNum)
    End If


        
        StrSql = " Select ltrim(rtrim(coalesce(MBMOthers.MBMDepartment,''))) AS MBMDepartment,MBMCostCenter,ltrim(rtrim(coalesce(MBMDeptCode,''))) AS MBMDeptCode,ltrim(rtrim(coalesce(subsidiary,''))) AS subsidiary," & _
        " ltrim(rtrim(GRPCompany.GRPCoCode)) AS GRPCoCode from hisdb..MBMOthers  inner join cisdb..GRPCompany on GRPCompany.GRPCoName=MBMOthers.mbmgroupcompany " & _
        " Where MBMNUmber ='" & tmpMBMNum & "'"
            tmpCISExc.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                tmpMBMDepartment = IIf(Trim(Len(tmpCISExc!MBMDepartment)), tmpCISExc!MBMDepartment, "NA")
                tmpMBMCostCenter = IIf(Trim(Len(tmpCISExc!MBMCostCenter)), tmpCISExc!MBMCostCenter, "NA")
                tmpMBMDeptCode = IIf(Trim(Len(tmpCISExc!MBMDeptCode)), tmpCISExc!MBMDeptCode, "NA")
                tmpsubsidiary = IIf(Trim(Len(tmpCISExc!Subsidiary)), tmpCISExc!Subsidiary, "NA")
                tmpGRPCoCode = IIf(Trim(Len(tmpCISExc!GRPCoCode)), tmpCISExc!GRPCoCode, "")
        tmpCISExc.Close
        Set tmpCISExc = Nothing
        
        
        
        tmpLimitExceed = 0
        SaveINVClaimSql = ""
        tmpCASnumber = NewCaseNo
        tmpClaimOpenDate = Format(Date, "yyyy/mm/dd")
        tmpClaimOpenBy = Logon
        tmpClaimCloseDate = Format(Date, "yyyy/mm/dd")
        tmpClaimClosedBy = Logon
        tmpClaimLastUpdateDate = Format(Date, "yyyy/mm/dd")
        tmpClaimLastUpdateUser = Logon
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(43)) Then
            tmpCLMCoPayColl = Format(lstInvoice.ListItems(AA).SubItems(43), "#,##0.00")
        Else
            tmpCLMCoPayColl = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(159)) Then
            tmpCLMCLNExcess = Format(lstInvoice.ListItems(AA).SubItems(159), "#,##0.00")
        Else
            tmpCLMCLNExcess = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(160)) Then
            If lstInvoice.ListItems(AA).SubItems(160) = 0 Then
                tmpCLMTotalActual = lblCLNBillAmt
            Else
                tmpCLMTotalActual = Format(lstInvoice.ListItems(AA).SubItems(160), "#,##0.00")
            End If
        Else
            tmpCLMTotalActual = lblCLNBillAmt
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(161)) Then
            If lstInvoice.ListItems(AA).SubItems(161) <= 0 Then
                tmpCLMFinalApprove = tmpCLMTotalActual
            Else
                tmpCLMFinalApprove = lstInvoice.ListItems(AA).SubItems(161)
            End If
        Else
            tmpCLMFinalApprove = tmpCLMTotalActual
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(162)) Then
            tmpCLMExcess = Format(lstInvoice.ListItems(AA).SubItems(162), "#,##0.00")
        Else
            tmpCLMExcess = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(163)) Then
            tmpCLMBalanceLimit = Format(lstInvoice.ListItems(AA).SubItems(163), "#,##0.00")
        Else
            tmpCLMBalanceLimit = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(164)) Then
            tmpCLMPreviousBalance = Format(lstInvoice.ListItems(AA).SubItems(164), "#,##0.00")
        Else
            tmpCLMPreviousBalance = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(165)) Then
            tmpCLMPaidBYMedix2M = Format(lstInvoice.ListItems(AA).SubItems(165), "#,##0.00")
        Else
            tmpCLMPaidBYMedix2M = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(166)) Then
            tmpCLMPaidBYMedix2C = Format(lstInvoice.ListItems(AA).SubItems(166), "#,##0.00")
        Else
            tmpCLMPaidBYMedix2C = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(113)) Then
            tmpCLMAdjAnnLmt = Format(lstInvoice.ListItems(AA).SubItems(113), "#,##0.00")
        Else
            tmpCLMAdjAnnLmt = 0
        End If
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(114)) Then
            tmpClmNonGPFee = Format(lstInvoice.ListItems(AA).SubItems(114), "#,##0.00")
        Else
            tmpClmNonGPFee = 0
        End If
        tmpNonGPFee = Format(IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(114)), lstInvoice.ListItems(AA).SubItems(114), 0), "#,##0.00")
        tmpCLMMCOProcStatus = IIf(Len(lstInvoice.ListItems(AA).SubItems(469)), lstInvoice.ListItems(AA).SubItems(469), "")
                 
        If tmpCLMMCOProcStatus = "Y" Then
        'get FFS percentage based on clinc code
            Dim sqlstr As String
            Dim FFSRs As New ADODB.Recordset
            Dim FFSpercentage As Double
            FFSpercentage = 0
            sqlstr = "Select CLNMCOProcfeePerc  from CLN where CLNCode ='" + TxtClinicCode + "'"
            FFSRs.Open sqlstr, medixmasterconn, adOpenDynamic, adLockOptimistic
            Do Until FFSRs.EOF
                If IsNumeric(FFSRs!CLNMCOProcfeePerc) Then
                    FFSpercentage = FFSRs!CLNMCOProcfeePerc
                End If
                FFSRs.MoveNext
            Loop
            tmpCLMMCOProcFee = (FFSpercentage * tmpCLMTotalActual) / 100
            lstInvoice.ListItems(AA).SubItems(468) = tmpCLMMCOProcFee
        
            'If IsNumeric(lstInvoice.ListItems(AA).SubItems(468)) Then
            '    If lstInvoice.ListItems(AA).SubItems(468) <= 0 Then
            '        tmpCLMMCOProcFee = (10 * tmpCLMTotalActual) / 100
            '    Else
            '        tmpCLMMCOProcFee = Format(IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(468)), lstInvoice.ListItems(AA).SubItems(468), 0), "#,##0.00")
            '    End If
            'End If
            
        Else
            tmpCLMMCOProcFee = 0
        End If
        
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(158)) Then
            If lstInvoice.ListItems(AA).SubItems(158) <= 0 Then
                tmpCLMPayableByMedix2C = CDbl(tmpCLMTotalActual) - CDbl(tmpCLMMCOProcFee)
            Else
                tmpCLMPayableByMedix2C = Format(lstInvoice.ListItems(AA).SubItems(158), "#,##0.00")
            End If
        Else
            tmpCLMPayableByMedix2C = CDbl(tmpCLMTotalActual) - CDbl(tmpCLMMCOProcFee)
        End If
        
        'tmpCLMMCOProcFee = Format(IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(468)), lstInvoice.ListItems(AA).SubItems(468), 0), "#,##0.00")
        'tmpCLMMCOProcStatus = IIf(Len(lstInvoice.ListItems(AA).SubItems(469)), lstInvoice.ListItems(AA).SubItems(469), "")
        tmpCLMProcMth = IIf(Len(lstInvoice.ListItems(AA).SubItems(470)), Mid(lstInvoice.ListItems(AA).SubItems(470), 1, 2), "")
        tmpCLMProcYR = IIf(Len(lstInvoice.ListItems(AA).SubItems(470)), Mid(lstInvoice.ListItems(AA).SubItems(470), 4, 4), "")
        tmpCLMCoPayCLNColl = Format(IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(471)), lstInvoice.ListItems(AA).SubItems(471), 0), "#,##0.00")
        
        If Trim(strGRPMCOStatus) = "F" Then
            If IsNumeric(tmpCLMFinalApprove) Then
                tmpCLMBordxAmt = Format(CDbl(tmpCLMTotalActual) + CDbl(tmpNonGPFee), "#,##0.00")
                'Format(CDbl(tmpCLMFinalApprove) + CDbl(tmpNonGPFee), "#,##0.00")
            Else
                tmpCLMBordxAmt = Format(CDbl(tmpNonGPFee), "#,##0.00")
            End If
        Else
            If IsNumeric(tmpCLMPayableByMedix2C) Then
                tmpCLMBordxAmt = Format(CDbl(tmpCLMTotalActual) + CDbl(tmpNonGPFee), "#,##0.00") 'Format(CDbl(tmpCLMPayableByMedix2C) + CDbl(tmpNonGPFee), "#,##0.00")
            Else
                tmpCLMBordxAmt = Format(CDbl(tmpNonGPFee), "#,##0.00")
            End If
        End If
        
        tmpLimitExceed = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(478)), lstInvoice.ListItems(AA).SubItems(478), 0)
        
        If IsNumeric(txtExcess) = False Then
            txtExcess = 0
        End If
        
        'New Checking for Bordx Amount
        If tmpCLMBordxAmt = 0 Then
            If IsNumeric(lblCLNBillAmt) Then
                tmpCLMBordxAmt = CDbl(lblCLNBillAmt)
            End If
        End If
        
        If tmpCLMFinalApprove = 0 Then
            If IsNumeric(lblCLNBillAmt) Then
                tmpCLMFinalApprove = CDbl(lblCLNBillAmt)
            End If
        End If
        
        SaveINVClaimSql = ""
        SaveINVClaimSql = "Exec SaveINVClaim_2015 '" & tmpCASnumber & "', '" & tmpClaimOpenDate & "', '" & _
                            tmpClaimOpenBy & "', '" & tmpClaimCloseDate & "', '" & tmpClaimClosedBy & "', '" & _
                            tmpClaimLastUpdateDate & "', '" & tmpClaimLastUpdateUser & "'," & CDbl(tmpCLMCoPayColl) & "," & _
                            CDbl(tmpCLMPayableByMedix2C) & "," & tmpCLMCLNExcess & ", " & CDbl(tmpCLMTotalActual) & ", " & _
                            CDbl(tmpCLMFinalApprove) & "," & tmpCLMExcess & "," & tmpCLMBalanceLimit & "," & _
                            CDbl(tmpCLMPreviousBalance) & "," & tmpCLMPaidBYMedix2M & "," & tmpCLMPaidBYMedix2C & "," & _
                            tmpCLMAdjAnnLmt & "," & tmpClmNonGPFee & "," & tmpCLMMCOProcFee & ",'" & _
                            tmpCLMMCOProcStatus & "','" & tmpCLMProcMth & "', '" & tmpCLMProcYR & "'," & tmpCLMCoPayCLNColl & "," & _
                            CDbl(tmpCLMBordxAmt) & "," & tmpLimitExceed & ", " & txtExcess & ",'" & tmpMBMCostCenter & "','" & tmpMBMDeptCode & "','" & tmpsubsidiary & "','" & tmpGRPCoCode & "','" & tmpMBMDepartment & "'"
    MediConnCISWS.Execute SaveINVClaimSql
End Sub

Public Sub GetClinicLimit()
GPLimit = 0
SPLimit = 0
GPCoPay = 0
SPCoPay = 0

Dim SqlQry As String
Dim CLNLimit As New ADODB.Recordset
        SqlQry = "select GPMax , SPMax , CLNCode,GPCopay,SPCoPay  from Tbl_CLNGPLimits where CLNCode ='" & Trim(TxtClinicCode) & "'"
        CLNLimit.Open SqlQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
        Do While Not CLNLimit.EOF
            GPLimit = CLNLimit!GPMax
            SPLimit = CLNLimit!SPMax
            GPCoPay = CLNLimit!GPCoPay
            SPCoPay = CLNLimit!SPCoPay
        CLNLimit.MoveNext
        Loop
        CLNLimit.Close
        Set CLNLimit = Nothing
End Sub
Private Function ExportToExcelError(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    On Error Resume Next
    Set objExcel = New Excel.Application

    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False

    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Error.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
            
    intCounter = 0
        
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
    intCounter = 0
    intStartRow = 5
    
    If lvw Is Nothing = False Then
        For Each itm In lvw.ListItems
            intCounter = intCounter + 1
            For i = 1 To intColumns
                If i = 1 Then
                    strArray(intCounter, 1) = itm.Text
                Else
                    strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
                End If
            Next i
        Next itm
    End If

    With objWorksheet
        If lvw Is Nothing = False And intCounter <> 0 Then
            .Range(.Cells(5, 1), _
            .Cells(4 + intCounter, intColumns)) = strArray
        End If
    End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "C") = Date
    End If
        
    objWorksheet.Columns.AutoFit
    objExcel.Interactive = True
    
    ExportToExcelError = True
ExitFunction:
    Screen.MousePointer = vbDefault
    Exit Function
HANDLE_ERROR:
    MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
    Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
    Set objRange = Nothing
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    objExcel.Quit
    GoTo ExitFunction
End Function
'add by shan 05022008
'Private Sub UpdateMBMBalanceMontly()
'
'Dim tmpDateVisitMonth As String
'tmpDateVisitMonth = Format(lstInvoice.ListItems(AA).SubItems(9), "yyyy/mm/dd")
'CASENumberMonth = Trim(lstInvoice.ListItems(AA).SubItems(2))

'Dim status As String
'Dim cmd As New ADODB.Command
'Dim Prm1 As New ADODB.Parameter
'Dim prm2 As New ADODB.Parameter
'Dim Prm3 As New ADODB.Parameter
'Dim rst2 As New ADODB.Recordset

'Set cmd.ActiveConnection = medixmasterconn
'cmd.CommandText = "UpdateMontlyLimit"
'cmd.CommandType = adCmdStoredProc
'cmd.CommandTimeout = 300
'Set Prm1 = cmd.CreateParameter("MBMNumber", adVarChar, adParamInput, 15, CASENumberMonth)
'cmd.Parameters.Append Prm1
'Set prm2 = cmd.CreateParameter("FinalAmountForMonthLimit", adCurrency, adParamInput, 20, FinalAmountForMonthLimit)
'cmd.Parameters.Append prm2
'Set Prm3 = cmd.CreateParameter("tmpDateVisit", adDate, adParamInput, 20, tmpDateVisitMonth)
'cmd.Parameters.Append Prm3
'rst2.CursorLocation = adUseClient
'rst2.CursorType = adOpenForwardOnly
'rst2.CacheSize = 100
'Set rst2 = cmd.Execute

'rst2.Close
'Set rst2 = Nothing

    
'End Sub

Private Sub UpdateMBMBalance()
Dim MBM As New ADODB.Recordset
Dim StrSql As String
Dim FinalBalance  As Double
Dim FinalNewBalance As Double
Dim TotalApprove As Double
Dim tmpAccumAmt As Double
Dim tmpMBMAdj As Double
Dim tmpNewMBMAdj As Double
Dim tmpPay2CLN As Double
Dim TempA32 As Double
Dim TempA33 As Double
Dim tmpMaternity As Double

Dim tmpMBMNum As String
Dim tmpMBMArr() As String
tmpMBMNum = Trim(lstInvoice.ListItems(AA).SubItems(2))
Dim Pos As Integer
Pos = 0
Pos = InStr(tmpMBMNum, "-")

If Pos > 0 Then
    tmpMBMArr() = Split(tmpMBMNum, "-")
    tmpMBMNum = tmpMBMArr(0)
Else
    tmpMBMNum = (tmpMBMNum)
End If


    'StrSql = ""
    'StrSql = "Select MBMNUmber , MBMCLNBalLmt, MBMAccumAmt, MBMAdjAnnLmt from MBMCLNOthers " & _
    '        " where MBMNUmber ='" & Trim(lstInvoice.ListItems(AA).SubItems(2)) & "'"
            
    StrSql = ""
    StrSql = "Select MBMNUmber , MBMCLNBalLmt, MBMAccumAmt, MBMAdjAnnLmt from MBMCLNOthers " & _
            " where MBMNUmber ='" & tmpMBMNum & "'"
               
    
    If Trim(lstInvoice.ListItems(AA).SubItems(3)) <> "00" Then
        If lstInvoice.ListItems(AA).SubItems(82) = "B" Or lstInvoice.ListItems(AA).SubItems(82) = "D" Then 'SuppLimitType
            StrSql = "Select MBMCNUmber , MBMCCLNBalLmt as MBMCLNBalLmt, MBMCAccumAmt as MBMAccumAmt, " & _
            "MBMCAdjAnnLmt as MBMAdjAnnLmt from MBMCoveredPersonsCLN " & _
            " where MBMCNUmber ='" & tmpMBMNum & "' And MBMCCoverID   ='" & "01" & "'"
        ElseIf lstInvoice.ListItems(AA).SubItems(82) = "C" Then
            StrSql = "Select MBMCNUmber , MBMCCLNBalLmt as MBMCLNBalLmt, MBMCAccumAmt as MBMAccumAmt, " & _
            "MBMCAdjAnnLmt as MBMAdjAnnLmt from MBMCoveredPersonsCLN " & _
            " where MBMCNUmber ='" & tmpMBMNum & "' And MBMCCoverID   ='" & Trim(lstInvoice.ListItems(AA).SubItems(3)) & "'"
        End If
    End If

    MBM.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    TotalApprove = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(161)), lstInvoice.ListItems(AA).SubItems(161), 0)
    
    tmpMBMAdj = IIf(IsNumeric(MBM!MBMAdjAnnLmt), MBM!MBMAdjAnnLmt, 0)
    tmpNewMBMAdj = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(113)), lstInvoice.ListItems(AA).SubItems(113), 0)
    tmpMBMAdj = tmpMBMAdj + tmpNewMBMAdj
    MBM!MBMAdjAnnLmt = CDbl(tmpMBMAdj)
    tmpPay2CLN = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(158)), lstInvoice.ListItems(AA).SubItems(158), 0)
    
    tmpAccumAmt = IIf(IsNumeric(MBM!MBMAccumAmt), MBM!MBMAccumAmt, 0)
    MBM!MBMAccumAmt = tmpAccumAmt + CDbl(TotalApprove)
    TempA32 = IIf(IsNumeric(tmpCLMTyphiad), tmpCLMTyphiad, 0)
    TempA33 = IIf(IsNumeric(tmpCLMOtherNoBalDed), tmpCLMOtherNoBalDed, 0)
    tmpMaternity = IIf(IsNumeric(A8), A8, 0)

    If lstInvoice.ListItems(AA).SubItems(70) <> "FULL" Then
        If Trim(strGRPMCOStatus) = "F" Then
            If Trim(strPLNGrossStatus) = "N" Then
                FinalBalance = MBM!MBMCLNBalLmt - (TotalApprove - (TempA32 + TempA33 + tmpMaternity))
            ElseIf Trim(strPLNGrossStatus) = "Y" Then
                FinalBalance = MBM!MBMCLNBalLmt - (tmpPay2CLN - (TempA32 + TempA33 + tmpMaternity))
             End If
        Else
            If Trim(strPLNGrossStatus) = "Y" Then
                FinalBalance = MBM!MBMCLNBalLmt - (tmpPay2CLN - (TempA32 + TempA33 + tmpMaternity))
            ElseIf Trim(strPLNGrossStatus) = "N" Then
                FinalBalance = MBM!MBMCLNBalLmt - (TotalApprove - (TempA32 + TempA33 + tmpMaternity))
            End If
        End If
            FinalBalance = IIf(FinalBalance < 0, 0, FinalBalance)
            MBM!MBMCLNBalLmt = FinalBalance
    End If
    
     ' add by shan 05022008
        If Trim(strGRPMCOStatus) = "F" Then
            If Trim(strPLNGrossStatus) = "N" Then
                FinalAmountForMonthLimit = (TotalApprove - (TempA32 + TempA33 + tmpMaternity))
            ElseIf Trim(strPLNGrossStatus) = "Y" Then
                FinalAmountForMonthLimit = (tmpPay2CLN - (TempA32 + TempA33 + tmpMaternity))
            End If
        Else
            If Trim(strPLNGrossStatus) = "Y" Then
                FinalAmountForMonthLimit = (tmpPay2CLN - (TempA32 + TempA33 + tmpMaternity))
            ElseIf Trim(strPLNGrossStatus) = "N" Then
                FinalAmountForMonthLimit = (TotalApprove - (TempA32 + TempA33 + tmpMaternity))
            End If
        End If
    
    MBM.update
    MBM.Close
    Set MBM = Nothing
End Sub

Private Sub DisplayBillAmt()
Dim tmpBillAmt As Double
    tmpBillAmt = getTotal
    lblDispBillAmt = Format(tmpBillAmt, "#,##0.00")
    lblDispCLNBillAmt = Format(tmpBillAmt, "#,##0.00")
    
    
    If tmpGrpCompany = "C014" Then   ' Group company for Prince Court
        If Len(Trim(C1)) Then
            'LimitValue = TmpVisit
            lblActCoPay = 5
            A29 = 5
        Else
            'LimitValue = TmpSPVisitLimit
            lblActCoPay = 0
            A29 = 0
        End If
    End If
    
    
    Call A29_Change
End Sub

Private Sub DrugList(tmpMedType As String)
Dim Drug As New ADODB.Recordset
    If tmpcboDrugcnt = 0 And tmpcboDrugLen = 2 Then
        Drug.Open "Select DrugName From CLNDrugList where DrugName like '%" & tmpMedType & "%' ", MediConnCISDB, adOpenKeyset, adLockOptimistic
        Do While Not Drug.EOF
            If tmpcboDrugList = 1 Then
                cboMedType1(0).AddItem Drug!DrugName
            ElseIf tmpcboDrugList = 2 Then
                cboMedType1(1).AddItem Drug!DrugName
            ElseIf tmpcboDrugList = 3 Then
                cboMedType1(2).AddItem Drug!DrugName
            ElseIf tmpcboDrugList = 4 Then
                cboMedType1(3).AddItem Drug!DrugName
            ElseIf tmpcboDrugList = 5 Then
                cboMedType1(4).AddItem Drug!DrugName
            ElseIf tmpcboDrugList = 6 Then
                cboMedType1(5).AddItem Drug!DrugName
            ElseIf tmpcboDrugList = 7 Then
                cboMedType1(6).AddItem Drug!DrugName
            End If
        Drug.MoveNext
        Loop
        Drug.Close
        Set Drug = Nothing
    End If
End Sub


