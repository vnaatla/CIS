VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPaymentCLNByInvoice 
   Caption         =   "Payment to Clinic By Invoice"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtPVStatus 
      Height          =   285
      Left            =   5160
      TabIndex        =   55
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView lstPaymentCLN 
      Height          =   3855
      Left            =   120
      TabIndex        =   16
      Top             =   1950
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6800
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   26
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Pymt Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Clinic Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Clinic Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Invoice No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Inv. Date"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Pay2CLN"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Amt Paid"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "OS Amt"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Amt Paying"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Bordx No"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Entry By"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Pay Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "PV No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Cheque No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "Cheque Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Chq Amt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Text            =   "A/C Period"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Bank"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "AccountNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Remarks"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "OPBatch"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Email"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "AccNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Payee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "WordingOne"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   57
      Top             =   240
      Width           =   11775
      Begin VB.CheckBox ChkBatchNo 
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
         Left            =   5280
         TabIndex        =   90
         Top             =   465
         Width           =   375
      End
      Begin VB.CheckBox ChkInvDate 
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
         Left            =   11280
         TabIndex        =   65
         Top             =   150
         Width           =   375
      End
      Begin VB.CheckBox ChkInvNo 
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
         Left            =   5280
         TabIndex        =   64
         Top             =   210
         Width           =   375
      End
      Begin VB.CheckBox chkPeriodMth 
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
         Left            =   11280
         TabIndex        =   63
         Top             =   990
         Width           =   375
      End
      Begin VB.CheckBox chkPVDate 
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
         Left            =   11280
         TabIndex        =   62
         Top             =   735
         Width           =   375
      End
      Begin VB.CheckBox chkPVNo 
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
         Left            =   11280
         TabIndex        =   61
         Top             =   450
         Width           =   375
      End
      Begin VB.CheckBox ChkPeriodYr 
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
         Left            =   11280
         TabIndex        =   60
         Top             =   1305
         Width           =   375
      End
      Begin VB.CheckBox ChkClinic 
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
         Left            =   5280
         TabIndex        =   59
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpComp 
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
         Left            =   5280
         TabIndex        =   58
         Top             =   720
         Width           =   255
      End
      Begin MSMask.MaskEdBox txtInvDate2 
         Height          =   315
         Left            =   9840
         TabIndex        =   8
         Top             =   135
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
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
         Left            =   7920
         TabIndex        =   7
         Top             =   135
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPVNo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7920
         MaxLength       =   15
         TabIndex        =   9
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox txtPVNo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   10
         Top             =   420
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtPVDate1 
         Height          =   315
         Left            =   7920
         TabIndex        =   11
         Top             =   675
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboACPeriodMthFr 
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
         ItemData        =   "frmPaymentCLNByInvoice.frx":0000
         Left            =   7920
         List            =   "frmPaymentCLNByInvoice.frx":0002
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodYRFr 
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
         ItemData        =   "frmPaymentCLNByInvoice.frx":0004
         Left            =   7920
         List            =   "frmPaymentCLNByInvoice.frx":0006
         TabIndex        =   15
         Top             =   1245
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtPVDate2 
         Height          =   315
         Left            =   9840
         TabIndex        =   12
         Top             =   675
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboACPeriodMthTo 
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
         ItemData        =   "frmPaymentCLNByInvoice.frx":0008
         Left            =   9840
         List            =   "frmPaymentCLNByInvoice.frx":000A
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodYRTo 
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
         ItemData        =   "frmPaymentCLNByInvoice.frx":000C
         Left            =   9840
         List            =   "frmPaymentCLNByInvoice.frx":000E
         TabIndex        =   21
         Top             =   1245
         Width           =   1335
      End
      Begin VB.TextBox TxtInvNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   0
         Top             =   180
         Width           =   1560
      End
      Begin VB.TextBox txtInvNo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         MaxLength       =   30
         TabIndex        =   1
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox txtOPBatchNo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   2
         Top             =   435
         Width           =   1575
      End
      Begin VB.TextBox txtOPBatchNo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   3
         Top             =   435
         Width           =   1575
      End
      Begin VB.ComboBox CboPayor 
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
         Left            =   1560
         TabIndex        =   5
         Top             =   690
         Width           =   3615
      End
      Begin VB.ComboBox CboPayment 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   975
         Width           =   3615
      End
      Begin VB.ComboBox CboCLN 
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1260
         Width           =   3615
      End
      Begin VB.Label Label35 
         Caption         =   "OP Batch No"
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
         TabIndex        =   92
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "To"
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
         Left            =   3240
         TabIndex        =   91
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label38 
         Caption         =   "To"
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
         Left            =   3240
         TabIndex        =   88
         Top             =   195
         Width           =   375
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   9885
         TabIndex        =   81
         Top             =   945
         Width           =   255
      End
      Begin VB.Label Label30 
         Caption         =   "A/C Period Year"
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
         Left            =   6360
         TabIndex        =   80
         Top             =   1305
         Width           =   1455
      End
      Begin VB.Label Label31 
         Caption         =   "To"
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
         Left            =   9450
         TabIndex        =   79
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label32 
         Caption         =   "To"
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
         Left            =   9450
         TabIndex        =   78
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label Label34 
         Caption         =   "A/C Period Month"
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
         Left            =   6360
         TabIndex        =   77
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "To"
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
         Left            =   9450
         TabIndex        =   76
         Top             =   135
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Invoice Date"
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
         Left            =   6360
         TabIndex        =   75
         Top             =   195
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "To"
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
         Left            =   9450
         TabIndex        =   74
         Top             =   735
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "PV Date"
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
         Left            =   6360
         TabIndex        =   73
         Top             =   765
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Left            =   9450
         TabIndex        =   72
         Top             =   465
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "PV Number"
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
         Left            =   6360
         TabIndex        =   71
         Top             =   495
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Payment Status"
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
         TabIndex        =   70
         Top             =   1035
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Invoice No"
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
         TabIndex        =   69
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Group Company"
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
         TabIndex        =   68
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Clinic"
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
         TabIndex        =   67
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5280
         TabIndex        =   66
         Top             =   1080
         Width           =   255
      End
   End
   Begin VB.Frame FramePV 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   120
      TabIndex        =   29
      Top             =   5760
      Width           =   11775
      Begin VB.ComboBox CboPayto 
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
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   150
         Width           =   5290
      End
      Begin VB.TextBox txtCLNEmail 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6840
         MaxLength       =   30
         TabIndex        =   32
         Top             =   150
         Width           =   4695
      End
      Begin MSMask.MaskEdBox txtChequeDate 
         Height          =   315
         Left            =   3000
         TabIndex        =   24
         Top             =   435
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtChequeAmt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         MaxLength       =   8
         TabIndex        =   25
         Top             =   435
         Width           =   1095
      End
      Begin VB.ComboBox CboACPeriodMth 
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
         ItemData        =   "frmPaymentCLNByInvoice.frx":0010
         Left            =   5280
         List            =   "frmPaymentCLNByInvoice.frx":0012
         TabIndex        =   17
         Top             =   1395
         Width           =   855
      End
      Begin VB.ComboBox CboACPeriodYr 
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
         ItemData        =   "frmPaymentCLNByInvoice.frx":0014
         Left            =   6840
         List            =   "frmPaymentCLNByInvoice.frx":0016
         TabIndex        =   18
         Top             =   1395
         Width           =   975
      End
      Begin VB.TextBox txtChequeNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   20
         TabIndex        =   23
         Top             =   440
         Width           =   1215
      End
      Begin VB.TextBox txtAmtWording 
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
         Height          =   285
         Left            =   840
         MaxLength       =   500
         TabIndex        =   26
         Top             =   690
         Width           =   5295
      End
      Begin MSMask.MaskEdBox txtPVDate 
         Height          =   315
         Left            =   2760
         TabIndex        =   28
         Top             =   945
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtPVNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   15
         TabIndex        =   27
         Top             =   950
         Width           =   1215
      End
      Begin VB.ComboBox CboBank 
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
         ItemData        =   "frmPaymentCLNByInvoice.frx":0018
         Left            =   4440
         List            =   "frmPaymentCLNByInvoice.frx":001A
         TabIndex        =   30
         Text            =   "CB - CITIBANK"
         Top             =   950
         Width           =   1695
      End
      Begin VB.TextBox txtAccNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6840
         MaxLength       =   30
         TabIndex        =   34
         Top             =   410
         Width           =   4695
      End
      Begin VB.TextBox TxtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6840
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   660
         Width           =   4695
      End
      Begin VB.TextBox txtCaseNo 
         Height          =   285
         Left            =   10080
         TabIndex        =   31
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Wording"
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
         Left            =   120
         TabIndex        =   95
         Top             =   700
         Width           =   735
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   120
         TabIndex        =   94
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "Email"
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
         Index           =   1
         Left            =   6240
         TabIndex        =   93
         Top             =   165
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "Chq. Date"
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
         Left            =   2160
         TabIndex        =   53
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Remark"
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
         Left            =   6240
         TabIndex        =   49
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "A/C Mth"
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
         Left            =   4320
         TabIndex        =   47
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label24 
         Caption         =   "A/C Yr"
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
         Left            =   6240
         TabIndex        =   45
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Chq. Amt"
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
         Index           =   0
         Left            =   4320
         TabIndex        =   43
         Top             =   465
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Chq. No"
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
         TabIndex        =   41
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "PV Date"
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
         Left            =   2160
         TabIndex        =   39
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "PV No"
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
         TabIndex        =   37
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Bank"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   4080
         TabIndex        =   35
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label27 
         Caption         =   "A/C No"
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
         Index           =   0
         Left            =   6240
         TabIndex        =   33
         Top             =   450
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   120
      TabIndex        =   82
      Top             =   7040
      Width           =   11775
      Begin VB.TextBox txttmpDollar 
         Height          =   285
         Left            =   3360
         TabIndex        =   97
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txttmpCents 
         Height          =   285
         Left            =   4200
         TabIndex        =   96
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11040
         TabIndex        =   56
         Top             =   500
         Width           =   615
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print to &File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9000
         TabIndex        =   54
         Top             =   500
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrintChqRange 
         Caption         =   "Print Chq"
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
         Left            =   8160
         TabIndex        =   52
         Top             =   500
         Width           =   855
      End
      Begin VB.CommandButton cmdPrintPV 
         Caption         =   "Print P&V"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7440
         TabIndex        =   50
         Top             =   500
         Width           =   735
      End
      Begin VB.CommandButton CmdPreviewPV 
         Caption         =   "Previ&w PV"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4080
         TabIndex        =   89
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7440
         TabIndex        =   38
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8160
         TabIndex        =   40
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10440
         TabIndex        =   48
         Top             =   500
         Width           =   615
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9840
         TabIndex        =   44
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9000
         TabIndex        =   42
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdCheck 
         Caption         =   "&Check All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10440
         TabIndex        =   46
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7440
         TabIndex        =   19
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8160
         TabIndex        =   20
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdPrintChq 
         Caption         =   "Print Chq"
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
         Left            =   7560
         TabIndex        =   51
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label LblAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Left            =   840
         TabIndex        =   86
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
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
         TabIndex        =   85
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Left            =   840
         TabIndex        =   84
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label Label29 
         Caption         =   "Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "   Payment to Clinic - Invoice"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   0
      TabIndex        =   87
      Top             =   0
      Width           =   11925
   End
End
Attribute VB_Name = "frmPaymentCLNByInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intExit As Integer
Dim m_blnDirection(1 To 22) As Boolean

Dim EditFlag As Boolean
Dim Check As Boolean

Private Sub CboACPeriodMth_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboACPeriodMth.Text, CboACPeriodMth.SelStart) + Chr(KeyAscii) + Mid(CboACPeriodMth.Text, CboACPeriodMth.SelStart + CboACPeriodMth.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboACPeriodMth.ListCount - 1
       
        If NewText = LCase(Left(CboACPeriodMth.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboACPeriodMth.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboACPeriodMth.Text = ValidValue
               CboACPeriodMth.SelStart = 0
               CboACPeriodMth.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub CboACPeriodYr_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboACPeriodYr.Text, CboACPeriodYr.SelStart) + Chr(KeyAscii) + Mid(CboACPeriodYr.Text, CboACPeriodYr.SelStart + CboACPeriodYr.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboACPeriodYr.ListCount - 1
       
        If NewText = LCase(Left(CboACPeriodYr.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboACPeriodYr.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
           CboACPeriodYr.Text = ValidValue
           CboACPeriodYr.SelStart = 0
           CboACPeriodYr.SelLength = Len(ValidValue)
        End If
    End If

End Sub

Private Sub CboBank_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboBank.Text, CboBank.SelStart) + Chr(KeyAscii) + Mid(CboBank.Text, CboBank.SelStart + CboBank.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboBank.ListCount - 1
       
        If NewText = LCase(Left(CboBank.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboBank.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboBank.Text = ValidValue
               CboBank.SelStart = 0
               CboBank.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub ChkClinic_Click()
Dim ClnRec As New ADODB.Recordset

    CboCLN.Clear
    If ChkClinic.Value = 1 Then
        ClnRec.Open "Select CLNCode, CLNName from CLN WHERE CLNPanel = 'Y' OR CLNPanel = 'C'", medixmasterconn, adOpenKeyset, adLockOptimistic
        Do Until ClnRec.EOF
            If Len(ClnRec!CLNName) Then
                CboCLN.AddItem ClnRec!CLNCode & " - " & Trim(ClnRec!CLNName)
            End If
            ClnRec.MoveNext
        Loop
        ClnRec.Close
        Set ClnRec = Nothing
    End If

End Sub

Private Sub ChkGrpComp_Click()
Dim Company As New ADODB.Recordset
    
    Company.Open "Select GRPCoName from GRPCompany", MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do Until Company.EOF
        CboPayor.AddItem Trim(Company!GrpCoName)
        Company.MoveNext
    Loop
    Company.Close
    Set Company = Nothing

End Sub

Private Sub CmdAdd_Click()
Dim ZZ As Integer
Dim ItemChk As Boolean
    
    ItemChk = False
    For ZZ = 1 To lstPaymentCLN.ListItems.Count
        If lstPaymentCLN.ListItems(ZZ).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next ZZ

    CmdUpdate.Enabled = False
    
    If ItemChk Then
        Call ClearFramePV
        cmdAdd.Visible = False
        CmdSave.Visible = True
        FramePV.Enabled = True
        CboBank.Enabled = True
        txtChequeAmt = CDbl(LblAmt)
        txtPVDate.SetFocus
    Else
        MsgBox "Please checked at least 1 Claim  ! ", vbCritical + vbOKOnly, DialogBoxTitle
    End If
End Sub

Private Sub CmdCheck_Click()
    If lstPaymentCLN.ListItems.Count <> 0 Then
        SetCheckAllItems True
    Else
        MsgBox "Please select a record.", vbInformation
    End If
End Sub

Private Sub SetCheckAllItems(bState As Boolean)
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim r As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstPaymentCLN.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstPaymentCLN.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub

Private Sub CmdPreviewPV_Click()
Dim ItemChecked As Boolean
Dim II As Integer
    ItemChecked = False
    For II = 1 To lstPaymentCLN.ListItems.Count
        If lstPaymentCLN.ListItems(II).Checked = True Then
            ItemChecked = True
        Exit For
        End If
    Next II
    
    If ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    Else
        Call PrintPV
    End If
End Sub

Private Sub cmdPrintChq_Click()
'On Error Resume Next
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim strsql As String
Dim chq As New ADODB.Recordset
Dim RecFound As Boolean
Dim RecCnt As Integer
Dim tmpDD As String
Dim tmpMM As String
Dim tmpYY As String
    Screen.MousePointer = vbHourglass
    If MsgBox("Do you want to print now?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then

        If TxtPVNo = "" Then
            MsgBox "Please Enter or Generate PV No", vbOKOnly + vbCritical
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If CboPayto = "" Then
            MsgBox "Please Enter Payee", vbOKOnly + vbCritical
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If txtChequeNo = "" Then
            MsgBox "Please Enter Cheque No", vbOKOnly + vbCritical
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If txtChequeDate = "__/__/____" Then
            MsgBox "Please Enter Cheque Date", vbOKOnly + vbCritical
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If IsNumeric(txtChequeAmt) = False Then
            MsgBox "Please Enter Cheque Amount", vbOKOnly + vbCritical
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If txtAmtWording = "" Then
            MsgBox "Please Enter Wording Amount", vbOKOnly + vbCritical
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            'If objExcel.Visible = False Then
                objExcel.Visible = False
            'End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\PaymentCheque.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
            
        'If Err.Number > 0 Then
        '    MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        '    Exit Sub
        'End If
            
        On Error GoTo HANDLE_ERROR
        ' Don't allow user to affect workbook
    '        strsql = "select CASNumber, ClaimVersion, BordxRefNo, MBMNumber, MBMPatientCovID, " & _
    '                " HLTCode, ClaimCloseDate, PVNo, PVChequeNo, PVDate, MBMPaidAmt, CLMPayableByMedix2M, " & _
    '                " PVBankCharges from View_Payment2MBMNew where PVNo = '" & TxtPVNo & "' AND PVStatus='A'"
    
            
        strsql = "select PVCLNChequeNo, PVCLNChequeDate, PVCLNChequeAmt, " & _
        " PVWordingOne,PVWordingTwo, CLNPName from VIEW_Payment2CLN_Chq where PVCLNChequeNo = '" & Trim(txtChequeNo) & "'"
        chq.Open strsql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
        
        Do While Not chq.EOF
                objWorksheet.Cells(3, "A") = chq!CLNPName
                objWorksheet.Cells(5, "A") = chq!PVWordingOne
                objWorksheet.Cells(7, "A") = chq!PVWordingTwo
                objWorksheet.Cells(5, "E") = Format(chq!PVCLNChequeAmt, "#,##0.00")
                tmpDD = Format(chq!PVCLNChequeDate, "DD")
                tmpMM = Format(chq!PVCLNChequeDate, "MM")
                tmpYY = Format(chq!PVCLNChequeDate, "YY")
                objWorksheet.Cells(1, "H") = Mid(tmpDD, 1, 1)
                objWorksheet.Cells(1, "I") = Mid(tmpDD, 2, 1)
                objWorksheet.Cells(1, "J") = Mid(tmpMM, 1, 1)
                objWorksheet.Cells(1, "K") = Mid(tmpMM, 2, 1)
                objWorksheet.Cells(1, "L") = Mid(tmpYY, 1, 1)
                objWorksheet.Cells(1, "M") = Mid(tmpYY, 2, 1)
            chq.MoveNext
        Loop
        chq.Close
        Set chq = Nothing
        
        objWorkbook.PrintOut
        objWorkbook.SaveAs (crdPath & "Template\PaymentVoucher\" & "tmpPaycheque" & ".xlt")
        objExcel.Quit

ExitFunction:
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            Screen.MousePointer = vbDefault
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            objExcel.Quit
            GoTo ExitFunction
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdPrintChqRange_Click()
    frmPrintChequeRange.chqMode = ""
    Load frmPrintChequeRange
    frmPrintChequeRange.chqMode = "MedixPayCLN"
    frmPrintChequeRange.Show vbModal
End Sub

Private Sub CmdPrintPV_Click()
    If TxtPVNo = "" Then
        MsgBox "Please Enter or Generate PV No", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
    Else
        Call PrintPV
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

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelAnalysis1(lstPaymentCLN)
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdEdit_Click()
Dim ZZ As Integer
Dim ItemChk As Boolean

    EditFlag = True
    ItemChk = False
    For ZZ = 1 To lstPaymentCLN.ListItems.Count
        If lstPaymentCLN.ListItems(ZZ).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next ZZ

    CmdSave.Enabled = False
    FramePV.Enabled = False
    If ItemChk = False And Len(Trim(TxtPVNo)) Then
        CmdEdit.Visible = False
        CmdUpdate.Visible = True
        FramePV.Enabled = True
        EditFlag = True
        If Len(Trim(lstPaymentCLN.SelectedItem.SubItems(1))) Then
            FramePV.Enabled = True
        End If
        txtPVDate.SetFocus
    End If
End Sub

Private Sub CmdSearch_Click()
    Screen.MousePointer = vbHourglass
    CmdEdit.Visible = True
    cmdAdd.Visible = True
    CmdUpdate.Visible = False
    CmdSave.Visible = False
    LblAmt = 0
    EditFlag = False
    Call SearchRecord
    Screen.MousePointer = Default
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    Check = False
    TxtInvNo.SetFocus
    CmdExit.Enabled = True
End Sub

Private Sub CmdUpdate_Click()
Dim RecordSaved
Dim ZZ As Integer
Dim ItemChk As Boolean
Dim PVSql As String
Dim CaseNo As String
Dim PVStatus As String
Dim PaymentAmt As Double
Dim PymtIndex As String
Dim UPdRec As New ADODB.Recordset
Dim StrInvSql As String

    ItemChk = False
    For ZZ = 1 To lstPaymentCLN.ListItems.Count
        If lstPaymentCLN.ListItems(ZZ).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next ZZ
    If ItemChk = True Then
    ElseIf lstPaymentCLN.ListItems.Count = 0 Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    ElseIf CboPayto = "" Then
        MsgBox "Please Enter Pay to! ", vbCritical
    ElseIf TxtPVNo = "" Then
        MsgBox "Please Confirm PV No!", vbCritical
    ElseIf IsDate(txtPVDate) = False Then
        MsgBox "Invalid PV Date!", vbCritical
        txtPVDate.SetFocus
    ElseIf txtPVDate = "__/__/____" Then
        MsgBox "Please Enter PV Date! ", vbCritical
        txtPVDate.SetFocus
    ElseIf txtChequeNo = "" Then
            MsgBox "Please Enter Cheque Number!", vbCritical
            txtChequeNo.SetFocus
    ElseIf txtChequeDate = "__/__/____" Then
            MsgBox "Please Enter Cheque Date!", vbCritical
            txtChequeDate.SetFocus
    ElseIf txtChequeAmt = "" Then
        MsgBox "Please Enter Cheque Amt!", vbCritical
        txtChequeAmt.SetFocus
    ElseIf txtAmtWording = "" Then
        MsgBox "Please Enter Wording Amount!", vbCritical
    ElseIf Trim(CboBank) = "" Then
        MsgBox "Please Enter Bank !", vbCritical
        CboBank.SetFocus
    ElseIf Trim(CboACPeriodMth) = "" Then
        MsgBox "Please keyin Account Period Month ! ", vbCritical
        CboACPeriodMth.SetFocus
    ElseIf Trim(CboACPeriodYr) = "" Then
        MsgBox "Please keyin Account Period Year ! ", vbCritical
        CboACPeriodYr.SetFocus
    Else
        RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            StrInvSql = "Select CASNumber, PVPayCLNStatus, PVCLNIndex, CLNPaidAmt" & _
                         " from View_Payment2CLn WHERE PVCLNNo='" & Trim(TxtPVNo) & _
                         "' AND PVPayCLNStatus='F'"
                        
            UPdRec.Open StrInvSql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
            If UPdRec.RecordCount Then
                Do Until UPdRec.EOF
                    CaseNo = ""
                    PVStatus = ""
                    PymtIndex = ""
                    CaseNo = Trim(UPdRec!CasNumber)
                    PVStatus = "F"
                    PaymentAmt = UPdRec!CLNPaidAmt
                    PymtIndex = UPdRec!PVCLNIndex
                    Call UpdateRecord(CaseNo, PVStatus, PaymentAmt, PymtIndex)
                    UPdRec.MoveNext
                Loop
            End If
            UPdRec.Close
            Set UPdRec = Nothing
            CmdSave.Enabled = True
            CmdEdit.Visible = True
            cmdAdd.Visible = True
            CmdUpdate.Visible = False
            CmdSave.Visible = False
            EditFlag = False
            LblAmt = 0
            PVSql = " and PVCLNNo = '" & Trim(TxtPVNo) & "'"
            Call ClearFramePV
            lstPaymentCLN.ListItems.Clear
            Call Payment2CLNListing(PVSql)
        End If
        
        
    End If
    
End Sub

Private Sub Form_Load()
    Call ComboListing
    CboPayment.ListIndex = 0
    intExit = 0
    lblCurrent = 0
    LblAmt = 0
    Check = False
    EditFlag = False
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    CmdSave.Visible = False
    cmdAdd.Visible = True
    CmdEdit.Visible = True
    CmdUpdate.Visible = False
    FramePV.Enabled = False
    EditFlag = False
    lstPaymentCLN.ListItems.Clear
    CboPayment.ListIndex = 0
    lblCurrent = 0
    LblAmt = 0
    Check = False
    TxtInvNo.SetFocus
End Sub


Private Sub cmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

'************************Start ComboListing**************************
Private Sub ComboListing()
Dim monthstart
Dim yearstart
Dim BankRec As New ADODB.Recordset
Dim ClnRec As New ADODB.Recordset
    
    BankRec.Open "Banks", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    Do Until BankRec.EOF
        CboBank.AddItem ChkStr(BankRec!BankCode) & " - " & ChkStr(BankRec!BankDesc)
        BankRec.MoveNext
    Loop
    BankRec.Close
    Set BankRec = Nothing
        
    With CboPayment
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
'        .AddItem "3 - Partial Paid"
'        .AddItem "4 - Paid & Partial Paid"
'        .AddItem "5 - Unpaid & Partial Paid"
    End With
    
    For monthstart = 1 To 12
        CboACPeriodMth.AddItem Format(monthstart, "00")
        cboACPeriodMthFr.AddItem Format(monthstart, "00")
        cboACPeriodMthTo.AddItem Format(monthstart, "00")
    Next monthstart
    
    For yearstart = 1999 To 3000
        CboACPeriodYr.AddItem yearstart
        cboACPeriodYRFr.AddItem yearstart
        cboACPeriodYRTo.AddItem yearstart
    Next yearstart
        

    CboPayto.Clear
    ClnRec.Open "Select CLNPName from CLN WHERE CLNPanel = 'Y' OR CLNPanel = 'C'", medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until ClnRec.EOF
        If Len(ClnRec!CLNPName) Then
            CboPayto.AddItem ClnRec!CLNPName
        End If
        ClnRec.MoveNext
    Loop
    ClnRec.Close
    Set ClnRec = Nothing
End Sub

Private Function SearchRecord()
Dim strsql As String
Dim StrAppend As String
Dim sql As String
Dim GrpCompanyName As String
    
    StrAppend = ""
    GrpCompanyName = ""
    
    If Len(Trim(CboPayor)) Then
         GrpCompanyName = Replace(CboPayor, "'", "''")
         StrAppend = StrAppend & " AND GroupCompany = '" & Trim(GrpCompanyName) & "'"
    End If
    If Len(Trim(CboCLN)) Then
        StrAppend = StrAppend & " AND ClinicCode = '" & Trim(Left(CboCLN, IIf(InStr(1, CboCLN, "-") - 1 < 0, Len(CboCLN), InStr(1, CboCLN, "-") - 1))) & "'"
    End If
    
    If Mid(CboPayment, 1, 1) = 1 Then
        StrAppend = StrAppend & " and PVPayCLNStatus = 'F'"
    End If
 
    If Mid(CboPayment, 1, 1) = 2 Then
        StrAppend = StrAppend & " and (PVPayCLNStatus = '' or PVPayCLNStatus is null) "
    End If
 
'    If Mid(CboPayment, 1, 1) = 3 Then
'        strAppend = strAppend & " and PVPayCLNStatus = 'P'"
'    End If
 
'    If Mid(CboPayment, 1, 1) = 4 Then
'        strAppend = strAppend & " and (PVPayCLNStatus = 'P' or PVPayCLNStatus = 'F') "
'    End If
    If ChkBatchNo.Value = 1 Then
        sql = ""
        sql = "AND (OPBatchNo Is null OR OPBatchNo = '') " & ""
        StrAppend = StrAppend & sql
    End If
    
    If Len(Trim(txtOPBatchNo1)) Then
        StrAppend = StrAppend & " AND OPBatchNo >= '" & Trim(txtOPBatchNo1) & "'"
    End If
    If Len(Trim(txtOPBatchNo2)) Then
        StrAppend = StrAppend & " AND OPBatchNo <= '" & Trim(txtOPBatchNo2) & "'"
    End If
    If Len(TxtInvNo) And Trim(txtInvNo2) = "" Then
        StrAppend = StrAppend & " and InvoiceNo = '" & Trim(TxtInvNo) & "'"
    End If
    If Trim(TxtInvNo) <> "" And Trim(txtInvNo2) <> "" Then
        StrAppend = StrAppend & " and InvoiceNo >= '" & Trim(TxtInvNo) & "'"
        StrAppend = StrAppend & " and InvoiceNo <= '" & Trim(txtInvNo2) & "'"
    End If
    If Trim(TxtInvNo) = "" And Trim(txtInvNo2) <> "" Then
        StrAppend = StrAppend & " and InvoiceNo <= '" & Trim(txtInvNo2) & "'"
    End If
    If ChkInvNo.Value = 1 Then
        sql = ""
        sql = "AND (InvoiceNo Is null OR InvoiceNo = '') " & ""
        StrAppend = StrAppend & sql
    End If
    
    If IsDate(txtInvDate1) = True Then
        StrAppend = StrAppend & " And InvoiceDate >= '" & Format(txtInvDate1, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(txtInvDate2) = True Then
        StrAppend = StrAppend & " And InvoiceDate <= '" & Format(txtInvDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkInvDate.Value = 1 Then
        sql = ""
        sql = "AND (PVDate Is null OR PVDate = '') " & ""
        StrAppend = StrAppend & sql
    End If
    
    If Len(txtPVNo1) Then
        StrAppend = StrAppend & " and PVCLNNo >= '" & Trim(txtPVNo1) & "'"
    End If
        
    If Len(txtPVNo2) Then
        StrAppend = StrAppend & " and PVCLNNo <= '" & Trim(txtPVNo2) & "'"
    End If
    
    If chkPVNo.Value = 1 Then
        sql = ""
        sql = "AND (PVCLNNo Is null OR PVCLNNo = '') " & ""
        StrAppend = StrAppend & sql
    End If
    
    If IsDate(txtPVDate1) = True Then
        StrAppend = StrAppend & " And PVCLNDate >= '" & Format(txtPVDate1, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(txtPVDate2) = True Then
        StrAppend = StrAppend & " And PVCLNDate <= '" & Format(txtPVDate2, "mm/dd/yyyy") & "'"
    End If
    
    If chkPVDate.Value = 1 Then
        sql = ""
        sql = "AND (PVCLNDate Is null OR PVCLNDate = '') " & ""
        StrAppend = StrAppend & sql
    End If
        
    If Len(cboACPeriodMthFr) Then
        StrAppend = StrAppend & " and CLNAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
        
    If Len(cboACPeriodMthTo) Then
        StrAppend = StrAppend & " and CLNAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If chkPeriodMth.Value = 1 Then
        sql = ""
        sql = " AND CLNAccPeriodMth Is null "
        StrAppend = StrAppend & sql
    End If
    
    If Len(cboACPeriodYRTo) Then
        StrAppend = StrAppend & " and CLNAccPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
    End If
        
    If Len(cboACPeriodYRTo) Then
        StrAppend = StrAppend & " and CLNAccPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
    
    If ChkPeriodYr.Value = 1 Then
        sql = ""
        sql = " AND CLNAccPeriodYr Is null "
        StrAppend = StrAppend & sql
    End If
    'strAppend = strAppend & " AND CLMPayableByMedix2c <> 0 "
    Call Payment2CLNListing(StrAppend)

End Function

Private Sub cboCLN_Change()
    If InStr(CboCLN.Text, "null") Then
       CboCLN.Enabled = False
    End If
    
End Sub

Private Sub CboCLN_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCLN.Text, CboCLN.SelStart) + Chr(KeyAscii) + Mid(CboCLN.Text, CboCLN.SelStart + CboCLN.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCLN.ListCount - 1
 
        If NewText = LCase(Left(CboCLN.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboCLN.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCLN.Text = ValidValue
                CboCLN.SelStart = 0
                CboCLN.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboPayment_Change()
    If InStr(CboPayment.Text, "null") Then
       CboPayment.Enabled = False
    End If
End Sub

Private Sub CboPayment_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayment.ListCount - 1
 
        If NewText = LCase(Left(CboPayment.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayment.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPayment.Text = ValidValue
                CboPayment.SelStart = 0
                CboPayment.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub Payment2CLNListing(Optional StrAppend As String)
Dim PVRec As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim strsqlPay As String
Dim PAY As New ADODB.Recordset
Dim INVAmt As Variant
Dim Current As String
Dim TotalAmt As Variant
Dim Check As Boolean
Dim tmpPay2CLN As Double
    Current = 0
    lstPaymentCLN.ListItems.Clear
    Call ClearFramePV
    'sql = " Select * From VIEW_Payment2CLN_Invoice where 0 = 0 "
    Call ReadRecord(StrAppend, PVRec)
    Do
        Do While Not PVRec.EOF
        txtPVStatus = ""
        If Len(PVRec!PVPayCLNStatus) Then
            txtPVStatus = PVRec!PVPayCLNStatus
        End If
            
            With PVRec
                Set Record = lstPaymentCLN.ListItems.Add(, , "")
                Record.SubItems(1) = ""
                If Trim(txtPVStatus) = "F" Then
                    Record.SubItems(1) = "Full"
                ElseIf Trim(txtPVStatus) = "P" Then
                    Record.SubItems(1) = "Partial"
                End If
                Record.SubItems(2) = IIf(Len(PVRec!ClinicCode), PVRec!ClinicCode, "")
                Record.SubItems(3) = IIf(Len(PVRec!CLNName), PVRec!CLNName, "")
                Record.SubItems(4) = IIf(Len(PVRec!InvoiceNo), PVRec!InvoiceNo, "")
                If Len(PVRec!InvoiceDate) Then
                    Record.SubItems(5) = Format(PVRec!InvoiceDate, "dd/mm/yyyy")
                End If
                tmpPay2CLN = IIf(IsNumeric(PVRec!CLMPayableByMedix2C), Format(PVRec!CLMPayableByMedix2C, "#,##0.00"), 0)
                Record.SubItems(6) = Format(tmpPay2CLN, "#,##0.00")
                Record.SubItems(7) = IIf(IsNumeric(PVRec!CLNPaidAmt), Format(PVRec!CLNPaidAmt, "#,##0.00"), 0)
                Record.SubItems(8) = CDbl(Record.SubItems(6) - Record.SubItems(7))
                If Len(PVRec!BordxRefNo) Then
                    Record.SubItems(10) = Trim(PVRec!BordxRefNo)
                End If
                If Len(PVRec!INVEnteredBy) Then
                    Record.SubItems(12) = Trim(PVRec!INVEnteredBy)
                End If
               
                If Len(PVRec!PVCLNDate) Then
                    Record.SubItems(11) = Format(PVRec!PVCLNDate, "dd/mm/yyyy")
                End If
                                
                If Len(PVRec!PVCLNNo) Then
                    Record.SubItems(13) = Trim(PVRec!PVCLNNo)
                End If
                
                If Len(PVRec!PVCLNChequeNo) Then
                    Record.SubItems(14) = Trim(PVRec!PVCLNChequeNo)
                End If
                
                If Len(!PVCLNChequeDate) Then
                    Record.SubItems(15) = Trim(PVRec!PVCLNChequeDate)
                End If
                If IsNumeric(PVRec!PVCLNChequeAmt) Then
                    Record.SubItems(16) = Trim(PVRec!PVCLNChequeAmt)
                End If

                If Len(Trim(PVRec!CLNAccPeriodMth) & Trim(PVRec!CLNAccPeriodYr)) Then
                    Record.SubItems(17) = Format(PVRec!CLNAccPeriodMth & "/" & PVRec!CLNAccPeriodYr, "mm/yyyy")
                End If
                If Len(Trim(PVRec!PVCLNBank)) Then
                    Record.SubItems(18) = Trim(PVRec!PVCLNBank)
                End If
                If Len(Trim(PVRec!PVCLNAccount)) Then
                    Record.SubItems(19) = Trim(PVRec!PVCLNAccount)
                End If
                If Len(Trim(PVRec!CLNRemark)) Then
                    Record.SubItems(20) = Trim(PVRec!CLNRemark)
                End If
                If Len(Trim(PVRec!OPBatchNo)) Then
                    Record.SubItems(21) = Trim(PVRec!OPBatchNo)
                End If
                
                If Len(Trim(PVRec!CLNEmail)) Then
                    Record.SubItems(22) = Trim(PVRec!CLNEmail)
                End If
                If Len(Trim(PVRec!CLNBankAcc)) Then
                    Record.SubItems(23) = Trim(PVRec!CLNBankAcc)
                End If
                
                If Len(Trim(PVRec!CLNPName)) Then
                    Record.SubItems(24) = Trim(PVRec!CLNPName)
                End If
                
                If Len(Trim(PVRec!PVWordingOne)) Then
                    Record.SubItems(25) = Trim(PVRec!PVWordingOne)
                End If
                
                If Len(Trim(PVRec!PVWordingTwo)) Then
                    Record.SubItems(25) = Record.SubItems(25) & PVRec!PVWordingTwo
                End If
                
                Current = Current + 1
                lblCurrent = Current
            End With

            PVRec.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
'    End If
    intExit = 0
    PVRec.Close
    Set PVRec = Nothing
End Sub

Private Sub CmdSave_Click()
Dim PAYCode As String
Dim CaseNo As String
Dim PVStatus As String
Dim II As Integer
Dim i As Integer
Dim Cnt As Integer
Dim Update
Dim ItemChecked As Boolean
Dim PymtIndex As String
Dim ZZ As String
Dim TempBank As String
Dim PVSql As String
Dim BankRec As New ADODB.Recordset
Dim BankSql As String
Dim TmpPVCode As String
Dim UPdRec As New ADODB.Recordset
Dim StrInvSql As String
Dim tmpNonGpAmt As String
Dim tmpAmtPay2C As Double
Dim PVNumber As String
Dim PrevCLN, CurrCLN As String
Dim DiffCLN As Boolean
Dim CLNCnt As Integer
Dim strlsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim strsqlpymt As String
Dim pymt As New ADODB.Recordset
Dim PYMTSTATUS As Boolean
    PYMTSTATUS = False
    Cnt = 0
    ItemChecked = False
    TmpPVCode = ""
    ZZ = InStr(1, CboBank, "-")
    TempBank = Trim(Left(CboBank, IIf(ZZ - 1 <= 0, Len(Trim(CboBank)), ZZ - 1)))
    BankSql = "SELECT PVCode from Banks WHERE BankCode='" & Trim(TempBank) & "'"
    BankRec.Open BankSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    If BankRec.RecordCount Then
        TmpPVCode = Trim(BankRec!PVCode)
    End If
    BankRec.Close
    Set BankRec = Nothing
    
    For i = 1 To lstPaymentCLN.ListItems.Count
        If lstPaymentCLN.ListItems(i).Checked = True Then
            ItemChecked = True
            Exit For
        End If
    Next i
    CLNCnt = 0
    DiffCLN = False
    For i = 1 To lstPaymentCLN.ListItems.Count
        If lstPaymentCLN.ListItems(i).Checked = True Then
            CLNCnt = CLNCnt + 1
            CurrCLN = Trim(lstPaymentCLN.ListItems(i).SubItems(2))
            If CLNCnt = 1 Then
                PrevCLN = Trim(lstPaymentCLN.ListItems(i).SubItems(2))
            End If
            If PrevCLN <> CurrCLN Then
                DiffCLN = True
                Exit For
            End If
        End If
    Next i

    If ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    ElseIf CboPayto = "" Then
        MsgBox "Please Enter Pay to! ", vbCritical
    ElseIf IsDate(txtPVDate) = False Then
        MsgBox "Invalid PV Date!", vbCritical
'        txtPVDate.SetFocus
    ElseIf txtPVDate = "__/__/____" Then
        MsgBox "Please Enter PV Date! ", vbCritical
        txtPVDate.SetFocus
    ElseIf DiffCLN Then
        MsgBox "Please verify the cases, different Clinic found !", vbOKOnly + vbCritical
    ElseIf txtChequeAmt = "" Then
        MsgBox "Please Enter Cheque Amt!", vbCritical
        txtChequeAmt.SetFocus
    ElseIf txtChequeNo = "" Then
        MsgBox "Please Enter Cheque Number!", vbCritical
        txtChequeNo.SetFocus
    ElseIf txtChequeDate = "__/__/____" Then
        MsgBox "Please Enter Cheque Date!", vbCritical
        txtChequeDate.SetFocus
    ElseIf txtChequeAmt = "" Then
        MsgBox "Please Enter Cheque Amt!", vbCritical
        txtChequeAmt.SetFocus
    ElseIf txtAmtWording = "" Then
        MsgBox "Please Enter Wording Amount!", vbCritical
    ElseIf Trim(CboBank) = "" Then
        MsgBox "Please Enter Bank !", vbCritical
        CboBank.SetFocus
    ElseIf Trim(CboACPeriodMth) = "" Then
        MsgBox "Please keyin Account Period Month ! ", vbCritical
        CboACPeriodMth.SetFocus
    ElseIf Trim(CboACPeriodYr) = "" Then
        MsgBox "Please keyin Account Period Year ! ", vbCritical
        CboACPeriodYr.SetFocus
    ElseIf Trim(TmpPVCode) = "" Then
        MsgBox "Please keyin Bank - PV Code ! ", vbCritical
    Else
        Update = MsgBox("Do you want to update the record ? ", vbYesNo + vbExclamation)
        If Update = vbYes Then
            PVNumber = GenerateSPVNo(TempBank, TmpPVCode)
            TxtPVNo = ""
            TxtPVNo = PVNumber
           
            For i = 1 To lstPaymentCLN.ListItems.Count
                If lstPaymentCLN.ListItems(i).Checked = True Then
                    'StrInvSql = "Select CASNumber, PVPayCLNStatus, PVCLNIndex, CLMPayableByMedix2C, CLMNonGpFee " & _
                    '            " from View_Payment2CLn WHERE InvoiceNo='" & Trim(lstPaymentCLN.ListItems(i).SubItems(4)) & _
                    '            "' AND InvoiceDate='" & Format(lstPaymentCLN.ListItems(i).SubItems(5), "yyyy/mm/dd") & _
                    '            "' AND ClinicCode='" & Trim(lstPaymentCLN.ListItems(i).SubItems(2)) & _
                    '            "' AND PVPayCLNStatus IS NULL"
                    'UPdRec.Open StrInvSql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
                    StrInvSql = ""
                    StrInvSql = "SELECT dbo.CISCas.CASNumber FROM CISWorksheet.dbo.Claim INNER JOIN dbo.CISCas ON CISWorksheet.dbo.Claim.CasNumber = dbo.CISCAS.CasNumber " & _
                                " WHERE InvoiceNo='" & Trim(lstPaymentCLN.ListItems(i).SubItems(4)) & _
                                "' AND InvoiceDate='" & Format(lstPaymentCLN.ListItems(i).SubItems(5), "yyyy/mm/dd") & _
                                "' AND BordxRefNo='" & lstPaymentCLN.ListItems(i).SubItems(10) & _
                                "' AND ClinicCode='" & Trim(lstPaymentCLN.ListItems(i).SubItems(2)) & "'"
                    UPdRec.Open StrInvSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                    'UPdRec.Open StrInvSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                    
                    If UPdRec.RecordCount Then
                        Do Until UPdRec.EOF
                            strsqlpymt = "Select CASNumber " & _
                                " from Payment2CLN WHERE CASNumber = '" & UPdRec!CasNumber & "'"
                                pymt.Open strsqlpymt, MediConnCISPayment, adOpenKeyset, adLockOptimistic
                            If pymt.RecordCount Then
                                PYMTSTATUS = True
                            Else
                                PYMTSTATUS = False
                            End If
                            pymt.Close
                            Set pymt = Nothing
                            
                            If PYMTSTATUS = False Then
                                strlsqlCLM = "Select CLMPayableByMedix2C  from Claim WHERE CASNumber ='" & UPdRec!CasNumber & "'"
                                CLM.Open strlsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
                                
                                CaseNo = ""
                                PVStatus = ""
                                PymtIndex = ""
                                CaseNo = Trim(UPdRec!CasNumber)
                                PVStatus = Mid(lstPaymentCLN.ListItems(i).SubItems(1), 1, 1)
                                tmpAmtPay2C = IIf(IsNumeric(CLM!CLMPayableByMedix2C), CLM!CLMPayableByMedix2C, 0)
                                PymtIndex = ""
                                Call UpdateRecord(CaseNo, PVStatus, tmpAmtPay2C, PymtIndex)
                                CLM.Close
                                Set CLM = Nothing
                            End If
                            UPdRec.MoveNext
                        Loop
                    End If
                    UPdRec.Close
                    Set UPdRec = Nothing
                    CaseNo = ""
                    PVStatus = ""
                    PymtIndex = ""
                End If
            Next i
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            CmdSave.Visible = False
            cmdAdd.Visible = True
            lstPaymentCLN.ListItems.Clear
            LblAmt = 0
            PVSql = " and PVCLNNo = '" & Trim(TxtPVNo) & "'"
            Call ClearFramePV
            lstPaymentCLN.ListItems.Clear
            Call Payment2CLNListing(PVSql)
        End If
    End If
End Sub

Private Sub UpdateRecord(CaseNo As String, status As String, PaymentAmt As Double, PymtIndex As String)
Dim Update
Dim strsql As String
Dim PAY As New ADODB.Recordset
Dim strsqlPay As String
Dim CLM As New ADODB.Recordset
Dim strsqlCLM As String
Dim ZZ As String
Dim TempBank As String
Dim tmpUpdWording1 As String
Dim tmpUpdWording2 As String
Dim strupdatepymt As String

    ZZ = InStr(1, CboBank, "-")
    tmpUpdWording1 = ""
    tmpUpdWording2 = ""
    TempBank = Trim(Left(CboBank, IIf(ZZ - 1 <= 0, Len(Trim(CboBank)), ZZ - 1)))

    strsqlPay = " Select PVCLNIndex, * from Payment2CLN " & _
                " Where PVCLNIndex = '" & Trim(PymtIndex) & "'"
    PAY.Open strsqlPay, MediConnCISPayment, adOpenKeyset, adLockOptimistic
    
    If PAY.RecordCount > 0 Then
    Else
        PAY.AddNew
    End If
    With PAY
        !CasNumber = Trim(CaseNo)
        !PVCLNNo = TxtPVNo
        !PVCLNDate = Format(txtPVDate, "dd/mm/yyyy")
        !PVCLNChequeDate = IIf(txtChequeDate <> "__/__/____", Format(txtChequeDate, "dd/mm/yyyy"), "")
        !PVCLNChequeNo = IIf(Len(Trim(txtChequeNo)), txtChequeNo, "")
        !PVCLNChequeAmt = IIf(IsNumeric(txtChequeAmt), txtChequeAmt, 0)
        !PVCLNBank = IIf(Len(TempBank), Trim(TempBank), "")
        !CLNRemark = IIf(Len(Trim(TxtRemark)), TxtRemark, "")
        !CLNPaidAmt = IIf(IsNumeric(PaymentAmt), PaymentAmt, 0)
        !CLNAccPeriodMth = IIf(CboACPeriodMth <> "", Trim(CboACPeriodMth), "")
        !CLNAccPeriodYr = IIf(CboACPeriodYr <> "", Trim(CboACPeriodYr), "")
        !Pay2CLNLastUpdateDate = Date
        !Pay2CLNLastUpdateUser = Logon
        !PVPayCLNStatus = Mid(status, 1, 1)
        If Len(txtAmtWording) > 55 Then
            If Mid(txttmpDollar, 1, 3) = "And" Then
                !PVWordingOne = Mid(txttmpDollar, 5, 500)
                tmpUpdWording1 = Mid(txttmpDollar, 5, 500)
            Else
                !PVWordingOne = txttmpDollar
                tmpUpdWording1 = Mid(txttmpDollar, 5, 500)
            End If
            
            !PVWordingTwo = IIf(Len(Trim(txttmpCents)), txttmpCents & " Only", "")
        Else
            !PVWordingOne = IIf(Len(Trim(txtAmtWording)), Trim(txtAmtWording), "")
            tmpUpdWording1 = txtAmtWording
            !PVWordingTwo = ""
            tmpUpdWording2 = ""
        End If
        
    End With
    PAY.Update
    
    PAY.Close
    Set PAY = Nothing
    
    'strupdatepymt = ""
    'strupdatepymt = " update Payment2CLN set " & _
                  " PVCLNChequeDate = '" & Format(txtChequeDate, "dd/mm/yyyy") & "', " & _
                  " PVCLNChequeNo = '" & txtChequeNo & "', " & _
                  " PVCLNChequeAmt = " & txtChequeAmt & ", " & _
                  " PVWordingOne = '" & tmpUpdWording1 & "', " & _
                  " PVWordingTwo = '" & tmpUpdWording2 & "' " & _
                  " Where PVCLNNo = '" & Trim(txtPVNo) & "'"
    'MediConnCISPayment.Execute strupdatepymt
    
End Sub

Private Sub lstPaymentCLN_Click()
    
    If lstPaymentCLN.ListItems.Count Then
        CmdEdit.Visible = True
        cmdAdd.Visible = True
        CmdUpdate.Visible = False
        CmdSave.Visible = False

        ClearFramePV
        If Trim(lstPaymentCLN.SelectedItem.SubItems(17)) <> "" Then
            CmdSave.Enabled = False
            FramePV.Enabled = False
            CboBank.Enabled = False
            TxtPVNo = lstPaymentCLN.SelectedItem.SubItems(13)
            If Len(lstPaymentCLN.SelectedItem.SubItems(11)) Then
                txtPVDate = lstPaymentCLN.SelectedItem.SubItems(11)
            End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(14)) Then
                txtChequeNo = lstPaymentCLN.SelectedItem.SubItems(14)
            End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(15)) Then
                txtChequeDate = lstPaymentCLN.SelectedItem.SubItems(15)
            End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(16)) Then
                txtChequeAmt = lstPaymentCLN.SelectedItem.SubItems(16)
            End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(17)) Then
                CboACPeriodMth = Mid(lstPaymentCLN.SelectedItem.SubItems(17), 1, 2)
                CboACPeriodYr = Mid(lstPaymentCLN.SelectedItem.SubItems(17), 4, 4)
            End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(18)) Then
                CboBank = lstPaymentCLN.SelectedItem.SubItems(18)
            End If
            'If Len(lstPaymentCLN.SelectedItem.SubItems(19)) Then
            '    txtAcNo = lstPaymentCLN.SelectedItem.SubItems(19)
            'End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(20)) Then
                TxtRemark = lstPaymentCLN.SelectedItem.SubItems(20)
            End If
            
            If Len(lstPaymentCLN.SelectedItem.SubItems(22)) Then
                txtCLNEmail = lstPaymentCLN.SelectedItem.SubItems(22)
            End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(23)) Then
                txtAccNo = lstPaymentCLN.SelectedItem.SubItems(23)
            End If
            
            If Len(lstPaymentCLN.SelectedItem.SubItems(24)) Then
                CboPayto = lstPaymentCLN.SelectedItem.SubItems(24)
            End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(25)) Then
                txtAmtWording = lstPaymentCLN.SelectedItem.SubItems(25)
            End If
                        
        Else
            Call ClearFramePV
        End If
    End If
End Sub

Private Sub lstPaymentCLN_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    CmdSave.Enabled = True
    EditFlag = False
    Call ProcessCheckedItem(Item)
End Sub

Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
    If Item.Checked = True Then
        If Mid(Item.SubItems(1), 1, 1) <> "" Then
            MsgBox "Payment has been made!", vbCritical
            Item.Checked = False
        Else
            Item.SubItems(9) = CDbl(Item.SubItems(8))
            Item.SubItems(1) = "Full"
            
            LblAmt = LblAmt + CDbl(Format(Item.SubItems(9), "#,##0.00"))
        End If
    Else
        LblAmt = LblAmt - CDbl(Format(Item.SubItems(9), "#,##0.00"))
        Item.SubItems(9) = ""
        Item.SubItems(1) = ""
    End If
End Function


'*************************** Start Disable Non Integer Value
Private Sub TxtChequeAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtChequeAmt_LostFocus()
Dim tmpAmt As String
    If IsNumeric(txtChequeAmt) Then
        tmpAmt = ""
        txttmpDollar = ""
        txttmpCents = ""
        tmpAmt = ConvertCurrencyToEnglish(txtChequeAmt.Text, "MedixPayCLN")
        If Mid(tmpAmt, 1, 3) = "And" Then
            tmpAmt = Mid(tmpAmt, 5, 500)
        End If
        txtAmtWording = tmpAmt & " Only"
    End If
End Sub

Private Sub txtChequeDate_LostFocus()
    If IsDate(txtChequeDate) Then
    Else
        If txtChequeDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtChequeDate.SetFocus
        End If
    End If

End Sub

Private Sub txtChequeNo_LostFocus()
    txtChequeNo = ChkStr(txtChequeNo)
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

Private Sub TxtInvNo_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub txtInvNo_LostFocus()
    TxtInvNo = ChkStr(TxtInvNo)
    CmdSearch.Default = False
End Sub

Private Sub txtOPBatchNo1_LostFocus()
    txtOPBatchNo1 = ChkStr(txtOPBatchNo1)
End Sub
Private Sub txtOPBatchNo2_LostFocus()
    txtOPBatchNo2 = ChkStr(txtOPBatchNo2)
End Sub

Private Sub TxtPVDate1_LostFocus()
    If IsDate(txtPVDate1) Then
    Else
        If txtPVDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtPVDate2_LostFocus()
    If IsDate(txtPVDate2) Then
    Else
        If txtPVDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtPVDate_LostFocus()
    
    If IsDate(txtPVDate) Then
        CboACPeriodMth.Text = Format(txtPVDate, "mm")
        CboACPeriodYr.Text = Format(txtPVDate, "yyyy")
    Else
        If txtPVDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDate.SetFocus
        End If
    End If

End Sub

Private Sub lstPaymentCLN_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.index) = Not m_blnDirection(ColumnHeader.index)
    Select Case ColumnHeader.index
    Case 1
        SortListView lstPaymentCLN, ColumnHeader.index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstPaymentCLN, ColumnHeader.index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstPaymentCLN, ColumnHeader.index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstPaymentCLN, ColumnHeader.index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstPaymentCLN, ColumnHeader.index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstPaymentCLN, ColumnHeader.index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstPaymentCLN, ColumnHeader.index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstPaymentCLN, ColumnHeader.index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstPaymentCLN, ColumnHeader.index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstPaymentCLN, ColumnHeader.index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstPaymentCLN, ColumnHeader.index, ldtNumber, m_blnDirection(11)
    Case 12
        SortListView lstPaymentCLN, ColumnHeader.index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstPaymentCLN, ColumnHeader.index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstPaymentCLN, ColumnHeader.index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstPaymentCLN, ColumnHeader.index, ldtDateTime, m_blnDirection(15)
    Case 16
        SortListView lstPaymentCLN, ColumnHeader.index, ldtDateTime, m_blnDirection(16)
    Case 17
        SortListView lstPaymentCLN, ColumnHeader.index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstPaymentCLN, ColumnHeader.index, ldtString, m_blnDirection(18)
    Case 19
        SortListView lstPaymentCLN, ColumnHeader.index, ldtString, m_blnDirection(19)
    Case 20
        SortListView lstPaymentCLN, ColumnHeader.index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstPaymentCLN, ColumnHeader.index, ldtDateTime, m_blnDirection(21)
    Case 22
        SortListView lstPaymentCLN, ColumnHeader.index, ldtNumber, m_blnDirection(22)
    End Select
End Sub

Private Sub CboPayor_Change()
    If CboPayor.Text = "" Then
        If InStr(CboPayor.Text, "null") Then
           CboPayor.Enabled = False
        End If
    End If
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayor.ListCount - 1
 
        If NewText = LCase(Left(CboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPayor.Text = ValidValue
                CboPayor.SelStart = 0
                CboPayor.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub TxtRemark_lostfocus()
    TxtRemark = ChkStr(TxtRemark)
End Sub
Private Sub ClearFramePV()
    TxtPVNo = ""
    txtPVDate.Text = "__/__/____"
    txtChequeNo = ""
    txtChequeDate = "__/__/____"
    txtChequeAmt = ""
    CboACPeriodMth.Text = ""
    CboACPeriodYr.Text = ""
    CboBank.Text = "CB - CITIBANK"
    txtAccNo = ""
    txtCaseNo = ""
    TxtRemark = ""
    txtCLNEmail = ""
    CboPayto = ""
    txtAmtWording = ""
End Sub

Private Sub ClinicHQInfo(tmpClinicCode, tmpCLNName, tmpCLNAdd1, tmpCLNAdd2, tmpCLNAdd3, tmpPCode, tmpCity, tmpState As String)
Dim StrCLNSql As String
Dim ClnRec As New ADODB.Recordset
    StrCLNSql = "SELECT CLNCode, CLNName, CLNAdd1, CLNAdd2,CLNAdd3, CLNPostCode, CLNCity, CLNState " & _
                " FROM CLN WHERE CLNCode='" & tmpClinicCode & "'"
    ClnRec.Open StrCLNSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    Do While Not ClnRec.EOF
        With ClnRec
            tmpClinicCode = Trim(!CLNCode)
            tmpCLNName = Trim(!CLNName)
            tmpCLNAdd1 = Trim(!CLNAdd1)
            tmpCLNAdd2 = Trim(!CLNAdd2)
            tmpCLNAdd3 = Trim(!CLNAdd3)
            tmpPCode = Trim(!CLNPostCode)
            tmpCity = Trim(!CLNCity)
            tmpState = Trim(!CLNState)
        End With
        ClnRec.MoveNext
    Loop
    ClnRec.Close
    Set ClnRec = Nothing
End Sub


Private Sub PrintPV()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim strsql As String
Dim PVRec As New ADODB.Recordset
Dim Cnt As Integer
Dim tmpClinicCode As String
Dim tmpCLNName As String
Dim tmpCLNAdd1 As String
Dim tmpCLNAdd2 As String
Dim tmpCLNAdd3 As String
Dim tmpPCode As String
Dim tmpCity As String
Dim tmpState As String
Dim tmpInvoice As String
Dim tmpCLNCode As String
Dim tmpInvDate As String
Dim StrAppend As String
Dim tmpPayee As String

    Screen.MousePointer = vbHourglass
    If cmdPrintPV Then
        StrAppend = " And PVCLNNo = '" & TxtPVNo & "'"
    Else
        If lstPaymentCLN.SelectedItem.Checked = True Then
            tmpInvoice = Trim(lstPaymentCLN.SelectedItem.SubItems(4))
            tmpInvDate = Trim(lstPaymentCLN.SelectedItem.SubItems(5))
            tmpCLNCode = Trim(lstPaymentCLN.SelectedItem.SubItems(2))
            StrAppend = " AND ClinicCode='" & Trim(tmpCLNCode) & "' AND InvoiceNo = '" & Trim(tmpInvoice) & _
                        "' AND InvoiceDate='" & Format(tmpInvDate, "YYYY/MM/DD") & "'"
        End If
    End If
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If
    
    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    Call ReadRecord(StrAppend, PVRec)
    objExcel.Interactive = False

    'strsql = "select BordxRefNo, InvoiceNo, InvoiceDate, PVCLNNo,PVCLNChequeNo, PVCLNDate, CLMTotalActual, CLMMCOProcFee, " & _
    '        "ClinicCode, ClinicGrpCode, CLNName, CLNAdd1, CLNAdd2,CLNAdd3, CLNPostCode, CLNCity, CLNState, CLNPaidAmt " & _
    '        "from View_Payment2CLN_Invoice where PVCLNNo = '" & TxtPVNo & "'"
    'PV.Open strsql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
        tmpPayee = IIf(Len(Trim(PVRec!CLNPName)), Trim(PVRec!CLNPName), "")
        tmpClinicCode = Trim(PVRec!ClinicCode)
        tmpCLNName = IIf(Len(Trim(PVRec!CLNName)), Trim(PVRec!CLNName), "")
        tmpCLNAdd1 = IIf(Len(Trim(PVRec!CLNPAdd1)), Trim(PVRec!CLNPAdd1), "")
        tmpCLNAdd2 = IIf(Len(Trim(PVRec!CLNPAdd2)), Trim(PVRec!CLNPAdd2), "")
        tmpCLNAdd3 = IIf(Len(Trim(PVRec!CLNPAdd3)), Trim(PVRec!CLNPAdd3), "")
    Do While Not PVRec.EOF

        'tmpPCode = IIf(Len(Trim(PVRec!CLNPostCode)), Trim(PVRec!CLNPostCode), "")
        'tmpCity = IIf(Len(Trim(PVRec!CLNCity)), Trim(PVRec!CLNCity), "")
        'tmpState = IIf(Len(Trim(PVRec!CLNState)), Trim(PVRec!CLNState), "")
        'If Trim(PVRec!ClinicGrpCode) <> "" Then
        '    tmpClinicCode = Trim(PVRec!ClinicGrpCode)
        '    Call ClinicHQInfo(tmpClinicCode, tmpCLNName, tmpCLNAdd1, tmpCLNAdd2, tmpCLNAdd3, tmpPCode, tmpCity, tmpState)
        'End If
        If objExcel.Visible = False Then
            objExcel.Visible = True
        End If
    
        objExcel.WindowState = xlMaximized
        Cnt = Cnt + 1
        If Cnt = 1 Then
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\PV-Clinic.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        objWorksheet.Cells(8, "C") = tmpPayee
        objWorksheet.Cells(9, "C") = tmpCLNName & " (" & tmpClinicCode & ")"
        objWorksheet.Cells(10, "C") = tmpCLNAdd1
        objWorksheet.Cells(11, "C") = tmpCLNAdd2
        objWorksheet.Cells(12, "C") = tmpCLNAdd3
        objWorksheet.Cells(8, "H") = TxtPVNo '& "/" & Format(PVRec!PVCLNDate, "MM/YY")
        objWorksheet.Cells(9, "H") = PVRec!PVCLNChequeNo
        objWorksheet.Cells(10, "H") = PVRec!PVCLNDate
        objWorksheet.Cells(37, "C") = Logon
         
        objWorksheet.Cells(15 + CDbl(Cnt), "A") = Cnt
        objWorksheet.Cells(15 + CDbl(Cnt), "B") = IIf(Len(PVRec!InvoiceNo), Trim(PVRec!InvoiceNo), "")
        objWorksheet.Cells(15 + CDbl(Cnt), "C") = PVRec!InvoiceDate
        objWorksheet.Cells(15 + CDbl(Cnt), "D") = IIf(Len(PVRec!BordxRefNo), PVRec!BordxRefNo, "")
        objWorksheet.Cells(15 + CDbl(Cnt), "F") = PVRec!CLMFinalApprove
        objWorksheet.Cells(15 + CDbl(Cnt), "G") = IIf(IsNumeric(PVRec!CLMMCOProcFee), PVRec!CLMMCOProcFee, 0)
        If cmdPrintPV Then
            objWorksheet.Cells(15 + CDbl(Cnt), "H") = PVRec!CLNPaidAmt
        Else
            objWorksheet.Cells(15 + CDbl(Cnt), "H") = PVRec!CLMPayableByMedix2C
        End If
        PVRec.MoveNext
    Loop
    PVRec.Close
    Set PVRec = Nothing
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    Screen.MousePointer = vbDefault
    Exit Sub
ExitFunction:
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            objExcel.Quit
            Screen.MousePointer = vbDefault
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            Screen.MousePointer = vbDefault
            GoTo ExitFunction
End Sub

Private Sub ReadRecord(StrAppend As String, PVRec As ADODB.Recordset)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter

    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "SP_CISPayment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("PaymentType", adVarChar, adParamInput, 10, "Pay2CLN")
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, StrAppend)
    cmd.Parameters.Append prm2
    PVRec.CursorLocation = adUseClient
    PVRec.CursorType = adOpenForwardOnly
    PVRec.CacheSize = 100
    Set PVRec = cmd.Execute
End Sub


