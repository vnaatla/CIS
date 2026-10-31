VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPaymentFFS 
   Caption         =   "FFS Payment"
   ClientHeight    =   7992
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7992
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "Check All"
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
      TabIndex        =   81
      Top             =   1880
      Width           =   855
   End
   Begin MSComctlLib.ListView LstFFSPmt 
      Height          =   3735
      Left            =   45
      TabIndex        =   61
      Top             =   2160
      Width           =   11730
      _ExtentX        =   20680
      _ExtentY        =   6583
      SortKey         =   12
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
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   25
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Pay Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "BordxNo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "BordxDate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "BordxAmt"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "ClmFFS"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "PaidFFS"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "OS Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Amt Paying"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "GroupCompany"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "PVNo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "PVDate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Chq No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "ChqDate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "ChqAmt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Mty/Yr"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Bank"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "AccNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Remarks"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "PaymentIndex"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Payee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "RecNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "RecDate"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame PVframe 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   50
      TabIndex        =   46
      Top             =   5880
      Width           =   11725
      Begin VB.ComboBox CboPayee 
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
         ItemData        =   "FrmPaymentFFS.frx":0000
         Left            =   8640
         List            =   "FrmPaymentFFS.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   195
         Width           =   2895
      End
      Begin VB.TextBox txtPayIndex 
         Height          =   285
         Left            =   8040
         TabIndex        =   79
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBordxNo3 
         Height          =   285
         Left            =   8040
         TabIndex        =   78
         Top             =   840
         Visible         =   0   'False
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
         Height          =   590
         Left            =   8640
         TabIndex        =   25
         Top             =   480
         Width           =   2895
      End
      Begin VB.ComboBox CboBank 
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
         Left            =   5145
         Sorted          =   -1  'True
         TabIndex        =   20
         Text            =   "MB - MAYBANK"
         Top             =   200
         Width           =   2800
      End
      Begin VB.TextBox TxtChqNo 
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
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   17
         Top             =   200
         Width           =   1200
      End
      Begin VB.TextBox TxtPVNo3 
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
         MaxLength       =   30
         TabIndex        =   15
         Top             =   200
         Width           =   1200
      End
      Begin VB.TextBox TxtACNo 
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
         Left            =   5145
         MaxLength       =   30
         TabIndex        =   21
         Top             =   480
         Width           =   2800
      End
      Begin VB.ComboBox CboACMth3 
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
         Left            =   5145
         TabIndex        =   22
         Top             =   760
         Width           =   1000
      End
      Begin VB.ComboBox CboACYr3 
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
         Left            =   6960
         TabIndex        =   23
         Top             =   760
         Width           =   1000
      End
      Begin MSMask.MaskEdBox TxtPVDate3 
         Height          =   315
         Left            =   800
         TabIndex        =   16
         Top             =   480
         Width           =   1200
         _ExtentX        =   2117
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
      Begin MSMask.MaskEdBox TxtChqDate 
         Height          =   315
         Left            =   3000
         TabIndex        =   18
         Top             =   480
         Width           =   1200
         _ExtentX        =   2117
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
      Begin VB.TextBox TxtChqAmt 
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
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   19
         Top             =   760
         Width           =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "Payee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8040
         TabIndex        =   80
         Top             =   225
         Width           =   615
      End
      Begin VB.Label LblRemark 
         Caption         =   "Remark"
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
         Left            =   8040
         TabIndex        =   71
         Top             =   510
         Width           =   735
      End
      Begin VB.Label LblACy 
         Caption         =   "A/C Yr"
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
         Left            =   6360
         TabIndex        =   70
         Top             =   825
         Width           =   855
      End
      Begin VB.Label LblACm 
         Caption         =   "A/C Mth"
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
         Left            =   4395
         TabIndex        =   69
         Top             =   825
         Width           =   735
      End
      Begin VB.Label LblACn 
         Caption         =   "A/C No"
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
         Left            =   4395
         TabIndex        =   68
         Top             =   510
         Width           =   735
      End
      Begin VB.Label LblBank 
         Caption         =   "Bank"
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
         Left            =   4395
         TabIndex        =   67
         Top             =   225
         Width           =   735
      End
      Begin VB.Label LblChqA 
         Caption         =   "Chq. Amt"
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
         Left            =   2100
         TabIndex        =   66
         Top             =   800
         Width           =   735
      End
      Begin VB.Label LblChqD 
         Caption         =   "Chq. Date"
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
         Left            =   2100
         TabIndex        =   65
         Top             =   510
         Width           =   795
      End
      Begin VB.Label LblChqN 
         Caption         =   "Chq. No"
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
         Left            =   2100
         TabIndex        =   64
         Top             =   225
         Width           =   735
      End
      Begin VB.Label LblPVd 
         Caption         =   "PV Date"
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
         Left            =   120
         TabIndex        =   63
         Top             =   510
         Width           =   615
      End
      Begin VB.Label LblPVn 
         Caption         =   "PV No"
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
         Left            =   120
         TabIndex        =   62
         Top             =   225
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   45
      TabIndex        =   42
      Top             =   240
      Width           =   11725
      Begin VB.CheckBox ChkBordxDate 
         Caption         =   "Check5"
         Height          =   255
         Left            =   5200
         TabIndex        =   40
         Top             =   520
         Width           =   255
      End
      Begin VB.CheckBox ChkBordxNo 
         Caption         =   "Check1"
         Height          =   255
         Left            =   5200
         TabIndex        =   77
         Top             =   210
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox TxtPVNo1 
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
         Left            =   7700
         MaxLength       =   30
         TabIndex        =   6
         Top             =   240
         Width           =   1500
      End
      Begin VB.CheckBox ChkACYr 
         Caption         =   "Check4"
         Height          =   255
         Left            =   11350
         TabIndex        =   14
         Top             =   1110
         Width           =   255
      End
      Begin VB.CheckBox ChkACMth 
         Caption         =   "Check8"
         Height          =   255
         Left            =   11350
         TabIndex        =   37
         Top             =   840
         Width           =   210
      End
      Begin VB.CheckBox ChkPVDate 
         Caption         =   "Check7"
         Height          =   255
         Left            =   11350
         TabIndex        =   38
         Top             =   555
         Width           =   255
      End
      Begin VB.CheckBox ChkPVNo 
         Caption         =   "Check6"
         Height          =   255
         Left            =   11350
         TabIndex        =   39
         Top             =   270
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpComp 
         Caption         =   "Check2"
         Height          =   255
         Left            =   5200
         TabIndex        =   41
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox TxtBordxNo1 
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
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   0
         Top             =   200
         Width           =   3525
      End
      Begin VB.TextBox TxtPVNo2 
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
         Left            =   9700
         MaxLength       =   30
         TabIndex        =   7
         Top             =   240
         Width           =   1500
      End
      Begin MSMask.MaskEdBox TxtPVDate1 
         Height          =   315
         Left            =   7695
         TabIndex        =   8
         Top             =   525
         Width           =   1500
         _ExtentX        =   2646
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
      Begin MSMask.MaskEdBox TxtPVDate2 
         Height          =   315
         Left            =   9705
         TabIndex        =   9
         Top             =   525
         Width           =   1500
         _ExtentX        =   2646
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
      Begin VB.ComboBox CboACMth2 
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
         Left            =   9705
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   795
         Width           =   1500
      End
      Begin VB.ComboBox CboACMth1 
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
         Left            =   7700
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   795
         Width           =   1500
      End
      Begin VB.ComboBox CboACYr2 
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
         Left            =   9700
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1065
         Width           =   1500
      End
      Begin VB.ComboBox CboACYr1 
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
         Left            =   7700
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1065
         Width           =   1500
      End
      Begin VB.TextBox TxtBordxNo2 
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
         Left            =   3580
         MaxLength       =   30
         TabIndex        =   1
         Top             =   200
         Width           =   1500
      End
      Begin MSMask.MaskEdBox TxtBordxDate1 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   1500
         _ExtentX        =   2646
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
      Begin MSMask.MaskEdBox TxtBodxDate2 
         Height          =   315
         Left            =   3585
         TabIndex        =   3
         Top             =   480
         Width           =   1500
         _ExtentX        =   2646
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
      Begin VB.ComboBox CboPmtStat 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   760
         Width           =   3540
      End
      Begin VB.ComboBox CboGComp 
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
         ItemData        =   "FrmPaymentFFS.frx":0004
         Left            =   1560
         List            =   "FrmPaymentFFS.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1040
         Width           =   3540
      End
      Begin VB.Label LblACYr 
         Caption         =   "A/C Period Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6105
         TabIndex        =   60
         Top             =   1125
         Width           =   1500
      End
      Begin VB.Label LblPmtStatStric 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5200
         TabIndex        =   59
         Top             =   790
         Width           =   180
      End
      Begin VB.Label Label18 
         Caption         =   "To"
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
         Left            =   9345
         TabIndex        =   58
         Top             =   855
         Width           =   270
      End
      Begin VB.Label Label17 
         Caption         =   "To"
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
         Left            =   9345
         TabIndex        =   57
         Top             =   1155
         Width           =   270
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Left            =   9345
         TabIndex        =   56
         Top             =   555
         Width           =   270
      End
      Begin VB.Label Label15 
         Caption         =   "To"
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
         Left            =   9345
         TabIndex        =   55
         Top             =   255
         Width           =   270
      End
      Begin VB.Label Label14 
         Caption         =   "To"
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
         Left            =   3200
         TabIndex        =   54
         Top             =   530
         Width           =   270
      End
      Begin VB.Label Label12 
         Caption         =   "To"
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
         Left            =   3200
         TabIndex        =   53
         Top             =   240
         Width           =   255
      End
      Begin VB.Label LblACMth 
         Caption         =   "A/C Period Month"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6105
         TabIndex        =   52
         Top             =   825
         Width           =   1500
      End
      Begin VB.Label LblPVDate 
         Caption         =   "PV Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6105
         TabIndex        =   51
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label LblPVNo 
         Caption         =   "PV Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6105
         TabIndex        =   50
         Top             =   255
         Width           =   1500
      End
      Begin VB.Label LblBordxDate 
         Caption         =   "Bordx Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   49
         Top             =   500
         Width           =   1500
      End
      Begin VB.Label LblPmtStat 
         Caption         =   "Payment Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   45
         Top             =   780
         Width           =   1455
      End
      Begin VB.Label LblGComp 
         Caption         =   "Group Company"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   44
         Top             =   1070
         Width           =   1455
      End
      Begin VB.Label LblBordxNo 
         Caption         =   "Bordx No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   43
         Top             =   200
         Width           =   1450
      End
   End
   Begin VB.Frame Frame3 
      Height          =   600
      Left            =   50
      TabIndex        =   47
      Top             =   7080
      Width           =   11725
      Begin VB.TextBox LblAmtR 
         Height          =   285
         Left            =   4200
         TabIndex        =   85
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox lblTotalAmt 
         Height          =   285
         Left            =   2160
         TabIndex        =   84
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox LblRecordR 
         Height          =   285
         Left            =   720
         TabIndex        =   83
         Top             =   225
         Width           =   645
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   11160
         TabIndex        =   36
         Top             =   180
         Width           =   500
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   10620
         TabIndex        =   35
         Top             =   180
         Width           =   550
      End
      Begin VB.CommandButton CmdPrinttoFile 
         Caption         =   "Print to &File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9565
         TabIndex        =   34
         Top             =   180
         Width           =   1065
      End
      Begin VB.CommandButton CmdPrintPV 
         Caption         =   "Print P&V"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   8785
         TabIndex        =   33
         Top             =   180
         Width           =   800
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   8200
         TabIndex        =   32
         Top             =   180
         Width           =   600
      End
      Begin VB.CommandButton CmdSearch 
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
         Height          =   350
         Left            =   7510
         TabIndex        =   31
         Top             =   180
         Width           =   700
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6780
         TabIndex        =   28
         Top             =   180
         Width           =   750
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6780
         TabIndex        =   30
         Top             =   180
         Width           =   750
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6200
         TabIndex        =   27
         Top             =   180
         Width           =   600
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6200
         TabIndex        =   29
         Top             =   180
         Width           =   600
      End
      Begin VB.CommandButton CmdPreviewPV 
         Caption         =   "Previe&w"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5440
         TabIndex        =   26
         Top             =   180
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Total Amt"
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
         Left            =   1440
         TabIndex        =   82
         Top             =   255
         Width           =   735
      End
      Begin VB.Label LblAmt 
         Caption         =   "Chk Amt"
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
         TabIndex        =   73
         Top             =   255
         Width           =   615
      End
      Begin VB.Label LblRecord 
         Caption         =   "Records"
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
         TabIndex        =   72
         Top             =   255
         Width           =   735
      End
   End
   Begin VB.Label Label36 
      Caption         =   "* Listing"
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
      Left            =   4920
      TabIndex        =   76
      Top             =   7680
      Width           =   2205
   End
   Begin VB.Label Label35 
      Caption         =   "* - Search Empty Data"
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
      Height          =   255
      Left            =   2280
      TabIndex        =   75
      Top             =   7680
      Width           =   2205
   End
   Begin VB.Label Label34 
      Caption         =   "Bold - Searchable"
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
      Left            =   120
      TabIndex        =   74
      Top             =   7680
      Width           =   2205
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "FFS Payment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmPaymentFFS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intExit As Boolean
Dim EditFlag As Boolean
Dim m_blnDirection(1 To 22) As Boolean

Private Sub CboACMth1_KeyPress(KeyAscii As Integer)
    Call ComboKeyPressNum(CboACMth1, KeyAscii)
End Sub

Private Sub CboACMth3_KeyPress(KeyAscii As Integer)
    Call ComboKeyPressNum(CboACMth3, KeyAscii)
End Sub

Private Sub CboACYr1_KeyPress(KeyAscii As Integer)
    Call ComboKeyPressNum(CboACYr1, KeyAscii)
End Sub

Private Sub CboACYr2_KeyPress(KeyAscii As Integer)
    Call ComboKeyPressNum(CboACYr2, KeyAscii)
End Sub

Private Sub CboACYr3_KeyPress(KeyAscii As Integer)
    Call ComboKeyPressNum(CboACYr3, KeyAscii)
End Sub

Private Sub CboBank_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboBank, KeyAscii)
End Sub

Private Sub CboGComp_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboGComp, KeyAscii)
End Sub

Private Sub CboPayee_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboPayee, KeyAscii)
End Sub

Private Sub CboPmtStat_Keypress(KeyAscii As Integer)
    Call ComboKeyPress(CboPmtStat, KeyAscii)
End Sub

Private Sub ChkBordxNo_Click()
    LstFFSPmt.ListItems.Clear
    If ChkBordxNo.Value = 1 Then
        Label12.Visible = False
        TxtBordxNo1.Width = 3540
        TxtBordxNo2.Visible = False
        TxtBordxNo1 = ""
    Else
        Label12.Visible = True
        TxtBordxNo1.Width = 1500
        TxtBordxNo2.Visible = True
    End If
    TxtBordxNo1.SetFocus
End Sub


Private Sub ChkGrpComp_Click()
Const METHOD_NAME = "ChkGrpComp_Click"
On Error GoTo debugs


Dim TempRec As New ADODB.Recordset
Dim TempStr As String
    
    If ChkGrpComp.Value = 1 Then
        CboGComp.Clear
            TempStr = "select GRPCoName from GRPCompany"
            TempRec.Open TempStr, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            Do While Not TempRec.EOF
                CboGComp.AddItem IIf(Len(TempRec!GrpCoName), TempRec!GrpCoName, "")
                TempRec.MoveNext
            Loop
            TempRec.Close
            Set TempRec = Nothing
    Else
        CboGComp.Clear
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim StrSql As String
Dim strAppend As String
Dim sql As String
Dim GrpCompanyName As String
    
    strAppend = ""
    GrpCompanyName = ""
    
    If Len(Trim(CboGComp)) Then
         GrpCompanyName = Replace(CboGComp, "'", "''")
         strAppend = strAppend & " AND GroupCompany = '" & Trim(GrpCompanyName) & "'"
    End If
    
    If Mid(CboPmtStat, 1, 1) = 1 Then
        strAppend = strAppend & " and PVPayStatus = 'F'"
    End If
 
    If Mid(CboPmtStat, 1, 1) = 2 Then
        strAppend = strAppend & " and (PVPayStatus = '' or PVPayStatus is null) "
    End If
 
    If ChkBordxNo.Value = 1 Then
        strAppend = strAppend & " AND BordxRefNo = '" & Trim(TxtBordxNo1) & "'"
    End If
        If Trim(TxtBordxNo1) <> "" Then
            strAppend = strAppend & " and BordxRefNo >= '" & Trim(TxtBordxNo1) & "'"
        End If
        If Trim(TxtBordxNo2) <> "" Then
            strAppend = strAppend & " and BordxRefNo <= '" & Trim(TxtBordxNo2) & "'"
        End If
    
    If ChkBordxDate.Value = 1 Then
        sql = ""
        sql = "AND (BordxDate Is null OR BordxDate = '') " & ""
        strAppend = strAppend & sql
    End If
        If IsDate(TxtBordxDate1) = True Then
            strAppend = strAppend & " And BordxDate >= '" & Format(TxtBordxDate1, "mm/dd/yyyy") & "'"
        End If
        If IsDate(TxtBodxDate2) = True Then
            strAppend = strAppend & " And BordxDate <= '" & Format(TxtBodxDate2, "mm/dd/yyyy") & "'"
        End If
    
    If ChkPVNo.Value = 1 Then
        sql = ""
        sql = "AND (PVNo Is null OR PVNo = '') " & ""
        strAppend = strAppend & sql
    End If
    
    If Len(TxtPVNo1) Then
        strAppend = strAppend & " and PVNo >= '" & Trim(TxtPVNo1) & "'"
    End If
    If Len(TxtPVNo2) Then
        strAppend = strAppend & " and PVNo <= '" & Trim(TxtPVNo2) & "'"
    End If
    
    If ChkPVDate.Value = 1 Then
        sql = ""
        sql = "AND (PVDate Is null OR PVDate = '') " & ""
        strAppend = strAppend & sql
    End If
    If IsDate(TxtPVDate1) = True Then
        strAppend = strAppend & " And PVDate >= '" & Format(TxtPVDate1, "mm/dd/yyyy") & "'"
    End If
    If IsDate(TxtPVDate2) = True Then
        strAppend = strAppend & " And PVDate <= '" & Format(TxtPVDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkACMth.Value = 1 Then
        sql = ""
        sql = " AND FFSAccMth Is null "
        strAppend = strAppend & sql
    End If
    If Len(CboACMth1) Then
        strAppend = strAppend & " and FFSAccMth >= '" & Trim(CboACMth1) & "'"
    End If
    If Len(CboACMth2) Then
        strAppend = strAppend & " and FFSAccMth <= '" & Trim(CboACMth2) & "'"
    End If
    
    If ChkACYr.Value = 1 Then
        sql = ""
        sql = " AND FFSAccYr Is null "
        strAppend = strAppend & sql
    End If
    If Len(CboACYr1) Then
        strAppend = strAppend & " and FFSAccYr >= '" & Trim(CboACYr1) & "'"
    End If
    If Len(CboACYr2) Then
        strAppend = strAppend & " and FFSAccYr <= '" & Trim(CboACYr2) & "'"
    End If
    
    Call PaymentFFSListing(strAppend)
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Sub CmdAdd_Click()
Dim ZZ As Integer
Dim ItemChk As Boolean
    
    ItemChk = False
    For ZZ = 1 To LstFFSPmt.ListItems.Count
        If LstFFSPmt.ListItems(ZZ).Checked = True Then
            ItemChk = True
            Exit For
        End If
    Next ZZ

    CmdUpdate.Enabled = False
    
    If ItemChk Then
        Call ClearFramePV
        CmdAdd.Visible = False
        CmdSave.Visible = True
        CmdSave.Enabled = True
        PVframe.Enabled = True
        CboBank.Enabled = True
        TxtChqAmt = CDbl(LblAmtR)
        TxtPVDate3.SetFocus
    Else
        MsgBox "Please checked at least 1 Claim  ! ", vbCritical + vbOKOnly, DialogBoxTitle
    End If
    'LstFFSPmt.Enabled = False
End Sub


Private Sub cmdCheckAll_Click()
Dim LstCnt As Integer
    If LstFFSPmt.ListItems.Count <> 0 Then
        LstCnt = LstFFSPmt.ListItems.Count
    
        For LstCnt = 1 To LstCnt
            If LstFFSPmt.ListItems(LstCnt).Checked = False Then
                LstFFSPmt.ListItems(LstCnt).Checked = True
                LstFFSPmt.ListItems(LstCnt).SubItems(1) = "Full"
                LstFFSPmt.ListItems(LstCnt).SubItems(9) = LstFFSPmt.ListItems(LstCnt).SubItems(8)
                LstFFSPmt.ListItems(LstCnt).SubItems(8) = 0
                LblAmtR = Format(LblAmtR + CDbl(LstFFSPmt.ListItems(LstCnt).SubItems(9)), "#,##0.00")
            ElseIf LstFFSPmt.ListItems(LstCnt).Checked = True Then
                LblAmtR = Format(LblAmtR - CDbl(LstFFSPmt.ListItems(LstCnt).SubItems(9)), "#,##0.00")
                LstFFSPmt.ListItems(LstCnt).Checked = False
                LstFFSPmt.ListItems(LstCnt).SubItems(1) = ""
                LstFFSPmt.ListItems(LstCnt).SubItems(8) = LstFFSPmt.ListItems(LstCnt).SubItems(9)
                LstFFSPmt.ListItems(LstCnt).SubItems(9) = 0
            End If
        Next LstCnt
    Else
        MsgBox "Please select a record.", vbInformation
    End If

End Sub

Private Sub cmdClear_Click()
Dim IntCount As Integer

    Call ClearForm(Me)
    CmdSave.Visible = False
    CmdAdd.Visible = True
    CmdEdit.Visible = True
    CmdUpdate.Visible = False
    PVframe.Enabled = False
    EditFlag = False
    LstFFSPmt.ListItems.Clear
    CboPmtStat.ListIndex = 0
    LblRecordR = 0
    LblAmtR = 0
    TxtBordxNo1.SetFocus
    ChkBordxNo.Value = 1
    Label12.Visible = False
    TxtBordxNo1.Width = 3540
    TxtBordxNo2.Visible = False
    LstFFSPmt.Enabled = True
    'CboPayee.Text = "UNITED HEALTH CONNECT SDN BHD"
    CboPayee.Text = "HEALTH CONNECT SDN BHD"
    CboACMth1.Clear
    CboACMth2.Clear
    CboACMth3.Clear
    For IntCount = 1 To 12
        CboACMth1.AddItem Format(IntCount, "00")
        CboACMth2.AddItem Format(IntCount, "00")
        CboACMth3.AddItem Format(IntCount, "00")
    Next IntCount
    CboACYr1.Clear
    CboACYr2.Clear
    CboACYr3.Clear
    For IntCount = 1999 To Year(Date) + 2
        CboACYr1.AddItem Format(IntCount, "00")
        CboACYr2.AddItem Format(IntCount, "00")
        CboACYr3.AddItem Format(IntCount, "00")
    Next IntCount

End Sub

Private Sub CmdEdit_Click()
Const METHOD_NAME = "CmdEdit_Click"
On Error GoTo debugs

Dim ZZ As Integer
Dim ItemChk As Boolean
Dim EditFlag As Boolean

    EditFlag = True
    ItemChk = False
    For ZZ = 1 To LstFFSPmt.ListItems.Count
        If LstFFSPmt.ListItems(ZZ).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next ZZ

    CmdSave.Enabled = False
    PVframe.Enabled = False
    If ItemChk = False And Len(Trim(TxtPVNo3)) Then
        CmdEdit.Visible = False
        CmdUpdate.Visible = True
        CmdUpdate.Enabled = True
        PVframe.Enabled = True
        EditFlag = True
        If Len(Trim(LstFFSPmt.SelectedItem.SubItems(1))) Then
            PVframe.Enabled = True
        End If
        TxtPVDate3.SetFocus
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub cmdExit_Click()
    intExit = True
    Unload Me
End Sub

Private Sub CmdPreviewPV_Click()
Const METHOD_NAME = "CmdPreviewPV_Click"
On Error GoTo debugs

Dim ItemChecked As Boolean
Dim EE As Integer
    ItemChecked = False
    For EE = 1 To LstFFSPmt.ListItems.Count
        If LstFFSPmt.ListItems(EE).Checked = True Then
            ItemChecked = True
        Exit For
        End If
    Next EE
    
    If ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    Else
            Call PrintPV
    End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdPrintPV_Click()
Const METHOD_NAME = "CmdPrintPV_Click"
On Error GoTo debugs


    If TxtPVNo3 = "" Then
        MsgBox "Please Enter or Generate PV No", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
    Else
        Call PrintPV
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdPrinttoFile_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelAnalysis1(LstFFSPmt)
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSearch_Click()
Const METHOD_NAME = "CmdSearch_Click"
On Error GoTo debugs

Dim AA As Integer
Dim ChkCas As Boolean

    LblAmtR = 0
    lblTotalAmt = 0
    If ChkBordxNo.Value = 1 Then
        If Trim(TxtBordxNo1) = "" Then
            MsgBox "Please Enter Bordx Ref. No. ", vbOKOnly + vbCritical, DialogBoxTitle
            TxtBordxNo1.SetFocus
        Else
            ChkCas = False
            Screen.MousePointer = vbHourglass
            If LstFFSPmt.ListItems.Count <> 0 Then
                For AA = 1 To LstFFSPmt.ListItems.Count
                   If Trim(ChkStr(TxtBordxNo1)) = Trim(LstFFSPmt.ListItems(AA).SubItems(2)) Then
                        ChkCas = True
                        Exit For
                   End If
                Next AA
            End If
            If ChkCas = True Then
                MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
            ElseIf ChkCas = False Then
                Call SearchRecord
            End If
            ChkCas = False
            Screen.MousePointer = Default
            TxtBordxNo1 = ""
            TxtBordxNo1.SetFocus
        End If
        TxtBordxNo1.SetFocus
    Else
       Call SearchRecord
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdSave_Click()
Const METHOD_NAME = "CmdSave_Click"
On Error GoTo debugs


Dim tmpBordxNo As String
Dim i As Integer
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
Dim tmpFFSAmt As Double
Dim PVNumber As String
Dim PrevGRP, CurrGRP As String
Dim DiffGRP As Boolean
Dim CLNCnt As Integer
Dim CompTrans As Boolean

    ItemChecked = False
    TmpPVCode = ""
    ZZ = InStr(1, CboBank, "-")
    
        TempBank = Trim(Left(CboBank, IIf(ZZ - 1 <= 0, Len(Trim(CboBank)), ZZ - 1)))
        BankSql = "SELECT PVCode from Banks WHERE BankCode='" & Trim(TempBank) & "'"
        BankRec.Open BankSql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
        If BankRec.EOF = False Then
            TmpPVCode = Trim(BankRec!PVCode)
        End If
        BankRec.Close
        Set BankRec = Nothing
    
    For i = 1 To LstFFSPmt.ListItems.Count
        If LstFFSPmt.ListItems(i).Checked = True Then
            ItemChecked = True
            Exit For
        End If
    Next i
    CLNCnt = 0
    DiffGRP = False
    
    For i = 1 To LstFFSPmt.ListItems.Count
        If LstFFSPmt.ListItems(i).Checked = True Then
            CLNCnt = CLNCnt + 1
            CurrGRP = Trim(LstFFSPmt.ListItems(i).SubItems(11))
            If CLNCnt = 1 Then
                PrevGRP = Trim(LstFFSPmt.ListItems(i).SubItems(11))
            End If
            If PrevGRP <> CurrGRP Then
                DiffGRP = True
                Exit For
            End If
        End If
    Next i

    If ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    ElseIf IsDate(TxtPVDate3) = False Then
        MsgBox "Invalid PV Date!", vbCritical
'        txtPVDate.SetFocus
    ElseIf TxtPVDate3 = "__/__/____" Then
        MsgBox "Please Enter PV Date! ", vbCritical
        TxtPVDate3.SetFocus
    'ElseIf DiffGRP Then
    '    MsgBox "Different Group Companies found !", vbOKOnly + vbCritical
    ElseIf TxtChqAmt = "" Then
        MsgBox "Please Enter Cheque Amt!", vbCritical
        TxtChqAmt.SetFocus
    ElseIf CDbl(TxtChqAmt) <> CDbl(LblAmtR) And TxtRemark = "" Then
        MsgBox "Amount to paid is incorrect" & vbNewLine & "Please key in remark for reason.", vbCritical + vbOKOnly
        TxtRemark.SetFocus
    ElseIf Trim(CboBank) = "" Then
        MsgBox "Please Enter Bank !", vbCritical
        CboBank.SetFocus
    ElseIf Trim(CboACMth3) = "" Then
        MsgBox "Please keyin Account Period Month ! ", vbCritical
        CboACMth3.SetFocus
    ElseIf Trim(CboACYr3) = "" Then
        MsgBox "Please keyin Account Period Year ! ", vbCritical
        CboACYr3.SetFocus
    ElseIf Trim(TmpPVCode) = "" Then
        MsgBox "Please keyin Bank - PV Code ! ", vbCritical
    ElseIf Trim(CboPayee) = "" Then
        MsgBox "Please keyin Payee Name ! ", vbCritical
        CboPayee.SetFocus
    Else
        Update = MsgBox("Do you want to update the record ? ", vbYesNo + vbExclamation)
        If Update = vbYes Then
            CompTrans = False
            MediConnCISPayment.beginTrans
            
            PVNumber = GenerateSPVNo(TempBank, TmpPVCode)
            TxtPVNo3 = ""
            TxtPVNo3 = PVNumber
           
            For i = 1 To LstFFSPmt.ListItems.Count
                If LstFFSPmt.ListItems(i).Checked = True Then
                    tmpBordxNo = Trim(LstFFSPmt.ListItems(i).SubItems(2))
                    tmpFFSAmt = IIf(IsNumeric(LstFFSPmt.ListItems(i).SubItems(9)), LstFFSPmt.ListItems(i).SubItems(9), 0)
                    PymtIndex = ""
                    Call UpdateRecord(tmpBordxNo, tmpFFSAmt, PymtIndex)
                End If
            Next i
            
            MediConnCISPayment.commitTrans
            CompTrans = True
            CmdSave.Visible = False
            CmdAdd.Visible = True
            LstFFSPmt.ListItems.Clear
            LblAmtR = 0
            PVSql = " and PVNo = '" & Trim(TxtPVNo3) & "'"
            Call ClearFramePV
            LstFFSPmt.ListItems.Clear
            Call PaymentFFSListing(PVSql)
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            
        End If
    End If
    LstFFSPmt.Enabled = True
        
        
ErrorSave:
        Screen.MousePointer = vbDefault
        If CompTrans = False Then
            MediConnCISPayment.RollbackTrans
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

Private Sub CmdStop_Click()
    intExit = 1
    TxtBordxNo1.SetFocus
    CmdExit.Enabled = True
End Sub

Private Sub CmdUpdate_Click()
Const METHOD_NAME = "CmdUpdate_Click"
On Error GoTo debugs


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
    For ZZ = 1 To LstFFSPmt.ListItems.Count
        If LstFFSPmt.ListItems(ZZ).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next ZZ
    If ItemChk = True Then
    ElseIf LstFFSPmt.ListItems.Count = 0 Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    ElseIf TxtPVNo3 = "" Then
        MsgBox "Please Confirm PV No!", vbCritical
    ElseIf IsDate(TxtPVDate3) = False Then
        MsgBox "Invalid PV Date!", vbCritical
        TxtPVDate3.SetFocus
    ElseIf TxtPVDate3 = "__/__/____" Then
        MsgBox "Please Enter PV Date! ", vbCritical
        TxtPVDate3.SetFocus
    ElseIf TxtChqAmt = "" Then
        MsgBox "Please Enter Cheque Amt!", vbCritical
        TxtChqAmt.SetFocus
    ElseIf Trim(CboBank) = "" Then
        MsgBox "Please Enter Bank !", vbCritical
        CboBank.SetFocus
    ElseIf Trim(CboACMth3) = "" Then
        MsgBox "Please keyin Account Period Month ! ", vbCritical
        CboACMth3.SetFocus
    ElseIf Trim(CboACYr3) = "" Then
        MsgBox "Please keyin Account Period Year ! ", vbCritical
        CboACYr3.SetFocus
    Else
        RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            Call UpdateRecord(txtBordxNo3, CDbl(TxtChqAmt), txtPayIndex)
            CmdSave.Enabled = True
            EditFlag = False
            LblAmt = 0
            PVSql = " and PVNo = '" & Trim(TxtPVNo3) & "'"
            Call ClearFramePV
            LstFFSPmt.ListItems.Clear
            Call PaymentFFSListing(PVSql)
        End If
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "FrmPaymentFFS")
    Call ComboListing
End Sub

Private Sub LstFFSPmt_Click()
Const METHOD_NAME = "LstFFSPmt_Click"
On Error GoTo debugs

    If LstFFSPmt.ListItems.Count Then
        If Trim(LstFFSPmt.SelectedItem.SubItems(12)) <> "" Then
            CmdSave.Enabled = False
            PVframe.Enabled = False
            CboBank.Enabled = False
            CmdEdit.Visible = True
            TxtPVNo3 = LstFFSPmt.SelectedItem.SubItems(12)
            If Len(LstFFSPmt.SelectedItem.SubItems(2)) Then
                txtBordxNo3 = LstFFSPmt.SelectedItem.SubItems(2)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(13)) Then
                TxtPVDate3 = LstFFSPmt.SelectedItem.SubItems(13)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(14)) Then
                TxtChqNo = LstFFSPmt.SelectedItem.SubItems(14)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(15)) Then
                TxtChqDate = LstFFSPmt.SelectedItem.SubItems(15)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(16)) Then
                TxtChqAmt = LstFFSPmt.SelectedItem.SubItems(16)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(17)) Then
                CboACMth3 = Mid(LstFFSPmt.SelectedItem.SubItems(17), 1, 2)
                CboACYr3 = Mid(LstFFSPmt.SelectedItem.SubItems(17), 4, 4)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(18)) Then
                CboBank.Text = LstFFSPmt.SelectedItem.SubItems(18)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(19)) Then
                TxtACNo = LstFFSPmt.SelectedItem.SubItems(19)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(20)) Then
                TxtRemark = LstFFSPmt.SelectedItem.SubItems(20)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(21)) Then
                txtPayIndex = LstFFSPmt.SelectedItem.SubItems(21)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(22)) Then
                CboPayee = LstFFSPmt.SelectedItem.SubItems(22)
            End If
        Else
            Call ClearFramePV
        End If
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub LstFFSPmt_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView LstFFSPmt, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView LstFFSPmt, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView LstFFSPmt, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView LstFFSPmt, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView LstFFSPmt, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView LstFFSPmt, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView LstFFSPmt, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView LstFFSPmt, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView LstFFSPmt, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView LstFFSPmt, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView LstFFSPmt, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
    Case 12
        SortListView LstFFSPmt, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView LstFFSPmt, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView LstFFSPmt, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView LstFFSPmt, ColumnHeader.Index, ldtDateTime, m_blnDirection(15)
    Case 16
        SortListView LstFFSPmt, ColumnHeader.Index, ldtDateTime, m_blnDirection(16)
    Case 17
        SortListView LstFFSPmt, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView LstFFSPmt, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 19
        SortListView LstFFSPmt, ColumnHeader.Index, ldtString, m_blnDirection(19)
    Case 20
        SortListView LstFFSPmt, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView LstFFSPmt, ColumnHeader.Index, ldtDateTime, m_blnDirection(21)
    Case 22
        SortListView LstFFSPmt, ColumnHeader.Index, ldtNumber, m_blnDirection(22)
    End Select

End Sub

Private Sub LstFFSPmt_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    EditFlag = False
    Call ProcessCheckedItem(Item)
End Sub

Private Sub LstFFSPmt_KeyDown(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "LstFFSPmt_KeyDown"
On Error GoTo debugs

    If LstFFSPmt.ListItems.Count Then
        If Trim(LstFFSPmt.SelectedItem.SubItems(12)) <> "" Then
            CmdSave.Enabled = False
            PVframe.Enabled = False
            CboBank.Enabled = False
            TxtPVNo3 = LstFFSPmt.SelectedItem.SubItems(12)
            If Len(LstFFSPmt.SelectedItem.SubItems(2)) Then
                txtBordxNo3 = LstFFSPmt.SelectedItem.SubItems(2)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(13)) Then
                TxtPVDate3 = LstFFSPmt.SelectedItem.SubItems(13)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(14)) Then
                TxtChqNo = LstFFSPmt.SelectedItem.SubItems(14)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(15)) Then
                TxtChqDate = LstFFSPmt.SelectedItem.SubItems(15)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(16)) Then
                TxtChqAmt = LstFFSPmt.SelectedItem.SubItems(16)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(17)) Then
                CboACMth3 = Mid(LstFFSPmt.SelectedItem.SubItems(17), 1, 2)
                CboACYr3 = Mid(LstFFSPmt.SelectedItem.SubItems(17), 4, 4)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(18)) Then
                CboBank.Text = LstFFSPmt.SelectedItem.SubItems(18)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(19)) Then
                TxtACNo = LstFFSPmt.SelectedItem.SubItems(19)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(20)) Then
                TxtRemark = LstFFSPmt.SelectedItem.SubItems(20)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(21)) Then
                txtPayIndex = LstFFSPmt.SelectedItem.SubItems(21)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(22)) Then
                CboPayee = LstFFSPmt.SelectedItem.SubItems(22)
            End If
            
            If Len(LstFFSPmt.SelectedItem.SubItems(23)) Then
                txtPayIndex = LstFFSPmt.SelectedItem.SubItems(23)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(24)) Then
                CboPayee = LstFFSPmt.SelectedItem.SubItems(24)
            End If
            
        Else
            Call ClearFramePV
        End If
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub LstFFSPmt_KeyUp(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "LstFFSPmt_KeyUp"
On Error GoTo debugs

    If LstFFSPmt.ListItems.Count Then
        If Trim(LstFFSPmt.SelectedItem.SubItems(12)) <> "" Then
            CmdSave.Enabled = False
            PVframe.Enabled = False
            CboBank.Enabled = False
            TxtPVNo3 = LstFFSPmt.SelectedItem.SubItems(12)
            If Len(LstFFSPmt.SelectedItem.SubItems(2)) Then
                txtBordxNo3 = LstFFSPmt.SelectedItem.SubItems(2)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(13)) Then
                TxtPVDate3 = LstFFSPmt.SelectedItem.SubItems(13)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(14)) Then
                TxtChqNo = LstFFSPmt.SelectedItem.SubItems(14)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(15)) Then
                TxtChqDate = LstFFSPmt.SelectedItem.SubItems(15)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(16)) Then
                TxtChqAmt = LstFFSPmt.SelectedItem.SubItems(16)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(17)) Then
                CboACMth3 = Mid(LstFFSPmt.SelectedItem.SubItems(17), 1, 2)
                CboACYr3 = Mid(LstFFSPmt.SelectedItem.SubItems(17), 4, 4)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(18)) Then
                CboBank.Text = LstFFSPmt.SelectedItem.SubItems(18)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(19)) Then
                TxtACNo = LstFFSPmt.SelectedItem.SubItems(19)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(20)) Then
                TxtRemark = LstFFSPmt.SelectedItem.SubItems(20)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(21)) Then
                txtPayIndex = LstFFSPmt.SelectedItem.SubItems(21)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(22)) Then
                CboPayee = LstFFSPmt.SelectedItem.SubItems(22)
            End If
            
            If Len(LstFFSPmt.SelectedItem.SubItems(23)) Then
                txtPayIndex = LstFFSPmt.SelectedItem.SubItems(23)
            End If
            If Len(LstFFSPmt.SelectedItem.SubItems(24)) Then
                CboPayee = LstFFSPmt.SelectedItem.SubItems(24)
            End If
        Else
            Call ClearFramePV
        End If
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub TxtACNo_lostfocus()
    TxtACNo = ChkStr(TxtACNo)
End Sub

Private Sub TxtBodxDate2_LostFocus()
    If IsDate(TxtBodxDate2) Then
    Else
        If TxtBodxDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtBodxDate2.SetFocus
        End If
    End If
End Sub

Private Sub TxtBordxDate1_LostFocus()
    If IsDate(TxtBordxDate1) Then
    Else
        If TxtBordxDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtBordxDate1.SetFocus
        End If
    End If

End Sub

Private Sub TxtBordxNo1_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtBordxNo1_LostFocus()
    TxtBordxNo1 = ChkStr(TxtBordxNo1)
    CmdSearch.Default = True
End Sub

Private Sub TxtBordxNo2_LostFocus()
    TxtBordxNo2 = ChkStr(TxtBordxNo2)
End Sub

Private Sub TxtChqAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtChqDate_LostFocus()
    If IsDate(TxtChqDate) Then
    Else
        If TxtChqDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtChqDate.SetFocus
        End If
    End If
End Sub

Private Sub TxtChqNo_lostfocus()
    TxtChqNo = ChkStr(TxtChqNo)
End Sub


Private Sub TxtPVDate1_LostFocus()
    If IsDate(TxtPVDate1) Then
    Else
        If TxtPVDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPVDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtPVDate2_LostFocus()
    If IsDate(TxtPVDate2) Then
    Else
        If TxtPVDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPVDate2.SetFocus
        End If
    End If
End Sub

Private Sub TxtPVDate3_LostFocus()
    If IsDate(TxtPVDate3) Then
    Else
        If TxtPVDate3 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPVDate3.SetFocus
        End If
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

Private Sub TxtPVNo1_lostfocus()
    TxtPVNo1 = ChkStr(TxtPVNo1)
End Sub

Private Sub TxtPVNo2_lostfocus()
    TxtPVNo2 = ChkStr(TxtPVNo2)
End Sub

Private Sub TxtPVNo3_lostfocus()
    TxtPVNo3 = ChkStr(TxtPVNo3)
End Sub

Private Sub TxtRemark_lostfocus()
    TxtRemark = ChkStr(TxtRemark)
End Sub

Private Sub ComboKeyPressNum(tmpComboBox As ComboBox, KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    'NewText = LCase(Left(tmpComboBox.Text, tmpComboBox.SelStart) + Chr(KeyAscii) + Mid(tmpComboBox.Text, tmpComboBox.SelStart + tmpComboBox.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To tmpComboBox.ListCount - 1
        If NewText = LCase(Left(tmpComboBox.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = tmpComboBox.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
         If ValidCount = 1 Then
           tmpComboBox.Text = ValidValue
           tmpComboBox.SelStart = 0
           tmpComboBox.SelLength = Len(ValidValue)
         End If
    End If
End Sub
Private Sub ComboKeyPress(tmpCbo As ComboBox, KeyAscii As Integer)
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
Private Sub ComboListing()
Const METHOD_NAME = "ComboListing"
On Error GoTo debugs

Dim IntMth As Integer
Dim BankRec As New ADODB.Recordset
Dim TempRec As New ADODB.Recordset
Dim TempStr As String
    
    CboBank.Clear
   
        BankRec.Open "Banks", medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
        Do While Not BankRec.EOF
            CboBank.AddItem ChkStr(BankRec!BankCode) & " - " & ChkStr(BankRec!BankDesc)
            BankRec.MoveNext
        Loop
        BankRec.Close
        Set BankRec = Nothing
    
    TempStr = "select GRPCoName from GRPCompany"
    TempRec.Open TempStr, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not TempRec.EOF
        CboPayee.AddItem IIf(Len(TempRec!GrpCoName), TempRec!GrpCoName, "")
        TempRec.MoveNext
    Loop
    TempRec.Close
    Set TempRec = Nothing
    CboPayee.Text = "HEALTH CONNECT SDN BHD"
    
    
    CboPmtStat.Clear
    With CboPmtStat
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
    End With
    CboPmtStat.ListIndex = 0
    
    CboACMth1.Clear
    CboACMth2.Clear
    CboACMth3.Clear
    For IntMth = 1 To 12
        CboACMth1.AddItem Format(IntMth, "00")
        CboACMth2.AddItem Format(IntMth, "00")
        CboACMth3.AddItem Format(IntMth, "00")
    Next IntMth
    
    CboACYr1.Clear
    CboACYr2.Clear
    CboACYr3.Clear
    For IntMth = 1999 To Year(Date) + 2
        CboACYr1.AddItem (IntMth)
        CboACYr2.AddItem (IntMth)
        CboACYr3.AddItem (IntMth)
    Next IntMth
    CboBank.Text = "MB - MAYBANK"
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub PaymentFFSListing(Optional strAppend As String)
Const METHOD_NAME = "PaymentFFSListing"
On Error GoTo debugs

Dim PVRec As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim strsqlPay As String
Dim PAY As New ADODB.Recordset
Dim INVAmt As Variant
Dim Current As String
Dim TotalAmt As Variant
Dim Check As Boolean
Dim tmpPayFFS As Double
Dim REcFound As Boolean

    Current = 0
    If ChkBordxNo.Value = 0 Then
        LstFFSPmt.ListItems.Clear
    End If
    Call ClearFramePV
    Screen.MousePointer = vbHourglass
    REcFound = False
    Call ReadRecord(strAppend, PVRec)
    Do
        Do While Not PVRec.EOF
        REcFound = True
            With PVRec
                Set Record = LstFFSPmt.ListItems.Add(, , "")
                If Trim(!PVPayStatus) = "F" Then
                    Record.SubItems(1) = "Full"
                End If
                Record.SubItems(2) = IIf(Len(!Bordxrefno), !Bordxrefno, "")
                If Len(!BordxDate) Then
                    Record.SubItems(3) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                If IsNumeric((!CLMBordxAmt)) Then
                    Record.SubItems(4) = Format(!CLMBordxAmt, "#,##0.00")
                Else
                    Record.SubItems(4) = Format("0", "#,##0.00")
                End If
                tmpPayFFS = IIf(IsNumeric(!CLMMCOProcFee), !CLMMCOProcFee, 0)
                Record.SubItems(6) = Format(tmpPayFFS, "#,##0.00")
                Record.SubItems(7) = Format(IIf(IsNumeric(!FFSPaidAmt), !FFSPaidAmt, 0), "#,##0.00")
                Record.SubItems(8) = Format(CDbl(Record.SubItems(6) - Record.SubItems(7)), "#,##0.00")
                Record.SubItems(9) = Format(Record.SubItems(9), "#,##0.00")
                If Len(!GroupCompany) Then
                    Record.SubItems(11) = Trim(!GroupCompany)
                End If
                If Len(!PVNo) Then
                    Record.SubItems(12) = Trim(!PVNo)
                End If
                If Len(!PVDate) Then
                    Record.SubItems(13) = Format(!PVDate, "dd/mm/yyyy")
                End If
                If Len(!PVChqNo) Then
                    Record.SubItems(14) = Trim(!PVChqNo)
                End If
                If Len(!PVChqDate) Then
                    Record.SubItems(15) = Trim(!PVChqDate)
                End If
                If IsNumeric(!PVChqAmt) Then
                    Record.SubItems(16) = Trim(!PVChqAmt)
                End If
                If Len(Trim(!FFSAccMth) & Trim(!FFSAccYr)) Then
                    Record.SubItems(17) = Format(!FFSAccMth & "/" & !FFSAccYr, "mm/yyyy")
                End If
                If Len(Trim(!PVBank)) Then
                    Record.SubItems(18) = Trim(!PVBank)
                End If
                If Len(Trim(!PVAcc)) Then
                    Record.SubItems(19) = Trim(!PVAcc)
                End If
                If Len(Trim(!Remark)) Then
                    Record.SubItems(20) = Trim(!Remark)
                End If
                If Len(Trim(!PaymentIndex)) Then
                    Record.SubItems(21) = Trim(!PaymentIndex)
                End If
                If Len(Trim(!GrpCoName)) Then
                    Record.SubItems(22) = Trim(!GrpCoName)
                End If
                If Len(!PVRecNo) Then
                    Record.SubItems(23) = !PVRecNo
                End If
                If Len(!PVRecDate) Then
                    Record.SubItems(24) = Format(!PVRecDate, "dd/mm/yyyy")
                End If
                lblTotalAmt = lblTotalAmt + tmpPayFFS
                Current = Current + 1
                LblRecordR = Current
            End With

            PVRec.MoveNext
            DoEvents
                If intExit = True Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    intExit = False
    PVRec.Close
    Set PVRec = Nothing
    If REcFound = False Then
        MsgBox "No Record Found ! ", vbOKOnly + vbCritical
    Else
    End If
    Screen.MousePointer = Default
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ClearFramePV()
    TxtPVNo3 = ""
    TxtPVDate3.Text = "__/__/____"
    TxtChqNo = ""
    TxtChqDate.Text = "__/__/____"
    TxtChqAmt = ""
    CboBank.Text = "MB - MAYBANK"
    TxtACNo = ""
    CboACMth3.Text = ""
    CboACYr3.Text = ""
    TxtRemark = ""
    'txtRecNo = ""
    'txtRecDate = "__/__/____"
End Sub

Private Sub ReadRecord(strAppend As String, PVRec As ADODB.Recordset)
Const METHOD_NAME = "ReadRecord"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter

    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "SP_CISPayment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("PaymentType", adVarChar, adParamInput, 10, "PayFFS")
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm2
    PVRec.CursorLocation = adUseClient
    PVRec.CursorType = adOpenForwardOnly
    PVRec.CacheSize = 100
    Set PVRec = cmd.Execute
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
    
End Sub

Private Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
    If Item.Checked = True Then
        If Mid(Item.SubItems(1), 1, 1) <> "" Then
            MsgBox "Payment has been made!", vbCritical
            Item.Checked = False
        Else
            Item.SubItems(9) = CDbl(Item.SubItems(8))
            Item.SubItems(1) = "Full"
            Item.SubItems(8) = 0
            LblAmtR = Format(LblAmtR + CDbl(Item.SubItems(9)), "#,##0.00")
        End If
    Else
        LblAmtR = Format(LblAmtR - CDbl(Item.SubItems(9)), "#,##0.00")
        Item.SubItems(8) = CDbl(Item.SubItems(9))
        Item.SubItems(9) = 0
        Item.SubItems(1) = ""
    End If
End Function

Private Sub UpdateRecord(tmpBordxNo As String, PaymentAmt As Double, PymtIndex As String)
Const METHOD_NAME = "UpdateRecord"
On Error GoTo debugs


Dim Update
Dim StrSql As String
Dim PAY As New ADODB.Recordset
Dim strsqlPay As String
Dim GRPC As New ADODB.Recordset
Dim strsqlGRPC As String
Dim ZZ As String
Dim TempBank As String
Dim TempCode As String

    ZZ = InStr(1, CboBank, "-")
    TempBank = Trim(Left(CboBank, IIf(ZZ - 1 <= 0, Len(Trim(CboBank)), ZZ - 1)))
    strsqlPay = " Select PaymentIndex, * from PaymentFFS " & _
                " Where PaymentIndex = '" & Trim(PymtIndex) & "'"
    PAY.Open strsqlPay, MediConnCISPayment, adOpenKeyset, adLockOptimistic
    
    strsqlGRPC = " Select GRPCoCode from GRPCompany " & _
                " Where GRPCoName = '" & Trim(CboPayee) & "'"
    GRPC.Open strsqlGRPC, MediConnCISDB, adOpenKeyset, adLockOptimistic
    TempCode = IIf(Len(Trim(GRPC!GRPCoCode)), Trim(GRPC!GRPCoCode), "")
    GRPC.Close
    
    If PAY.RecordCount > 0 Then
    Else
        PAY.AddNew
    End If
    With PAY
        !Bordxrefno = Trim(tmpBordxNo)
        !PVNo = TxtPVNo3
        !PVDate = Format(TxtPVDate3, "dd/mm/yyyy")
        !PVChqDate = IIf(TxtChqDate <> "__/__/____", Format(TxtChqDate, "dd/mm/yyyy"), "")
        !PVChqNo = IIf(Len(Trim(TxtChqNo)), TxtChqNo, "")
        !PVChqAmt = IIf(IsNumeric(TxtChqAmt), TxtChqAmt, 0)
        !PVBank = IIf(Len(TempBank), Trim(TempBank), "")
        !PVAcc = IIf(Len(Trim(TxtACNo)), Trim(TxtACNo), "")
        !Remark = IIf(Len(Trim(TxtRemark)), TxtRemark, "")
        !FFSPaidAmt = IIf(IsNumeric(PaymentAmt), PaymentAmt, 0)
        !FFSAccMth = IIf(CboACMth3 <> "", Trim(CboACMth3), "")
        !FFSAccYr = IIf(CboACYr3 <> "", Trim(CboACYr3), "")
        !PVLastUpdateDate = Date
        !PVLastUpdateUser = Logon
        !PVPayStatus = "F"
        !GRPCoCode = Trim(TempCode)
    End With
    PAY.Update
    
    PAY.Close
    Set PAY = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub PrintPV()
Const METHOD_NAME = "PrintPV"
On Error GoTo debugs

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim StrSql As String
Dim strAppend As String
Dim PVRec As New ADODB.Recordset
Dim Cnt As Integer
Dim tmpGRPCode As String
Dim tmpGRPName As String
Dim tmpGRPAdd1 As String
Dim tmpGRPAdd2 As String
Dim tmpGRPAdd3 As String
Dim TempBordxNo As String
Dim TempBordxDate As String
    Screen.MousePointer = vbHourglass
    If CmdPrintPV Then
        strAppend = " And PVNo = '" & TxtPVNo3 & "'"
    Else
        If LstFFSPmt.SelectedItem.Checked = True Then
            TempBordxNo = Trim(LstFFSPmt.SelectedItem.SubItems(2))
            TempBordxDate = Trim(LstFFSPmt.SelectedItem.SubItems(3))
           'strAppend = " AND GRPCoCode='" & Trim(tmpGRPCode) & "' AND BordxRefNo = '" & Trim(TempBordxNo) & _ del
            strAppend = " AND BordxRefNo = '" & Trim(TempBordxNo) & _
                        "' AND BordxDate='" & Format(TempBordxDate, "YYYY/MM/DD") & "'"
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
  
    Call ReadRecord(strAppend, PVRec)
    objExcel.Interactive = False

   
    Do While Not PVRec.EOF
        tmpGRPCode = IIf(Len(Trim(PVRec!GRPCoCode)), Trim(PVRec!GRPCoCode), "")
        tmpGRPName = IIf(Len(Trim(PVRec!GrpCoName)), Trim(PVRec!GrpCoName), "")
        tmpGRPAdd1 = IIf(Len(Trim(PVRec!GRPCoAdd1)), Trim(PVRec!GRPCoAdd1), "")
        tmpGRPAdd2 = IIf(Len(Trim(PVRec!GRPCoAdd2)), Trim(PVRec!GRPCoAdd2), "")
        tmpGRPAdd3 = IIf(Len(Trim(PVRec!GRPCoAdd3)), Trim(PVRec!GRPCoAdd3), "")
        'If Trim(PVRec!GRPCoCode) <> "" Then
        '    tmpGRPCode = Trim(PVRec!GRPCoCode)
        '    Call ClinicHQInfo(tmpGRPCode, tmpGRPName, tmpGRPAdd1, tmpGRPAdd2, tmpGRPAdd3)
        'End If
        If objExcel.Visible = False Then
            objExcel.Visible = True
        End If
    
        objExcel.WindowState = xlMaximized
        Cnt = Cnt + 1
        If Cnt = 1 Then
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\PV-FFS.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        objWorksheet.Cells(8, "C") = tmpGRPName
        objWorksheet.Cells(9, "C") = tmpGRPAdd1
        objWorksheet.Cells(10, "C") = tmpGRPAdd2
        objWorksheet.Cells(11, "C") = tmpGRPAdd3
        objWorksheet.Cells(8, "H") = TxtPVNo3 '& "/" & Format(PVRec!PVDate, "MM/YY")
        objWorksheet.Cells(9, "H") = PVRec!PVChqNo
        objWorksheet.Cells(10, "H") = PVRec!PVChqDate
        objWorksheet.Cells(37, "C") = Logon
         
        objWorksheet.Cells(15 + CDbl(Cnt), "A") = Cnt
        objWorksheet.Cells(15 + CDbl(Cnt), "B") = IIf(Len(PVRec!Bordxrefno), Trim(PVRec!Bordxrefno), "")
        objWorksheet.Cells(15 + CDbl(Cnt), "C") = PVRec!BordxDate
        If CmdPrintPV Then
            objWorksheet.Cells(15 + CDbl(Cnt), "H") = PVRec!FFSPaidAmt
        Else
            'objWorksheet.Cells(15 + CDbl(Cnt), "H") = PVRec!CLMPayableByMedix2C
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

            MediConnCISDB.Close
            Set MediConnCISDB = Nothing
            medixmasterconn.Close
            Set medixmasterconn = Nothing
            MediConnCISWS.Close
            Set MediConnCISWS = Nothing
            MediConnCISPayment.Close
            Set MediConnCISPayment = Nothing

            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            objExcel.Quit
            Screen.MousePointer = vbDefault
            Exit Sub
HANDLE_ERROR:
            MediConnCISDB.Close
            Set MediConnCISDB = Nothing
            medixmasterconn.Close
            Set medixmasterconn = Nothing
            MediConnCISWS.Close
            Set MediConnCISWS = Nothing
            MediConnCISPayment.Close
            Set MediConnCISPayment = Nothing
            
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            Screen.MousePointer = vbDefault
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            GoTo ExitFunction
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub
