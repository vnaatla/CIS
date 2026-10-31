VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReceiptFFS 
   Caption         =   "Receipt for FFS"
   ClientHeight    =   8160
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11832
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   11832
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
      Left            =   75
      TabIndex        =   61
      Top             =   1720
      Width           =   855
   End
   Begin MSComctlLib.ListView LstFFSRcpt 
      Height          =   4350
      Left            =   75
      TabIndex        =   34
      Top             =   1995
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   7684
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
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Receipt Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "BordxDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Bordx Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "FFS Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "FFS Rcpt Amt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "OS Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Receipt Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Receipt No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Receipt Date"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Remarks"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Acc Yr"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Acc Mth"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Chq No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "TotalReceipt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Group Company"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame RecpFrame 
      Height          =   975
      Left            =   80
      TabIndex        =   32
      Top             =   6300
      Width           =   11700
      Begin VB.ComboBox CboACMth3 
         Height          =   315
         Left            =   6600
         TabIndex        =   17
         Top             =   240
         Width           =   1250
      End
      Begin VB.TextBox TxtRemark 
         Height          =   615
         Left            =   8880
         TabIndex        =   19
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox TxtChqNo 
         Height          =   285
         Left            =   3450
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtRecpNo3 
         Height          =   285
         Left            =   900
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtRecpAmt 
         Height          =   285
         Left            =   3450
         TabIndex        =   16
         Top             =   500
         Width           =   1455
      End
      Begin MSMask.MaskEdBox TxtRecpDate3 
         Height          =   285
         Left            =   900
         TabIndex        =   14
         Top             =   495
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboACYr3 
         Height          =   315
         Left            =   6600
         TabIndex        =   18
         Top             =   520
         Width           =   1250
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
         Height          =   255
         Left            =   8250
         TabIndex        =   54
         Top             =   255
         Width           =   645
      End
      Begin VB.Label LblACYr2 
         Caption         =   "A/C Period Yr"
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
         Left            =   5350
         TabIndex        =   53
         Top             =   540
         Width           =   1250
      End
      Begin VB.Label LblACMth2 
         Caption         =   "A/C Period Mth"
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
         Left            =   5350
         TabIndex        =   52
         Top             =   260
         Width           =   1250
      End
      Begin VB.Label LblRecpAmt 
         Caption         =   "Recp Amt"
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
         Left            =   2500
         TabIndex        =   51
         Top             =   540
         Width           =   900
      End
      Begin VB.Label LblChqNo 
         Caption         =   "Cheque No"
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
         Left            =   2500
         TabIndex        =   50
         Top             =   260
         Width           =   900
      End
      Begin VB.Label LblRecpDate2 
         Caption         =   "Date"
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
         Left            =   180
         TabIndex        =   49
         Top             =   540
         Width           =   900
      End
      Begin VB.Label LblRecpNo2 
         Caption         =   "Recp No"
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
         Left            =   180
         TabIndex        =   48
         Top             =   260
         Width           =   900
      End
   End
   Begin VB.Frame BtnFrame 
      Height          =   600
      Left            =   80
      TabIndex        =   33
      Top             =   7210
      Width           =   11700
      Begin VB.TextBox LblChkedAmtR 
         Height          =   285
         Left            =   5400
         TabIndex        =   65
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox lblTotalAmt 
         Height          =   285
         Left            =   3000
         TabIndex        =   64
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox LblRecordR 
         Height          =   285
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   6660
         TabIndex        =   21
         Top             =   165
         Width           =   705
      End
      Begin VB.TextBox TxtPVStatus 
         Height          =   225
         Left            =   6720
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
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
         Height          =   355
         Left            =   11000
         TabIndex        =   30
         Top             =   165
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
         Height          =   355
         Left            =   10450
         TabIndex        =   29
         Top             =   165
         Width           =   560
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   9460
         TabIndex        =   28
         Top             =   165
         Width           =   1000
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
         Height          =   355
         Left            =   8880
         TabIndex        =   27
         Top             =   165
         Width           =   600
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "S&earch"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   8140
         TabIndex        =   26
         Top             =   165
         Width           =   750
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   7360
         TabIndex        =   22
         Top             =   165
         Width           =   800
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   6660
         TabIndex        =   24
         Top             =   165
         Width           =   705
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
         Height          =   355
         Left            =   7360
         TabIndex        =   25
         Top             =   165
         Width           =   800
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
         Left            =   2160
         TabIndex        =   62
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblChkedAmt 
         Caption         =   "Checked Amt"
         Height          =   255
         Left            =   4275
         TabIndex        =   56
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LblRecord 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1200
         TabIndex        =   55
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame SchFrame 
      Height          =   1455
      Left            =   80
      TabIndex        =   31
      Top             =   240
      Width           =   11700
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   2900
      End
      Begin VB.CheckBox ChkBordx 
         Caption         =   "Check1"
         Height          =   255
         Left            =   4815
         TabIndex        =   46
         Top             =   240
         Value           =   1  'Checked
         Width           =   255
      End
      Begin MSMask.MaskEdBox TxtRecpDate2 
         Height          =   285
         Left            =   9840
         TabIndex        =   6
         Top             =   255
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtRecpDate1 
         Height          =   285
         Left            =   8040
         TabIndex        =   5
         Top             =   255
         Width           =   1333
         _ExtentX        =   2350
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   3480
         TabIndex        =   23
         Top             =   240
         Width           =   1222
      End
      Begin VB.ComboBox CboPayStat 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   495
         Width           =   2895
      End
      Begin VB.TextBox TxtRecpNo2 
         Height          =   285
         Left            =   3480
         TabIndex        =   3
         Top             =   780
         Width           =   1215
      End
      Begin VB.TextBox TxtRecpNo1 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   780
         Width           =   1215
      End
      Begin MSMask.MaskEdBox TxtBordxDate1 
         Height          =   315
         Left            =   8040
         TabIndex        =   7
         Top             =   495
         Width           =   1333
         _ExtentX        =   2350
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
      Begin MSMask.MaskEdBox TxtBodxDate2 
         Height          =   315
         Left            =   9840
         TabIndex        =   8
         Top             =   480
         Width           =   1335
         _ExtentX        =   2350
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
         ItemData        =   "FrmReceiptFFS.frx":0000
         Left            =   1800
         List            =   "FrmReceiptFFS.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1035
         Width           =   2940
      End
      Begin VB.ComboBox CboACMth1 
         Height          =   315
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   750
         Width           =   1333
      End
      Begin VB.ComboBox CboACYr1 
         Height          =   315
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1035
         Width           =   1333
      End
      Begin VB.ComboBox CboACMth2 
         Height          =   315
         Left            =   9840
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   750
         Width           =   1333
      End
      Begin VB.ComboBox CboACYr2 
         Height          =   315
         Left            =   9840
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1035
         Width           =   1333
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
         Left            =   180
         TabIndex        =   60
         Top             =   1100
         Width           =   1455
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
         Left            =   6480
         TabIndex        =   59
         Top             =   520
         Width           =   1500
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
         Left            =   9480
         TabIndex        =   58
         Top             =   525
         Width           =   270
      End
      Begin VB.Label LblTo5 
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
         Left            =   9495
         TabIndex        =   45
         Top             =   1065
         Width           =   255
      End
      Begin VB.Label LblTo4 
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
         Left            =   9495
         TabIndex        =   44
         Top             =   795
         Width           =   255
      End
      Begin VB.Label LblTo3 
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
         Left            =   9495
         TabIndex        =   43
         Top             =   285
         Width           =   255
      End
      Begin VB.Label LblTo2 
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
         Left            =   3120
         TabIndex        =   42
         Top             =   825
         Width           =   255
      End
      Begin VB.Label LblACYr1 
         Caption         =   "A/C Period Yr"
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
         Left            =   6510
         TabIndex        =   41
         Top             =   1065
         Width           =   1455
      End
      Begin VB.Label LblACMth1 
         Caption         =   "A/C Period Mth"
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
         Left            =   6510
         TabIndex        =   40
         Top             =   795
         Width           =   1455
      End
      Begin VB.Label LblRecpDate1 
         Caption         =   "Recp Date"
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
         Left            =   6510
         TabIndex        =   39
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label LblRecpNo1 
         Caption         =   "Recp No"
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
         Left            =   180
         TabIndex        =   38
         Top             =   825
         Width           =   1455
      End
      Begin VB.Label LblTo1 
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
         Left            =   3150
         TabIndex        =   37
         Top             =   255
         Width           =   255
      End
      Begin VB.Label LblPayStat 
         Caption         =   "Receipt Status"
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
         Left            =   180
         TabIndex        =   36
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label LblBordx 
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
         Height          =   255
         Left            =   180
         TabIndex        =   35
         Top             =   260
         Width           =   1450
      End
   End
   Begin VB.Label LblBold 
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
      TabIndex        =   47
      Top             =   7830
      Width           =   1575
   End
   Begin VB.Label LblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "FFS Receipt"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "FrmReceiptFFS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpBordxNo As String
Dim intExit As Boolean
Dim EditFlag As Boolean
Dim RecIndex As String
Dim tmpRecIndex As String
Dim listloop1 As Integer
Dim Check1 As Boolean

Private Sub CboACMth1_KeyPress(KeyAscii As Integer)
    Call ComboKeyPressNum(CboACMth1, KeyAscii)
End Sub

Private Sub CboACMth2_KeyPress(KeyAscii As Integer)
    Call ComboKeyPressNum(CboACMth3, KeyAscii)
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



Private Sub CboPayStat_Keypress(KeyAscii As Integer)
    Call ComboKeyPress(CboPayStat, KeyAscii)
End Sub

Private Sub ComboKeyPressNum(tmpComboBox As ComboBox, KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
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

Private Sub ClearRecpFrame()
    TxtRecpNo3 = ""
    TxtRecpDate3.Text = "__/__/____"
    TxtChqNo = ""
'   CboACMth3.Text = ""
'   CboACYr3.Text = ""
    TxtRemark = ""
End Sub

Private Sub ChkBordx_Click()
    LstFFSRcpt.ListItems.Clear
    If ChkBordx.Value = 1 Then
        LblTo1.Visible = False
        TxtBordx1.Width = 2900
        TxtBordx2.Visible = False
        TxtBordx1 = ""
    Else
        LblTo1.Visible = True
        TxtBordx1.Width = 1222
        TxtBordx2.Visible = True
    End If
    TxtBordx1.SetFocus
End Sub
Private Sub CmdAdd_Click()
Dim ZZ As Integer
Dim ItemChk As Boolean

    EditFlag = False
    ItemChk = False
    For ZZ = 1 To LstFFSRcpt.ListItems.Count
        If LstFFSRcpt.ListItems(ZZ).Checked = True Then
            ItemChk = True
            Exit For
        End If
    Next ZZ

    CmdUpdate.Enabled = False
    CmdEdit.Visible = True
    If ItemChk Then
        Call ClearFramePV
        cmdAdd.Visible = False
        CmdSave.Visible = True
        CmdSave.Enabled = True
        RecpFrame.Enabled = True
        'CboBank.Enabled = True
    Else
        MsgBox "Please checked at least 1 Claim  ! ", vbCritical + vbOKOnly, DialogBoxTitle
    End If
    
End Sub
Private Sub ClearFramePV()
    TxtRecpNo3 = ""
    TxtRecpDate3.Text = "__/__/____"
    TxtChqNo = ""
    TxtRecpAmt = 0
    CboACMth3 = ""
    CboACYr3 = ""
    TxtRemark = ""
End Sub

Private Sub cmdCheckAll_Click()
Dim LstCnt As Integer
    If LstFFSRcpt.ListItems.Count <> 0 Then
        LstCnt = LstFFSRcpt.ListItems.Count
    
        For LstCnt = 1 To LstCnt
            If LstFFSRcpt.ListItems(LstCnt).Checked = False Then
                LstFFSRcpt.ListItems(LstCnt).Checked = True
                LstFFSRcpt.ListItems(LstCnt).SubItems(1) = "Full"
                LstFFSRcpt.ListItems(LstCnt).SubItems(8) = LstFFSRcpt.ListItems(LstCnt).SubItems(7)
                LstFFSRcpt.ListItems(LstCnt).SubItems(7) = 0
                LblChkedAmtR = Format(LblChkedAmtR + CDbl(LstFFSRcpt.ListItems(LstCnt).SubItems(8)), "#,##0.00")
            ElseIf LstFFSRcpt.ListItems(LstCnt).Checked = True Then
                LblChkedAmtR = Format(LblChkedAmtR - CDbl(LstFFSRcpt.ListItems(LstCnt).SubItems(8)), "#,##0.00")
                LstFFSRcpt.ListItems(LstCnt).Checked = False
                LstFFSRcpt.ListItems(LstCnt).SubItems(1) = ""
                LstFFSRcpt.ListItems(LstCnt).SubItems(7) = LstFFSRcpt.ListItems(LstCnt).SubItems(8)
                LstFFSRcpt.ListItems(LstCnt).SubItems(8) = 0
            End If
        Next LstCnt
    Else
        MsgBox "Please select a record.", vbInformation
    End If

End Sub

Private Sub cmdClear_Click()
Dim IntCount As Integer

    Call ClearForm(Me)
    'CmdSave.Visible = False
    'CmdAdd.Visible = True
    'CmdEdit.Visible = True
    'CmdUpdate.Visible = False
    'EditFlag = False
    RecpFrame.Enabled = False
    LstFFSRcpt.ListItems.Clear
    CboPayStat.ListIndex = 0
    LblRecordR = ""
    LblChkedAmtR = ""
    lblTotalAmt = ""
    TxtBordx1.SetFocus
    ChkBordx.Value = 1
    LblTo1.Visible = False
    TxtBordx1.Width = 2900
    TxtBordx2.Visible = False
    LstFFSRcpt.Enabled = True
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
Dim ZZ As Integer
Dim ItemChk As Boolean

    EditFlag = True
    CmdSave.Enabled = False
    RecpFrame.Enabled = False
        CmdEdit.Visible = False
        CmdUpdate.Visible = True
        CmdUpdate.Enabled = True
        RecpFrame.Enabled = True
        EditFlag = True
        If Len(Trim(LstFFSRcpt.SelectedItem.SubItems(1))) Then
            RecpFrame.Enabled = True
        End If
End Sub

Private Sub cmdExit_Click()
    intExit = True
    Unload Me
End Sub

Private Sub CmdSave_Click()

Const METHOD_NAME = "CmdSave_Click"
On Error GoTo debugs

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
Dim RecNumber As String
Dim PrevGRP, CurrGRP As String
Dim DiffGRP As Boolean
Dim CLNCnt As Integer
Dim CompTrans As Boolean
    
    ItemChecked = False
    TmpPVCode = ""
    
    For i = 1 To LstFFSRcpt.ListItems.Count
        If LstFFSRcpt.ListItems(i).Checked = True Then
            ItemChecked = True
            Exit For
        End If
    Next i
    
    If ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    ElseIf TxtRecpNo3 = "" Then
        MsgBox "Please Enter Receipt No", vbOKOnly + vbCritical
    ElseIf IsDate(TxtRecpDate3) = False Then
        MsgBox "Invalid Receipt Date!", vbCritical
    ElseIf TxtRecpDate3 = "__/__/____" Then
        MsgBox "Please Enter Receipt Date! ", vbCritical
        TxtRecpDate3.SetFocus
    ElseIf TxtRecpAmt = "" Then
        MsgBox "Please Enter Receipt Amt!", vbCritical
        TxtRecpAmt.SetFocus
    ElseIf CDbl(TxtRecpAmt) <> CDbl(LblChkedAmtR) And TxtRemark = "" Then
        MsgBox "Amount to paid is incorrect" & vbNewLine & "Please key in remarks.", vbCritical + vbOKOnly
        TxtRemark.SetFocus
    ElseIf Trim(TxtChqNo) = "" Then
        MsgBox "Please Cheque Number!", vbCritical
        TxtChqNo.SetFocus
    ElseIf Trim(CboACMth3) = "" Then
        MsgBox "Please keyin Account Period Month ! ", vbCritical
        CboACMth3.SetFocus
    ElseIf Trim(CboACYr3) = "" Then
        MsgBox "Please keyin Account Period Year ! ", vbCritical
        CboACYr3.SetFocus
    Else
        Update = MsgBox("Do you want to update the record ? ", vbYesNo + vbExclamation)
        If Update = vbYes Then
            CompTrans = False
            MediConnCISPayment.beginTrans
            For i = 1 To LstFFSRcpt.ListItems.Count
                If LstFFSRcpt.ListItems(i).Checked = True Then
                    RecIndex = ""
                    tmpBordxNo = LstFFSRcpt.ListItems(i).SubItems(2)
                    tmpFFSAmt = LstFFSRcpt.ListItems(i).SubItems(5)
                    Call UpdateRecord(RecIndex, tmpBordxNo, tmpFFSAmt)
                End If
            Next i
            
            MediConnCISPayment.commitTrans
            CompTrans = True
            
            CmdSave.Visible = False
            cmdAdd.Visible = True
            LstFFSRcpt.ListItems.Clear
            'LblAmtR = 0
            PVSql = " and ReceiptNo = '" & Trim(TxtRecpNo3) & "'"
            Call ClearFramePV
            RecpFrame.Enabled = False
            LstFFSRcpt.ListItems.Clear
            Call ReceiptBordxListing(PVSql)
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
        End If
ErrorSave:
        Screen.MousePointer = vbDefault
        If CompTrans = False Then
            MediConnCISPayment.RollbackTrans
            MsgBox Err.Description
        End If
        
    End If
    LstFFSRcpt.Enabled = True
    
debugs:
        If CompTrans = False Then
            MediConnCISPayment.RollbackTrans
            MsgBox Err.Description
        End If
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
    
End Sub

Private Sub CmdSearch_Click()
Const METHOD_NAME = "CmdSearch_Click"
On Error GoTo debugs

Dim BordxNoArray() As String
Dim listrecord As Integer
Dim one As Boolean
Dim two As Boolean
Dim ScanCode As String
    lblTotalAmt = 0
    LblChkedAmtR = 0
    TxtRecpAmt = 0
    If Mid(TxtBordx1, 1, 2) = "LO" Then
        ScanCode = TxtBordx1
        TxtBordx1 = ""
        Call SearchForm(ScanCode)
    Else
        one = False
        two = False
        CmdEdit.Visible = True
        CmdEdit.Enabled = False
        CmdUpdate.Enabled = False
        'cmdEXPtoExcel.Enabled = False
        CmdSave.Enabled = False
        CmdClear.Enabled = False
        CmdExit.Enabled = False
        ReDim BordxNoArray(0)
        
        Screen.MousePointer = vbHourglass

        If ChkBordx.Value = 1 Then
            If TxtBordx1 = "" Then
                MsgBox "Please Enter Bordx No", vbOKOnly + vbCritical, DialogBoxTitle
                TxtBordx1.SetFocus
                CmdExit.Enabled = True
                CmdEdit.Enabled = True
                CmdUpdate.Enabled = True
                'cmdEXPtoExcel.Enabled = True
                CmdSave.Enabled = True
                CmdClear.Enabled = True
            Else
            
                Screen.MousePointer = vbHourglass
                
                listrecord = LstFFSRcpt.ListItems.Count
                If listrecord <> 0 Then
                    
                    For listloop1 = 1 To listrecord
                        Call BordxNoCheck
                    Next listloop1
                End If
                If Check1 = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    CmdExit.Enabled = True
                    CmdEdit.Enabled = True
                    CmdUpdate.Enabled = True
                    CmdSave.Enabled = True
                    CmdClear.Enabled = True
                ElseIf Check1 = False Then
                    Call SearchRecord
                End If
                Check1 = False
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

Private Function BordxNoCheck()
    If Trim(ChkStr(TxtBordx1)) = Trim(LstFFSRcpt.ListItems(listloop1).SubItems(2)) Then
        Check1 = True
    End If
End Function

Private Sub CmdStop_Click()
    intExit = 1
    TxtBordx1.SetFocus
    CmdExit.Enabled = True
End Sub

Private Sub CmdUpdate_Click()
Const METHOD_NAME = "CmdUpdate_Click"
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
Dim RecNumber As String
Dim PrevGRP, CurrGRP As String
Dim DiffGRP As Boolean
Dim CLNCnt As Integer
Dim CompTrans As Boolean
    
    If TxtRecpNo3 = "" Then
        MsgBox "Please Enter Receipt No", vbOKOnly + vbCritical
    ElseIf IsDate(TxtRecpDate3) = False Then
        MsgBox "Invalid Receipt Date!", vbCritical
    ElseIf TxtRecpDate3 = "__/__/____" Then
        MsgBox "Please Enter Receipt Date! ", vbCritical
        TxtRecpDate3.SetFocus
    ElseIf TxtRecpAmt = "" Then
        MsgBox "Please Enter Receipt Amt!", vbCritical
        TxtRecpAmt.SetFocus
    ElseIf Trim(TxtChqNo) = "" Then
        MsgBox "Please Cheque Number!", vbCritical
        TxtChqNo.SetFocus
    ElseIf Trim(CboACMth3) = "" Then
        MsgBox "Please keyin Account Period Month ! ", vbCritical
        CboACMth3.SetFocus
    ElseIf Trim(CboACYr3) = "" Then
        MsgBox "Please keyin Account Period Year ! ", vbCritical
        CboACYr3.SetFocus
    Else
        Update = MsgBox("Do you want to update the record ? ", vbYesNo + vbExclamation)
        If Update = vbYes Then
            CompTrans = False
            MediConnCISPayment.beginTrans
            
                RecIndex = LstFFSRcpt.SelectedItem.SubItems(15)
                tmpBordxNo = LstFFSRcpt.SelectedItem.SubItems(2)
                Call UpdateRecord(RecIndex, tmpBordxNo, TxtRecpAmt)
            MediConnCISPayment.commitTrans
            CompTrans = True
            CmdSave.Visible = False
            cmdAdd.Visible = True
            LstFFSRcpt.ListItems.Clear
            PVSql = " and ReceiptNo = '" & Trim(TxtRecpNo3) & "'"
            Call ClearFramePV
            RecpFrame.Enabled = False
            LstFFSRcpt.ListItems.Clear
            Call ReceiptBordxListing(PVSql)
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
        End If
ErrorSave:
        Screen.MousePointer = vbDefault
        If CompTrans = False Then
            MediConnCISDB.Close
            Set MediConnCISDB = Nothing
            medixmasterconn.Close
            Set medixmasterconn = Nothing
            MediConnCISWS.Close
            Set MediConnCISWS = Nothing
            MediConnCISPayment.Close
            Set MediConnCISPayment = Nothing
        
            MediConnCISPayment.RollbackTrans
            MsgBox Err.Description
        End If
    End If
    LstFFSRcpt.Enabled = True
    
debugs:

        If CompTrans = False Then
            MediConnCISDB.Close
            Set MediConnCISDB = Nothing
            medixmasterconn.Close
            Set medixmasterconn = Nothing
            MediConnCISWS.Close
            Set MediConnCISWS = Nothing
            MediConnCISPayment.Close
            Set MediConnCISPayment = Nothing
            MediConnCISPayment.RollbackTrans
        End If
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
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
Call ActionLogUpdate(12, "FrmReceiptFFS")
    Call ComboListing
    RecpFrame.Enabled = False
End Sub

Private Sub ComboListing()
Dim IntCount As Integer
Dim TempRec As New ADODB.Recordset
Dim TempStr As String

    ChkBordx.Value = 1
    LblTo1.Visible = False
    TxtBordx1.Width = 2900
    TxtBordx2.Visible = False
    
    CboPayStat.Clear
    With CboPayStat
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
    End With
    CboPayStat.ListIndex = 0
    
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
        CboACYr1.AddItem (IntCount)
        CboACYr2.AddItem (IntCount)
        CboACYr3.AddItem (IntCount)
    Next IntCount
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        TempStr = "select GRPCoName from GRPCompany"
        TempRec.Open TempStr, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not TempRec.EOF
            CboGComp.AddItem IIf(Len(TempRec!GrpCoName), TempRec!GrpCoName, "")
            TempRec.MoveNext
        Loop
        TempRec.Close
        Set TempRec = Nothing
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
End Sub

Private Sub ReadRecord(strAppend As String, PVRec As ADODB.Recordset)
Const METHOD_NAME = "ReadRecord"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter

    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("strAppendCase", adVarChar, adParamInput, 10, "17")
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strAppend)
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
Private Sub LstFFSRcpt_Click()
Const METHOD_NAME = "LstFFSRcpt_Click"
On Error GoTo debugs

Dim rst3 As New ADODB.Recordset
Dim one As Boolean
Dim StrSql As String
Dim BordxNo As String
Dim ReceiptNo As String
    CmdSave.Enabled = False
    CmdUpdate.Enabled = False
    CmdEdit.Visible = True
    cmdAdd.Visible = True
    Call ClearFramePV
        If LstFFSRcpt.ListItems.Count Then
            If LstFFSRcpt.SelectedItem.SubItems(15) <> "" Then
                TxtRecpNo3 = LstFFSRcpt.SelectedItem.SubItems(9)
                If LstFFSRcpt.SelectedItem.SubItems(10) <> "" Then
                    TxtRecpDate3 = LstFFSRcpt.SelectedItem.SubItems(10)
                End If
                TxtRecpAmt = LstFFSRcpt.SelectedItem.SubItems(5)
                TxtChqNo = LstFFSRcpt.SelectedItem.SubItems(14)
                TxtRemark = LstFFSRcpt.SelectedItem.SubItems(11)
                If LstFFSRcpt.SelectedItem.SubItems(13) <> "" Then
                    CboACMth3 = LstFFSRcpt.SelectedItem.SubItems(13)
                End If
                If LstFFSRcpt.SelectedItem.SubItems(12) <> "" Then
                    CboACYr3 = LstFFSRcpt.SelectedItem.SubItems(12)
                End If
                tmpBordxNo = LstFFSRcpt.SelectedItem.SubItems(2)
                RecIndex = LstFFSRcpt.SelectedItem.SubItems(15)
            End If
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Function SearchRecord()

Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim strAppend As String
Dim GrpCompanyName As String
    strAppend = ""
    If Len(Trim(TxtBordx1)) Then
        strAppend = strAppend + " and BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
    End If
    If Len(Trim(TxtBordx2)) Then
        strAppend = strAppend + " and BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
    End If
    If Mid(CboPayStat, 1, 1) = 1 Then
        strAppend = strAppend + " and ReceiptStatus = 'F'"
    End If
    If Mid(CboPayStat, 1, 1) = 2 Then
        strAppend = strAppend + " and (ReceiptStatus = '' or ReceiptStatus is null) "
    End If
    If Len(Trim(TxtRecpNo1)) Then
        strAppend = strAppend + " AND ReceiptNo >= '" & RemStr(Trim(TxtRecpNo1)) & "'"
    End If
    If Len(Trim(TxtRecpNo2)) Then
        strAppend = strAppend + " AND ReceiptNo <= '" & RemStr(Trim(TxtRecpNo2)) & "'"
    End If
    If Len(Trim(TxtRecpDate1)) > 0 And IsDate(TxtRecpDate1) = True Then
        strAppend = strAppend + " And ReceiptDate >= '" & Format(TxtRecpDate1, "mm/dd/yyyy") & "'"
    End If
    If Len(Trim(TxtRecpDate2)) > 0 And IsDate(TxtRecpDate2) = True Then
        strAppend = strAppend + " And ReceiptDate <= '" & Format(TxtRecpDate2, "mm/dd/yyyy") & "'"
    End If
    If Len(CboACMth1) Then
        strAppend = strAppend + " AND ReceiptAccMth >= '" & Trim(CboACMth1) & "'"
    End If
    If Len(CboACMth2) Then
        strAppend = strAppend + " AND ReceiptAccMth <= '" & Trim(CboACMth2) & "'"
    End If
    If Len(CboACYr1) Then
        strAppend = strAppend + " AND ReceiptAccYr >= '" & Trim(CboACYr1) & "'"
    End If
    If Len(CboACYr2) Then
        strAppend = strAppend + " AND ReceiptAccYr <= '" & Trim(CboACYr2) & "'"
    End If
    If ChkBordx.Value = 1 Then
        If Len(Trim(TxtBordx1)) Then
            strAppend = strAppend + " and BordxRefNo = '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
    ElseIf ChkBordx.Value = 0 Then
        If Len(Trim(TxtBordx1)) Then
            strAppend = strAppend + " and BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
        If Len(Trim(TxtBordx2)) Then
            strAppend = strAppend + " and BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
        End If
    End If
    If IsDate(TxtBordxDate1) = True Then
        strAppend = strAppend & " And BordxDate >= '" & Format(TxtBordxDate1, "mm/dd/yyyy") & "'"
    End If
    If IsDate(TxtBodxDate2) = True Then
        strAppend = strAppend & " And BordxDate <= '" & Format(TxtBodxDate2, "mm/dd/yyyy") & "'"
    End If
    If Len(Trim(CboGComp)) Then
         GrpCompanyName = Replace(CboGComp, "'", "''")
         strAppend = strAppend & " AND GroupCompany = '" & Trim(GrpCompanyName) & "'"
    End If
    Call ReceiptBordxListing(strAppend)
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Sub ReceiptBordxListing(strAppend As String)
Const METHOD_NAME = "ReceiptBordxListing"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim cmd As New ADODB.Command
Dim strAppendCase As String
Dim Current As String
Dim Total As String
Dim sql As String
Dim Check As Boolean
Dim Record As ListItem
Dim tmpRecAmt As Double
Dim tmpFFSFee As Double
    lblTotalAmt = 0
    TxtRecpAmt = 0
    Total = 0
    Current = 0
    LblRecordR = 0
    If ChkBordx.Value = 0 Then
        LstFFSRcpt.ListItems.Clear
    End If
    strAppendCase = "17"
    sql = sql + strAppend
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    Do
        Do Until rst2.EOF
            TxtPVStatus = ""
            If Len(rst2!ReceiptStatus) Then
                TxtPVStatus = rst2!ReceiptStatus
            End If
            
            With rst2
                Set Record = LstFFSRcpt.ListItems.Add(, , "")
                If TxtPVStatus = "F" Then
                    Record.SubItems(1) = "Full"
                ElseIf TxtPVStatus = "P" Then
                    Record.SubItems(1) = "Partial"
                ElseIf TxtPVStatus = "" Then
                    Record.SubItems(1) = ""
                End If
                If Len(!Bordxrefno) Then
                    Record.SubItems(2) = !Bordxrefno
                End If
                If Len(!BordxDate) Then
                    Record.SubItems(3) = !BordxDate
                End If
                If Len(!CLMBordxAmt) Then
                    Record.SubItems(4) = !CLMBordxAmt
                End If
                If Len(!CLMMCOProcFee) Then
                    Record.SubItems(5) = Format(!CLMMCOProcFee, "#,##0.00")
                End If
                If Len(!ReceiptAmt) Then
                    Record.SubItems(6) = Format(!ReceiptAmt, "#,##0.00")
                End If
                tmpRecAmt = IIf(IsNumeric(!ReceiptAmt), !ReceiptAmt, 0)
                tmpFFSFee = IIf(IsNumeric(!CLMMCOProcFee), !CLMMCOProcFee, 0)
                'If Len(!CLMMCOProcFee) Then
                Record.SubItems(7) = Format(CDbl(tmpFFSFee) - CDbl(tmpRecAmt), "#,##0.00")
                'End If
                If Len(!ReceiptAmt) Then
                    Record.SubItems(8) = Format(!ReceiptAmt, "#,##0.00")
                Else
                    Record.SubItems(8) = 0
                End If
                
                
                If Len(!ReceiptNo) Then
                    Record.SubItems(9) = !ReceiptNo
                End If
                If Len(!ReceiptDate) Then
                    Record.SubItems(10) = !ReceiptDate
                End If
                If Len(!ReceiptRemarks) Then
                    Record.SubItems(11) = !ReceiptRemarks
                End If
                If Len(!ReceiptAccYr) Then
                    Record.SubItems(12) = !ReceiptAccYr
                End If
                If Len(!ReceiptAccMth) Then
                    Record.SubItems(13) = !ReceiptAccMth
                End If
                If Len(!ReceiptChqNo) Then
                    Record.SubItems(14) = !ReceiptChqNo
                End If
                If Len(!ReceiptIndex) Then
                    Record.SubItems(15) = !ReceiptIndex
                End If
                If Len(!TotalRecAmt) Then
                    Record.SubItems(16) = !TotalRecAmt
                Else
                    Record.SubItems(16) = 0
                End If

                If Len(!GroupCompany) Then
                    Record.SubItems(17) = !GroupCompany
                Else
                    Record.SubItems(17) = 0
                End If
                lblTotalAmt = CDbl(lblTotalAmt) + CDbl(Record.SubItems(5))
                Current = Current + 1
                LblRecordR = Current
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
    
    If LstFFSRcpt.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        LblRecordR = 0
        CmdExit.Enabled = True
        CmdEdit.Enabled = True
        CmdUpdate.Enabled = True
        'cmdEXPtoExcel.Enabled = True
        CmdSave.Enabled = True
        CmdClear.Enabled = True
    End If
    
    rst2.Close
    Set rst2 = Nothing
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    CmdUpdate.Enabled = True
    'cmdEXPtoExcel.Enabled = True
    CmdSave.Enabled = True
    CmdClear.Enabled = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub LstFFSRcpt_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Const METHOD_NAME = "LstFFSRcpt_ItemCheck"
On Error GoTo debugs
    If Item.Checked = True Then
        If Item.SubItems(1) = "Full" Then
            MsgBox "FFS has been received!", vbCritical
            Item.Checked = False
        Else
            Item.SubItems(8) = CDbl(Item.SubItems(5))
            Item.SubItems(1) = "Full"
            Item.SubItems(7) = 0
            LblChkedAmtR = Format(CDbl(LblChkedAmtR) + CDbl(Item.SubItems(5)), "#,##0.00")
        End If
    Else
        LblChkedAmtR = Format(CDbl(LblChkedAmtR) - CDbl(Item.SubItems(5)), "#,##0.00")
        Item.SubItems(7) = CDbl(Item.SubItems(5))
        Item.SubItems(8) = 0
        Item.SubItems(1) = ""
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub UpdateRecord(RecIndex As String, BordxNo As String, tmpFFSFees As Double)
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
    
    strsqlPay = " Select ReceiptIndex, * from ReceiptFrFFS " & _
                " Where ReceiptIndex = '" & Trim(RecIndex) & "'"
    PAY.Open strsqlPay, MediConnCISPayment, adOpenKeyset, adLockOptimistic
    
    If EditFlag = False Then
        PAY.AddNew
    End If
    With PAY
            !Bordxrefno = Trim(BordxNo)
            !ReceiptNo = TxtRecpNo3
            !ReceiptDate = Format(TxtRecpDate3, "dd/mm/yyyy")
            !ReceiptChqNo = IIf(Len(Trim(TxtChqNo)), TxtChqNo, "")
            !ReceiptAmt = IIf(IsNumeric(tmpFFSFees), tmpFFSFees, 0)
            !ReceiptRemarks = IIf(Len(Trim(TxtRemark)), TxtRemark, "")
            !ReceiptAccMth = IIf(CboACMth3 <> "", Trim(CboACMth3), "")
            !ReceiptAccYr = IIf(CboACYr3 <> "", Trim(CboACYr3), "")
            !ReceiptLastUpdDate = Date
            !ReceiptLastUpdUser = Logon
            !ReceiptStatus = "F"
            
        End With
        PAY.Update
        
        PAY.Close
        Set PAY = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub LstFFSRcpt_KeyDown(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "LstFFSRcpt_KeyDown"
On Error GoTo debugs

Dim rst3 As New ADODB.Recordset
Dim one As Boolean
Dim StrSql As String
Dim BordxNo As String
Dim ReceiptNo As String

    CmdSave.Enabled = False
    CmdUpdate.Enabled = False
    CmdEdit.Visible = True
    cmdAdd.Visible = True
    
    Call ClearFramePV
        If LstFFSRcpt.ListItems.Count Then
            If LstFFSRcpt.SelectedItem.SubItems(15) <> "" Then
                TxtRecpNo3 = LstFFSRcpt.SelectedItem.SubItems(9)
                If LstFFSRcpt.SelectedItem.SubItems(10) <> "" Then
                    TxtRecpDate3 = LstFFSRcpt.SelectedItem.SubItems(10)
                End If
                TxtRecpAmt = LstFFSRcpt.SelectedItem.SubItems(5)
                TxtChqNo = LstFFSRcpt.SelectedItem.SubItems(14)
                TxtRemark = LstFFSRcpt.SelectedItem.SubItems(11)
                If LstFFSRcpt.SelectedItem.SubItems(13) <> "" Then
                    CboACMth3 = LstFFSRcpt.SelectedItem.SubItems(13)
                End If
                If LstFFSRcpt.SelectedItem.SubItems(12) <> "" Then
                    CboACYr3 = LstFFSRcpt.SelectedItem.SubItems(12)
                End If
                tmpBordxNo = LstFFSRcpt.SelectedItem.SubItems(2)
                RecIndex = LstFFSRcpt.SelectedItem.SubItems(15)
            End If
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub LstFFSRcpt_KeyUp(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "LstFFSRcpt_KeyUp"
On Error GoTo debugs
Dim rst3 As New ADODB.Recordset
Dim one As Boolean
Dim StrSql As String
Dim BordxNo As String
Dim ReceiptNo As String

    CmdSave.Enabled = False
    CmdUpdate.Enabled = False
    CmdEdit.Visible = True
    cmdAdd.Visible = True
    
    Call ClearFramePV
        If LstFFSRcpt.ListItems.Count Then
            If LstFFSRcpt.SelectedItem.SubItems(15) <> "" Then
                TxtRecpNo3 = LstFFSRcpt.SelectedItem.SubItems(9)
                If LstFFSRcpt.SelectedItem.SubItems(10) <> "" Then
                    TxtRecpDate3 = LstFFSRcpt.SelectedItem.SubItems(10)
                End If
                TxtRecpAmt = LstFFSRcpt.SelectedItem.SubItems(5)
                TxtChqNo = LstFFSRcpt.SelectedItem.SubItems(14)
                TxtRemark = LstFFSRcpt.SelectedItem.SubItems(11)
                If LstFFSRcpt.SelectedItem.SubItems(13) <> "" Then
                    CboACMth3 = LstFFSRcpt.SelectedItem.SubItems(13)
                End If
                If LstFFSRcpt.SelectedItem.SubItems(12) <> "" Then
                    CboACYr3 = LstFFSRcpt.SelectedItem.SubItems(12)
                End If
                tmpBordxNo = LstFFSRcpt.SelectedItem.SubItems(2)
                RecIndex = LstFFSRcpt.SelectedItem.SubItems(13)
            End If
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
        
End Sub
