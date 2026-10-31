VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReceiptMember 
   Caption         =   "Receipt from Member"
   ClientHeight    =   8772
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8772
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   63
      Top             =   6480
      Width           =   11775
      Begin VB.TextBox TxtRecRemarks 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7680
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   240
         Width           =   3825
      End
      Begin VB.ComboBox cboAccPeriodMth 
         Height          =   315
         ItemData        =   "FrmReceiptMember.frx":0000
         Left            =   5640
         List            =   "FrmReceiptMember.frx":0002
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtAmt 
         Height          =   285
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtReceiptNo 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtReceiptDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   27
         Top             =   480
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtChequeNo 
         Height          =   285
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   29
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cboAccPeriodYR 
         Height          =   315
         ItemData        =   "FrmReceiptMember.frx":0004
         Left            =   5640
         List            =   "FrmReceiptMember.frx":0006
         TabIndex        =   31
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Receipt No"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Receipt Date"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Cheque No"
         Height          =   255
         Left            =   2280
         TabIndex        =   68
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Amount"
         Height          =   255
         Left            =   2280
         TabIndex        =   67
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   6960
         TabIndex        =   66
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "A/C Period Yr"
         Height          =   255
         Left            =   4440
         TabIndex        =   65
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "A/C Period Mth"
         Height          =   255
         Left            =   4440
         TabIndex        =   64
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1035
      Left            =   0
      TabIndex        =   71
      Top             =   7200
      Width           =   11775
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   1440
         TabIndex        =   80
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtCHKAmt 
         Height          =   285
         Left            =   4080
         TabIndex        =   79
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox lblCurRecord 
         Height          =   285
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   330
         Left            =   6000
         TabIndex        =   33
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   330
         Left            =   8040
         TabIndex        =   36
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   11040
         TabIndex        =   40
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10440
         TabIndex        =   39
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   8760
         TabIndex        =   37
         Top             =   180
         Width           =   735
      End
      Begin VB.TextBox txtPVStatus 
         Height          =   285
         Left            =   240
         TabIndex        =   72
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Height          =   330
         Left            =   6600
         TabIndex        =   34
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   7320
         TabIndex        =   35
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdEXPtoExcel 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9480
         TabIndex        =   38
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   1080
         TabIndex        =   75
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1440
         TabIndex        =   74
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Checked Amt"
         Height          =   255
         Left            =   3000
         TabIndex        =   73
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11925
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Receipt from Member"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   90
         Width           =   3285
      End
   End
   Begin MSComctlLib.ListView lstReceiptMBM 
      Height          =   4095
      Left            =   0
      TabIndex        =   62
      Top             =   2400
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   7218
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Reference No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Pymt Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "WKS No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amt Rec'd"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "OS Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Amt Receiving"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ReceiptIndex"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Rept No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Rept Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Acc Period"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1980
      Left            =   0
      TabIndex        =   42
      Top             =   360
      Width           =   11775
      Begin VB.TextBox txtRECNoFr 
         Height          =   285
         Left            =   8280
         MaxLength       =   20
         TabIndex        =   8
         Top             =   300
         Width           =   1215
      End
      Begin VB.CheckBox ChkRecpNo 
         Height          =   255
         Left            =   11235
         TabIndex        =   10
         Top             =   270
         Width           =   375
      End
      Begin VB.CheckBox chkPeriodMth 
         Height          =   255
         Left            =   11235
         TabIndex        =   22
         Top             =   1350
         Width           =   375
      End
      Begin VB.CheckBox chkDRAmt 
         Height          =   255
         Left            =   11235
         TabIndex        =   19
         Top             =   1095
         Width           =   375
      End
      Begin VB.CheckBox chkDRNo 
         Height          =   255
         Left            =   11235
         TabIndex        =   16
         Top             =   810
         Width           =   375
      End
      Begin VB.CheckBox chkRecpDate 
         Height          =   255
         Left            =   11235
         TabIndex        =   13
         Top             =   540
         Width           =   375
      End
      Begin VB.CheckBox ChkPeriodYr 
         Height          =   255
         Left            =   11235
         TabIndex        =   25
         Top             =   1665
         Width           =   375
      End
      Begin VB.TextBox txtRECNoTo 
         Height          =   285
         Left            =   9960
         MaxLength       =   20
         TabIndex        =   9
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox TxtCaseNo1 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   300
         Width           =   1510
      End
      Begin VB.TextBox txtCaseNo2 
         Height          =   285
         Left            =   3840
         TabIndex        =   2
         Top             =   300
         Width           =   1510
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   5400
         TabIndex        =   3
         Top             =   280
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtMBMNo 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   4
         Top             =   540
         Width           =   3560
      End
      Begin VB.TextBox TxtMBMName 
         Height          =   285
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   5
         Top             =   795
         Width           =   3560
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   1050
         Width           =   3560
      End
      Begin VB.ComboBox CboPayment 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1335
         Width           =   3560
      End
      Begin VB.TextBox txtRECDateFr 
         Height          =   285
         Left            =   8280
         MaxLength       =   20
         TabIndex        =   11
         Top             =   555
         Width           =   1215
      End
      Begin VB.TextBox txtRECDateTo 
         Height          =   285
         Left            =   9960
         MaxLength       =   20
         TabIndex        =   12
         Top             =   555
         Width           =   1215
      End
      Begin VB.TextBox txtDRNoFr 
         Height          =   285
         Left            =   8280
         MaxLength       =   20
         TabIndex        =   14
         Top             =   810
         Width           =   1215
      End
      Begin VB.TextBox txtDrNoTo 
         Height          =   285
         Left            =   9960
         MaxLength       =   20
         TabIndex        =   15
         Top             =   810
         Width           =   1215
      End
      Begin VB.TextBox txtDRAmtFr 
         Height          =   285
         Left            =   8280
         MaxLength       =   20
         TabIndex        =   17
         Top             =   1065
         Width           =   1215
      End
      Begin VB.TextBox txtDRAmtTo 
         Height          =   285
         Left            =   9960
         MaxLength       =   20
         TabIndex        =   18
         Top             =   1065
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         Height          =   315
         ItemData        =   "FrmReceiptMember.frx":0008
         Left            =   8280
         List            =   "FrmReceiptMember.frx":000A
         TabIndex        =   20
         Top             =   1305
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         Height          =   315
         ItemData        =   "FrmReceiptMember.frx":000C
         Left            =   9960
         List            =   "FrmReceiptMember.frx":000E
         TabIndex        =   21
         Top             =   1305
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         Height          =   315
         ItemData        =   "FrmReceiptMember.frx":0010
         Left            =   8280
         List            =   "FrmReceiptMember.frx":0012
         TabIndex        =   23
         Top             =   1590
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         Height          =   315
         ItemData        =   "FrmReceiptMember.frx":0014
         Left            =   9960
         List            =   "FrmReceiptMember.frx":0016
         TabIndex        =   24
         Top             =   1590
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Mem Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Mem No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   585
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Payment Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Group Company"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         TabIndex        =   57
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "DR No From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   56
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         TabIndex        =   55
         Top             =   615
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Recp Date from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   54
         Top             =   555
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Recp No from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   53
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   9600
         TabIndex        =   52
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label25 
         Caption         =   "A/C Period Yr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   51
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         TabIndex        =   50
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         TabIndex        =   49
         Top             =   1395
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "A/C Period Mth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   48
         Top             =   1395
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "DR Amt From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   47
         Top             =   1110
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         TabIndex        =   46
         Top             =   1110
         Width           =   375
      End
      Begin VB.Label Label27 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   11235
         TabIndex        =   45
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label29 
         Caption         =   "Case No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label30 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   43
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.Label Label33 
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
      Left            =   0
      TabIndex        =   77
      Top             =   8250
      Width           =   1695
   End
   Begin VB.Label Label39 
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
      Height          =   255
      Left            =   2280
      TabIndex        =   76
      Top             =   8250
      Width           =   2295
   End
End
Attribute VB_Name = "FrmReceiptMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CASNo As String
Dim CaseNoArray() As String
Public intExit As Integer
Public Source As Integer
Public intMBMAccount As Integer
Dim one As Boolean
Dim ScanCode As String
Dim listloop1 As Integer
Private m_blnDirection(1 To 12) As Boolean

Private Sub ChkScan_Click()
lstReceiptMBM.ListItems.Clear
If ChkScan = 1 Then
    Label30.Visible = False
    txtCaseNo2.Visible = False
    ChkScan.Left = 3480
    TxtCaseNo1 = ""
    TxtCaseNo1.SetFocus
Else
    Label30.Visible = True
    txtCaseNo2.Visible = True
    ChkScan.Left = 5400
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

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    ReDim CaseNoArray(0)
    CboPayment.ListIndex = 0
    CmdSave.Enabled = True
    one = False
    lstReceiptMBM.ListItems.Clear
    txtCHKAmt = 0
    lblCurRecord = 0
    lblTotal = 0
    ChkScan.Value = 1
End Sub

Private Sub cmdEdit_Click()
    one = True
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

'************************Start ComboListing**************************

Private Sub ComboListing()
Const METHOD_NAME = "ComboListing"
On Error GoTo debugs
Dim Company As New ADODB.Recordset
Dim MonthStart
Dim YearStart
        Company.Open "Select GRPCoName from GRPCompany", MediConnCISDB, adOpenKeyset, adLockOptimistic
        Do Until Company.EOF
            CboPayor.AddItem Trim(Company!GrpCoName)
            Company.MoveNext
        Loop
        Company.Close
        Set Company = Nothing
    With CboPayment
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
        .AddItem "3 - Partial Paid"
        .AddItem "4 - Paid & Partial Paid"
        .AddItem "5 - Unpaid & Partial Paid"
    End With
    
    For MonthStart = 1 To 12
        With cboAccPeriodMth
            .AddItem MonthStart
        End With
        With cboACPeriodMthFr
            .AddItem MonthStart
        End With
        With cboACPeriodMthTo
            .AddItem MonthStart
        End With
    Next MonthStart
    
    For YearStart = 1999 To 3000
        With cboAccPeriodYR
            .AddItem YearStart
        End With
        With cboACPeriodYRFr
            .AddItem YearStart
        End With
        With cboACPeriodYRTo
            .AddItem YearStart
        End With
    Next YearStart
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'************************Start ComboListing**************************

Private Function SearchRecord()

Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim strsql As String
Dim StrAppend As String
Dim Sql As String
Dim GrpCompany As String
    
    StrAppend = ""
    GrpCompany = ""
    
    If ChkScan.Value = 1 Then
        If Len(Trim(TxtCaseNo1)) Then
            StrAppend = StrAppend + " And CASNUMBER ='" & ChkStr(TxtCaseNo1) & "'"
        End If
    Else
        If Len(Trim(TxtCaseNo1)) Then
            StrAppend = StrAppend + " And CASNUMBER >='" & ChkStr(TxtCaseNo1) & "'"
        End If
    
        If Len(Trim(txtCaseNo2)) Then
            StrAppend = StrAppend + " And CASNUMBER <='" & ChkStr(txtCaseNo2) & "'"
        End If
    End If
    
    If Len(Trim(txtMBMNo)) Then
        StrAppend = StrAppend + " And MBMNumber like '%" & ChkStr(txtMBMNo) & "%'"
    End If
    
    If Len(Trim(TxtMBMName)) Then
        StrAppend = StrAppend + " And MBMName like '%" & ChkStr(TxtMBMName) & "%'"
    End If
    
    'If Len(Trim(CboPayor)) Then
        'StrAppend = StrAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
    'End If
    If Len(Trim(GrpCompany)) Then
        StrAppend = StrAppend + " And GroupCompany = '" & Trim(GrpCompany) & "'"
    End If
    
    If Mid(CboPayment, 1, 1) = 1 Then
        StrAppend = StrAppend + " and PVRecMBMStatus = 'F'"
    End If
 
    If Mid(CboPayment, 1, 1) = 2 Then
      StrAppend = StrAppend + " and (PVRecMBMStatus = '' or PVRecMBMStatus is null) "
    End If
    
    If Mid(CboPayment, 1, 1) = 3 Then
      StrAppend = StrAppend + " and PVRecMBMStatus = 'P'"
    End If
    
    If Mid(CboPayment, 1, 1) = 4 Then
      StrAppend = StrAppend + " and (PVRecMBMStatus = 'P' or PVRecMBMStatus = 'F') "
    End If
    
    If Mid(CboPayment, 1, 1) = 5 Then
      StrAppend = StrAppend + " and (PVRecMBMStatus = '' or PVRecMBMStatus is null or PVRecMBMStatus = 'P') "
    End If
        
    If Len(Trim(txtRECNoFr)) Then
        StrAppend = StrAppend + " AND ReceiptNo >= '" & RemStr(Trim(txtRECNoFr)) & "'"
    End If
    
    If Len(Trim(txtRECNoFr)) Then
        StrAppend = StrAppend + " AND ReceiptNo <= '" & RemStr(Trim(txtRECNoFr)) & "'"
    End If
    
    If ChkRecpNo.Value = 1 Then
        Sql = ""
        Sql = " AND (ReceiptNo Is null OR ReceiptNo = '')" & ""
        StrAppend = StrAppend + Sql
    End If
    
    If Len(Trim(txtRECDateFr)) > 0 And IsDate(txtRECDateFr) = True Then
        StrAppend = StrAppend + " And ReceiptDate >= '" & Format(txtRECDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtRECDateTo)) > 0 And IsDate(txtRECDateTo) = True Then
        StrAppend = StrAppend + " And ReceiptDate <= '" & Format(txtRECDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(txtDRNoFr) Then
        StrAppend = StrAppend + " AND DebitCreditRefNo >= '" & Trim(txtDRNoFr) & "'"
    End If

    If Len(txtDrNoTo) Then
        StrAppend = StrAppend + " AND DebitCreditRefNo <= '" & Trim(txtDrNoTo) & "'"
    End If
    
    If chkDRNo.Value = 1 Then
        Sql = ""
        Sql = " AND (DebitCreditRefNo Is null OR DebitCreditRefNo = '')" & ""
        StrAppend = StrAppend + Sql
    End If

    If Len(txtDRAmtFr) Then
        StrAppend = StrAppend + " AND CLMPayableByMedix2M >= '" & Trim(Format(txtDRAmtFr, "#,##0.00")) & "'"
    End If
    
    If Len(txtDRAmtTo) Then
        StrAppend = StrAppend + " AND CLMPayableByMedix2M <= '" & Trim(Format(txtDRAmtTo, "#,##0.00")) & "'"
    End If
    
    If chkDRAmt.Value = 1 Then
        Sql = ""
        Sql = " AND (CLMPayableByMedix2M Is null OR CLMPayableByMedix2M = '')" & ""
        StrAppend = StrAppend + Sql
    End If

    If Len(cboACPeriodMthFr) Then
        StrAppend = StrAppend + " AND RECMBMAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
    
    If Len(cboACPeriodMthTo) Then
        StrAppend = StrAppend + " AND RECMBMAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If chkPeriodMth.Value = 1 Then
        Sql = ""
        Sql = " AND RECMBMAccPeriodMth Is null " & ""
        StrAppend = StrAppend + Sql
    End If

    If Len(cboACPeriodYRFr) Then
        StrAppend = StrAppend + " AND RECMBMAccPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
    End If
    
    If Len(cboACPeriodYRTo) Then
        StrAppend = StrAppend + " AND RECMBMAccPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
    
    If ChkPeriodYr.Value = 1 Then
        Sql = ""
        Sql = " AND RECMBMAccPeriodYr Is null " & ""
        StrAppend = StrAppend + Sql
    End If

       
    searchCriteria = StrAppend
    Call ReceiptFrMBMListing(StrAppend)
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Sub ReceiptFrMBMListing(Optional StrAppend As String)
Const METHOD_NAME = "ReceiptFrMBMListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim REC As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim strsqlREC As String
Dim Total As String
Dim Current As String
Dim Check As Boolean
Dim RECFull As Boolean
Dim RECPartial As String
Dim CASNumber As String
Dim WSVersion As String
Dim INVAmt As Variant
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
    Total = 0
    Current = 0
    lstReceiptMBM.ListItems.Clear
    strAppendCase = "1"
    Sql = Sql + StrAppend
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, Sql)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    Do
        Do Until rst2.EOF
        txtPVStatus = ""
        If Trim(rst2!PVRecMBMStatus) <> "" Then
            txtPVStatus = rst2!PVRecMBMStatus
        End If
            
            With rst2
                Set Record = lstReceiptMBM.ListItems.Add(, , "")
                
                If Len(rst2!DebitCreditRefNo) Then
                    Record.SubItems(1) = Trim(rst2!DebitCreditRefNo)
                End If
                
                If txtPVStatus = "F" Then
                    Record.SubItems(2) = "Full"
                ElseIf txtPVStatus = "P" Then
                    Record.SubItems(2) = "Partial"
                Else
                    Record.SubItems(2) = ""
                End If

                                
                If Len(rst2!CASNumber) Then
                    Record.SubItems(3) = Trim(rst2!CASNumber)
                End If
                
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(4) = Format(rst2!CLMPayableByMedix2M, "#,##0.00") * (-1)
                End If
                
                If Len(rst2!ReceiptAmt) Then
                    Record.SubItems(5) = Format(rst2!ReceiptAmt, "#,##0.00")
                End If
                
                CASNumber = rst2!CASNumber
                
                strsqlREC = "Select Top 100 * from ReceiptfrMBM where CASNumber = '" & Trim(CASNumber) & "'"
                
                REC.Open strsqlREC, MediConnCISPayment, adOpenKeyset, adLockOptimistic
                INVAmt = 0
                If txtPVStatus = "P" Then
                    Do Until REC.EOF
                        INVAmt = INVAmt + REC!ReceiptAmt
                    REC.MoveNext
                    Loop
                    If Len(INVAmt) Then
                        Record.SubItems(5) = Format(INVAmt, "#,##0.00")
                        Record.SubItems(6) = Format(rst2!CLMPayableByMedix2M, "#,##0.00") * (-1) - INVAmt
                    End If
                ElseIf txtPVStatus = "" Then
                    If Len(rst2!CLMPayableByMedix2M) Then
                        Record.SubItems(6) = Format(rst2!CLMPayableByMedix2M, "#,##0.00") * (-1)
                    End If
                End If
                
                If Len(rst2!ReceiptIndex) Then
                    Record.SubItems(8) = rst2!ReceiptIndex
                End If
                
                If Len(rst2!ReceiptNo) Then
                    Record.SubItems(9) = rst2!ReceiptNo
                End If
                
                If Len(rst2!ReceiptDate) Then
                    Record.SubItems(10) = rst2!ReceiptDate
                End If
                
                If Len(rst2!RECMBMAccPeriodMth) Or Len(rst2!RECMBMAccPeriodYr) Then
                    Record.SubItems(11) = rst2!RECMBMAccPeriodMth & "/" & rst2!RECMBMAccPeriodYr
                End If
                
                Current = Current + 1
                lblCurRecord = Current
            End With
            REC.Close
            Set REC = Nothing
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
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdEXPtoExcel_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelReceiptMBM(lstReceiptMBM)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSave_Click()
Const METHOD_NAME = "cmdSave_Click"
On Error GoTo debugs

Dim CASNo As String
Dim status As String
Dim InvoiceNo As String
Dim PAYCode As String
Dim ii As Integer
Dim Cnt As Integer
Dim delimiter1, delimiter2 As Integer
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim Update
    Cnt = 0
    If txtReceiptNo = "" Then
        MsgBox "Please Enter Receipt No", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptNo.SetFocus
    ElseIf txtReceiptDate = "__/__/____" Then
        MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDate.SetFocus
    ElseIf txtAmt = "" Then
        MsgBox "Please Enter Amount", vbOKOnly + vbCritical, DialogBoxTitle
        txtAmt.SetFocus
    ElseIf txtChequeNo = "" Then
        MsgBox "Please Enter Cheque No", vbOKOnly + vbCritical, DialogBoxTitle
        txtChequeNo.SetFocus
    Else
        Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
        If Update = vbYes Then
            For ii = 0 To UBound(CaseNoArray)
             If Trim(CaseNoArray(ii)) <> "" Then
                delimiter1 = InStr(1, CaseNoArray(ii), "~")
                CASNo = Trim(Mid(CaseNoArray(ii), 1, IIf(delimiter1 = 0, 10, delimiter1 - 1)))
                delimiter2 = InStr(delimiter1 + 1, CaseNoArray(ii), "~")
                status = Trim(Mid(CaseNoArray(ii), delimiter1 + 1, delimiter2 - delimiter1 - 1))
                PaymentAmt = Trim(Mid(CaseNoArray(ii), delimiter2 + 1, Len(CaseNoArray(ii)) - delimiter2))
                Call UpdateRecord(CASNo, status, PaymentAmt)
             End If
           Next ii
           Erase CaseNoArray
           Arraycnt = 0
           ReDim CaseNoArray(0)
            Call SearchRecord
            MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
        End If
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub cmdSearch_Click()
Const METHOD_NAME = "cmdSearch_Click"
On Error GoTo debugs


    ScanCode = TxtCaseNo1
    If Mid(ScanCode, 1, 2) = "LO" Then
        TxtCaseNo1 = ""
        'Call SearchForm(ScanCode)
    Else
        one = False
        If ChkScan.Value = 1 Then
            If TxtCaseNo1 = "" Then
                MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
                TxtCaseNo1.SetFocus
            Else
            
                Screen.MousePointer = vbHourglass
                
                listrecord = lstReceiptMBM.ListItems.Count
                If listrecord <> 0 Then
                    
                    For listloop1 = 1 To listrecord
                        Call CaseNoCheck
                    Next listloop1
                End If
                    If Check = True Then
                        MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    ElseIf Check = False Then
                        Call SearchRecord
                    End If
                    Check = False
                Screen.MousePointer = Default
                TxtCaseNo1 = ""
                TxtCaseNo1.SetFocus
            End If
            TxtCaseNo1.SetFocus
        Else
            Call SearchRecord
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
End Sub


Private Sub cmdUpdate_Click()
Const METHOD_NAME = "cmdUpdate_Click"
On Error GoTo debugs
Dim RecordSaved
Dim RptMBM As New ADODB.Recordset
Dim UpdateRptMBM As New ADODB.Recordset
Dim strsql As String
    
    If one = False Then
    ElseIf lstReceiptMBM.ListItems.Count = 0 Then
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
        ElseIf txtAmt = "" Then
            MsgBox "Please Enter Cheque Amt!", vbCritical
            txtAmt.SetFocus
        Else
            strsql = "Select Top 100 ReceiptIndex, * From ReceiptFrMBM Where ReceiptIndex = '" & Trim(lstReceiptMBM.SelectedItem.SubItems(8)) & "'"
            RptMBM.Open strsql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
            RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If RptMBM.RecordCount = 0 Then
                    RptMBM.AddNew
                End If
                With RptMBM
                    !ReceiptNo = txtReceiptNo
                    !ReceiptDate = txtReceiptDate
                    !ReceiptAmt = txtAmt
                    !ReceiptChequeNo = txtChequeNo
                    !ReceiptRemarks = Trim(TxtRecRemarks)
                    !ReceiptLastUpdateDate = Date
                    !ReceiptLastUpdateUser = Logon
                    If cboAccPeriodMth <> "" Then
                        !RECMBMAccPeriodMth = cboAccPeriodMth
                    End If
                    If cboAccPeriodYR <> "" Then
                        !RECMBMAccPeriodYr = cboAccPeriodYR
                    End If
            
                    .Update
                End With
                
          
                Call SearchRecord
                one = False
            End If
            RptMBM.Close
            Set RptMBM = Nothing
            
        txtReceiptNo.Text = ""
        txtReceiptDate = "__/__/____"
        txtAmt.Text = ""
        txtChequeNo.Text = ""
        TxtRecRemarks.Text = ""
        cboAccPeriodMth = ""
        cboAccPeriodYR = ""
        CmdSave.Enabled = True
        MsgBox "Record Updated...", vbOKOnly + vbInformation, "Information"
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "FrmReceiptMember")
    one = False
    Call ComboListing
    CboPayment.ListIndex = 0
    ReDim CaseNoArray(0)
    txtCHKAmt = 0
    intExit = 0
    intMBMAccount = 1
    Label30.Visible = False
    txtCaseNo2.Visible = False
    ChkScan.Value = 1
    ChkScan.Left = 3480
End Sub

Private Sub lstReceiptMBM_Click()
Const METHOD_NAME = "lstReceiptMBM_Click"
On Error GoTo debugs
Dim rst3 As New ADODB.Recordset
Dim strsql As String
If one = True Then
CmdSave.Enabled = False
    If lstReceiptMBM.ListItems.Count Then
        txtReceiptNo = ""
        txtReceiptDate.mask = ""
        txtReceiptDate.Text = ""
        txtReceiptDate.mask = "##/##/####"
        txtAmt = ""
        txtChequeNo = ""
        TxtRecRemarks = ""
        cboAccPeriodMth = ""
        cboAccPeriodYR = ""
    strsql = " SELECT Top 100 * From ReceiptFrMBM Where ReceiptIndex = '" & lstReceiptMBM.SelectedItem.SubItems(8) & "'"
    rst3.Open strsql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
    If rst3.RecordCount = 0 Then
        MsgBox "Payment hasn't been made!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        txtReceiptNo.SetFocus
        txtReceiptNo = rst3!ReceiptNo
        txtReceiptDate = rst3!ReceiptDate
        txtAmt = rst3!ReceiptAmt
        txtChequeNo = rst3!ReceiptChequeNo
        TxtRecRemarks = rst3!ReceiptRemarks
        If Len(rst3!RECMBMAccPeriodMth) Then
            cboAccPeriodMth = rst3!RECMBMAccPeriodMth
        End If
        If Len(rst3!RECMBMAccPeriodYr) Then
            cboAccPeriodYR = rst3!RECMBMAccPeriodYr
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
Private Sub lstReceiptMBM_ItemCheck(ByVal Item As MSComctlLib.ListItem)
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
Dim caseno As String
Dim caseversion As String
Dim InvoiceNo As String
Static Arraycnt As Integer
        caseno = Trim(Item.SubItems(3))
        status = Trim(Item.SubItems(2))
        If Item.Checked = True Then
        strsqlREC = "Select Top 100 * From VIEW_CISReceiptFrMBM where CASNumber = '" & Trim(caseno) & "'"
        REC.Open strsqlREC, MediConnCISPayment, adOpenKeyset, adLockPessimistic
        If REC!PVRecMBMStatus = "F" Then
            MsgBox "Payment has been made!", vbCritical
            Item.Checked = False
        Else
            
            If Partial = False Then
                If Item.SubItems(5) = "" Then
                   Item.SubItems(7) = CDbl(Item.SubItems(4))
                Else
                   Item.SubItems(7) = CDbl(Item.SubItems(6))
                End If
                   Item.SubItems(2) = "Full"
           End If
                txtCHKAmt = txtCHKAmt + CDbl(Format(Item.SubItems(7), "#,##0.00"))
                ReDim Preserve CaseNoArray(Arraycnt)
                CaseNoArray(Arraycnt) = Item.SubItems(3) & "~" & Item.SubItems(2) & "~" & Item.SubItems(7)
                Arraycnt = Arraycnt + 1
        End If
        REC.Close
        Set REC = Nothing
    Else
        newbound = 0
        caseno = Trim(Item.SubItems(3))
        strsqlREC = "Select Top 100 * From VIEW_CISReceiptFrMBM where CASNumber = '" & Trim(caseno) & "'" '& _
                    '"And InvNo = '" & Trim(InvoiceNo) & "'"
        REC.Open strsqlREC, MediConnCISPayment, adOpenKeyset, adLockPessimistic
        
        If REC!PVRecMBMStatus = "P" Then
            status = "Partial"
        Else
            status = Trim(Mid(Item.SubItems(2), InStr(1, Item.SubItems(2), "-") + 1, Len(Item.SubItems(3)) - InStr(1, Item.SubItems(3), "-")))
        End If
           
        If UBound(CaseNoArray) <> 0 Then
            
            For i = 0 To UBound(CaseNoArray)
                If (Item.SubItems(3)) <> CaseNoArray(i) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = CaseNoArray(i)
                    newbound = newbound + 1
                End If
            Next
            If Item.SubItems(2) = "Full" Then
                If Item.SubItems(5) = "" Then
                    txtCHKAmt = txtCHKAmt - CDbl(Format(Item.SubItems(4), "#,##0.00"))
                Else
                    txtCHKAmt = txtCHKAmt - CDbl(Format(Item.SubItems(7), "#,##0.00"))
                End If
            Else
                If Len(Item.SubItems(7)) Then
                    txtCHKAmt = txtCHKAmt - CDbl(Format(Item.SubItems(7), "#,##0.00"))
                End If
            End If
            Erase CaseNoArray
            ReDim CaseNoArray(0)
            ReDim tempArray(0)
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
            ReDim tempArray(0)
            
            If REC.RecordCount Then
                
                If REC!PVRecMBMStatus = "P" Then
                    Item.SubItems(2) = "Partial"
                ElseIf REC!PVRecMBMStatus = "F" Then
                    Item.SubItems(2) = "Full"
                Else
                    Item.SubItems(2) = ""
                End If
            Else
                Item.SubItems(2) = ""
            End If
            If REC.RecordCount Then
                Item.SubItems(7) = ""
                
            Else
                Item.SubItems(5) = ""
                Item.SubItems(7) = ""
            End If

        Else
            If Item.SubItems(2) = "Full" Then
                If Item.SubItems(5) = "" Then
                    txtCHKAmt = txtCHKAmt - CDbl(Format(Item.SubItems(4), "#,##0.00"))
                Else
                    txtCHKAmt = txtCHKAmt - CDbl(Format(Item.SubItems(7), "#,##0.00"))
                End If
            ElseIf Item.SubItems(2) = "Partial" Then
                If Len(Item.SubItems(7)) Then
                    txtCHKAmt = txtCHKAmt - CDbl(Format(Item.SubItems(7), "#,##0.00"))
                End If
            End If
    
            Erase CaseNoArray
            ReDim CaseNoArray(0)
            Arraycnt = 0
            
            If REC.RecordCount Then
                If Len(REC!PVRecMBMStatus) Then
                    Item.SubItems(2) = "Partial"
                    Item.SubItems(7) = ""
                Else
                    Item.SubItems(2) = ""
                    Item.SubItems(7) = ""
                End If
            End If
                            
        End If
    
    REC.Close
    Set REC = Nothing
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Sub UpdateRecord(CASNo As String, status As String, PaymentAmt As Variant)
Const METHOD_NAME = "UpdateRecord"
On Error GoTo debugs

Dim Update
Dim strsql As String
Dim REC As New ADODB.Recordset
Dim strsqlREC As String
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
    strsqlREC = " Select Top 100 * from ReceiptfrMBM " & _
                " where CASNumber = '" & Trim(CASNo) & "'"
    REC.Open strsqlREC, MediConnCISPayment, adOpenKeyset, adLockOptimistic
        If REC.RecordCount = 0 Then
            REC.AddNew
        End If
        With REC
            !ReceiptNo = ChkStr(txtReceiptNo)
            !ReceiptDate = txtReceiptDate
            !ReceiptChequeNo = ChkStr(txtChequeNo)
            !ReceiptLastUpdateDate = Date
            !ReceiptLastUpdateUser = Logon
            !CASNumber = ChkStr(CASNo)
            !ReceiptRemarks = ChkStr(TxtRecRemarks)
            
            If cboAccPeriodMth <> "" Then
                !RECMBMAccPeriodMth = cboAccPeriodMth
            End If
            If cboAccPeriodYR <> "" Then
                !RECMBMAccPeriodYr = cboAccPeriodYR
            End If
            REC!PVRecMBMStatus = Mid(status, 1, 1)
            If status = "Full" Then
                If !ReceiptAmt <> PaymentAmt Then
                    !ReceiptAmt = PaymentAmt + !ReceiptAmt
                Else
                    !ReceiptAmt = PaymentAmt
                End If
            Else
                If IsNumeric(!ReceiptAmt) = False Then
                    !ReceiptAmt = 0
                End If
                    !ReceiptAmt = PaymentAmt + !ReceiptAmt
            End If
        End With
        REC.Update
        REC.Close
        Set REC = Nothing
    txtCHKAmt = 0
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub TxtCaseNo1_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtCaseNo1_LostFocus()
    TxtCaseNo1 = ChkStr(TxtCaseNo1)
    CmdSearch.Default = False
End Sub

Private Sub txtCaseNo2_Change()
txtCaseNo2 = ChkStr(txtCaseNo2)
End Sub

Private Sub txtMBMNo_LostFocus()
txtMBMNo = ChkStr(txtMBMNo)
End Sub

Private Sub txtReceiptDate_LostFocus()
    If IsDate(txtReceiptDate) Then
    Else
        If txtReceiptDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDate.SetFocus
        End If
    End If
End Sub

Private Sub lstReceiptMBM_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(12)
    End Select
End Sub

Private Sub cboPayor_Change()
If CboPayor.Text = "" Then
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End If
End Sub

Private Sub cboPayor_KeyPress(KeyAscii As Integer)
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

Private Function CaseNoCheck()
    If Trim(ChkStr(TxtCaseNo1)) = Trim(lstReceiptMBM.ListItems(listloop1).SubItems(3)) Then
        Check = True
    End If
End Function
