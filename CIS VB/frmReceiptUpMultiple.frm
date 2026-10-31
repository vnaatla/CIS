VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReceiptUpMultiple 
   Caption         =   "Upload Receipt (Multiple Receipts)"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   15516
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   15516
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lstReceiptListing 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   15255
      _ExtentX        =   26903
      _ExtentY        =   8911
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Case No. (Excel)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bordx No. (Excel)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Receipt Amt(Excel)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Case No. (CIS)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Bordx No. (CIS)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Amount (CIS)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Case Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bordx Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Amount (Diff)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Receipt Status"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Receipt No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Receipt Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Client PV No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Receipt No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Medix Receipt No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Total Rec Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Cheque No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Receipt Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   1200
      TabIndex        =   35
      Top             =   3000
      Visible         =   0   'False
      Width           =   11775
      Begin VB.TextBox txtAllocRefTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   90
         Top             =   240
         Width           =   1545
      End
      Begin VB.TextBox txtAllocRefFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   89
         Top             =   240
         Width           =   1545
      End
      Begin VB.TextBox txtDB 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   88
         Top             =   240
         Width           =   1905
      End
      Begin VB.TextBox txtT9To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   87
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtT9Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   86
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtT8To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   85
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtT8Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   84
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtT7To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   83
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtT7Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   82
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtT6To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   81
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtT6Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   80
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtT5To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   79
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtT5Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   78
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtT4To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   77
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtT4Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   76
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtT3To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   75
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtT3Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   74
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtT2To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   73
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtT2Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   72
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtT1To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   71
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtT1Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   70
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtJournalTypeTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         MaxLength       =   15
         TabIndex        =   69
         Top             =   2040
         Width           =   1785
      End
      Begin VB.TextBox txtJournalTypeFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   68
         Top             =   2040
         Width           =   1905
      End
      Begin VB.TextBox txtDRCR 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   67
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtAllocated 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         MaxLength       =   15
         TabIndex        =   66
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtTranRefFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   65
         Top             =   1680
         Width           =   1905
      End
      Begin VB.TextBox txtTranRefTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         MaxLength       =   15
         TabIndex        =   64
         Top             =   1680
         Width           =   1785
      End
      Begin VB.TextBox txtAmtFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   63
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtAmtTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   62
         Top             =   600
         Width           =   1545
      End
      Begin VB.TextBox txtPeriodYrTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         MaxLength       =   4
         TabIndex        =   61
         Top             =   960
         Width           =   960
      End
      Begin VB.TextBox txtPeriodYrFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   60
         Top             =   960
         Width           =   1080
      End
      Begin VB.TextBox txtPeriodMthFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   59
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtPeriodMthTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         MaxLength       =   2
         TabIndex        =   58
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox ChkBatchNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   57
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkPeriodMth 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   56
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox ChkPeriodYr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   55
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox ChkGrpComp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   54
         Top             =   3240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtACCCodeFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   53
         Top             =   600
         Width           =   1665
      End
      Begin VB.TextBox txtACCCodeTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         MaxLength       =   30
         TabIndex        =   52
         Top             =   3360
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.TextBox txtOPBatchNo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7680
         MaxLength       =   20
         TabIndex        =   51
         Top             =   3360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtOPBatchNo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8160
         MaxLength       =   20
         TabIndex        =   50
         Top             =   3360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox CboPayor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   49
         Top             =   3360
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox txtJournalNoFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   48
         Top             =   1320
         Width           =   1905
      End
      Begin VB.TextBox txtJournalNoTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         MaxLength       =   15
         TabIndex        =   47
         Top             =   1320
         Width           =   1785
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmReceiptUpMultiple.frx":0000
         Left            =   7680
         List            =   "frmReceiptUpMultiple.frx":0002
         TabIndex        =   46
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmReceiptUpMultiple.frx":0004
         Left            =   7200
         List            =   "frmReceiptUpMultiple.frx":0006
         TabIndex        =   45
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmReceiptUpMultiple.frx":0008
         Left            =   7920
         List            =   "frmReceiptUpMultiple.frx":000A
         TabIndex        =   44
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmReceiptUpMultiple.frx":000C
         Left            =   7680
         List            =   "frmReceiptUpMultiple.frx":000E
         TabIndex        =   43
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboBordxMonthfr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   42
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboBordxMonthTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   41
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboBordxYearfr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5040
         TabIndex        =   40
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboBordxYearTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         TabIndex        =   39
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkAppendRcd 
         Caption         =   "Append Records"
         Height          =   255
         Left            =   10320
         TabIndex        =   38
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cmdLookup 
         Caption         =   ">"
         Height          =   255
         Left            =   3000
         TabIndex        =   37
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtAccountname 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   14160
         MaxLength       =   15
         TabIndex        =   36
         Top             =   1320
         Width           =   4305
      End
      Begin MSMask.MaskEdBox txtBordxDateFr 
         Height          =   315
         Left            =   6000
         TabIndex        =   91
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateTo 
         Height          =   315
         Left            =   6600
         TabIndex        =   92
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
         Caption         =   "Allo. Ref"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   5640
         TabIndex        =   138
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   8040
         TabIndex        =   137
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Database"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   136
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "T9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   10080
         TabIndex        =   135
         Top             =   1695
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   12000
         TabIndex        =   134
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   10080
         TabIndex        =   133
         Top             =   1335
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   12000
         TabIndex        =   132
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   10080
         TabIndex        =   131
         Top             =   975
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   12000
         TabIndex        =   130
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   10080
         TabIndex        =   129
         Top             =   615
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   12000
         TabIndex        =   128
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   10080
         TabIndex        =   127
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   12000
         TabIndex        =   126
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   5880
         TabIndex        =   125
         Top             =   2055
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   8040
         TabIndex        =   124
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   5880
         TabIndex        =   123
         Top             =   1695
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   8040
         TabIndex        =   122
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5880
         TabIndex        =   121
         Top             =   1335
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   8040
         TabIndex        =   120
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   5880
         TabIndex        =   119
         Top             =   975
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   8040
         TabIndex        =   118
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Journal Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   117
         Top             =   2055
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   116
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "DR/CR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   115
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Allocated"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   114
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   113
         Top             =   1695
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Trans. Ref."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   112
         Top             =   1695
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8040
         TabIndex        =   111
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   110
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label38 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   109
         Top             =   975
         Width           =   135
      End
      Begin VB.Label Label38 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   108
         Top             =   975
         Width           =   135
      End
      Begin VB.Label Label38 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   107
         Top             =   975
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6720
         TabIndex        =   106
         Top             =   3240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Bordx Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   105
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   104
         Top             =   3240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Bordx Month"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   103
         Top             =   3360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Bordx Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   102
         Top             =   3240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label35 
         Caption         =   "OP Batch No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   101
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   100
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label30 
         Caption         =   "A/C Period Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   99
         Top             =   3240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label34 
         Caption         =   "A/C Period Month"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   98
         Top             =   3360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   97
         Top             =   1335
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Journal Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   96
         Top             =   1335
         Width           =   1455
      End
      Begin VB.Label Period 
         Caption         =   "Period"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   975
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Account Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   94
         Top             =   615
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Account Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14160
         TabIndex        =   93
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdCheck 
      Caption         =   "Check All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5760
      TabIndex        =   34
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   945
      Left            =   120
      TabIndex        =   6
      Top             =   6600
      Width           =   15255
      Begin VB.TextBox txtMedixRecNo 
         Height          =   285
         Left            =   9480
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cboAccPeriodMth 
         Height          =   315
         ItemData        =   "frmReceiptUpMultiple.frx":0010
         Left            =   7320
         List            =   "frmReceiptUpMultiple.frx":0038
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtChequeNo 
         Height          =   285
         Left            =   4440
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtReceiptNo 
         Height          =   285
         Left            =   10800
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   12600
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   14280
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13440
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtReceiptAmt 
         Height          =   285
         Left            =   11040
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cboAccPeriodYR 
         Height          =   315
         ItemData        =   "frmReceiptUpMultiple.frx":0063
         Left            =   7320
         List            =   "frmReceiptUpMultiple.frx":00EB
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtReceiptDate 
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Top             =   240
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         Caption         =   "Medix Rec No."
         Height          =   255
         Index           =   4
         Left            =   8400
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "A/C Period Mth"
         Height          =   255
         Left            =   6120
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "A/C Period Yr"
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   31
         Top             =   555
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Receipt Amt"
         Height          =   255
         Index           =   3
         Left            =   9840
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Receipt Date"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Cheque Number"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Receipt Number"
         Height          =   255
         Index           =   0
         Left            =   9360
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "DR :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblDebit 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "CR :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblCrebit 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblTotalCR 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10200
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Total CR :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9240
         TabIndex        =   21
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblTotalDR 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   20
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Total DR :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblBalance 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Balance :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   15255
      Begin VB.CommandButton cmpUpload 
         Caption         =   "Verify"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdFileLocation 
         Caption         =   " >"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   435
      End
      Begin VB.TextBox txtImportFilename 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2265
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4560
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         Caption         =   "Filename"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "  Upload Receipt (Multiple Receipts)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   0
      TabIndex        =   139
      Top             =   0
      Width           =   15645
   End
End
Attribute VB_Name = "frmReceiptUpMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_blnDirection(1 To 28) As Boolean
Dim xl As New Excel.Application
Dim xlsheet As Excel.Worksheet
Dim xlwbook As Excel.Workbook
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim tmpPathName As String

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFileLocation_Click()
  
    'If lstSUNListing.ListItems.Count > 0 Then
        CommonDialog1.CancelError = True
        On Error GoTo ErrHandler
        CommonDialog1.InitDir = "C:\"
        CommonDialog1.Flags = cdlOFNHideReadOnly
        CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Dat Files (*.Dat)|*.Dat"
        CommonDialog1.ShowOpen
        txtImportFilename = CommonDialog1.FileTitle
        Exit Sub
    'Else
    '    MsgBox "Please select record!", vbOKOnly + vbCritical
    'End If
  
ErrHandler:
  Exit Sub
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelBordxToPayor(lstReceiptListing)
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdUpdate_Click()
    If lstReceiptListing.ListItems.Count = 0 Then
        MsgBox "Please Upload file..! ", vbCritical
    Else
        cmdUpdate.Enabled = False
        Call SaveBordx
    End If
End Sub

Private Sub txtReceiptDate_LostFocus()
    If IsDate(txtReceiptDate) Then
        cboAccPeriodMth.Text = Format(txtReceiptDate, "m")
        cboAccPeriodYR.Text = Format(txtReceiptDate, "yyyy")
    ElseIf txtReceiptDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDate.SetFocus
    ElseIf txtReceiptDate = "__/__/____" Then
        cboAccPeriodMth.Text = ""
        cboAccPeriodYR.Text = ""
    End If
End Sub


Private Function SaveBordx()

Dim II As Integer
Dim update
Dim BordxNo As String
Dim CaseNo As String
Dim status As String
Dim delimiter1 As Integer
Dim StrSqlBordx As String
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim Bordx As New ADODB.Recordset
Dim tmpCaseNumber As String
Dim tmpReceiptAmt As Double
Dim tmnpReceiptAmtExcel As Double
Dim Recordtmp As ListItem
'Dim update
Dim StrSql As String
Dim REC1 As New ADODB.Recordset
Dim REC2 As New ADODB.Recordset
Dim InvCheck As New ADODB.Recordset
Dim strsqlREC1 As String
Dim strsqlREC2 As String
'Dim strsqlCLM As String
'Dim CLM As New ADODb.Recordset
Dim tmpPayAmt As String
Dim tmpBordxNo As String
Dim tmpUpdClientPVNo As String
Dim i As Integer
Dim tmpRecStatus As Boolean
Dim tmpRecNoExcelUpd As String
Dim tmpMedixRecNoExcelUpd As String
Dim tmpRecAmtExcelUpd As Double
Dim tmpChqExcelUpd As String
Dim tmpRecDateExcelUpd As String
Dim tmpAccPerMthExcelUpd As String
Dim tmpAccPerYrExcelUpd As String

    tmpRecStatus = False
    
   ' If txtReceiptNo <> "" Then
   '
   '     strsqlREC1 = " Select * from ReceiptfrPAY " & _
   '                  " Where ReceiptNo = '" & Trim(txtReceiptNo) & "'"
   '
   '     REC1.Open strsqlREC1, MediConnCISPayment, adOpenKeyset, adLockOptimistic
   '
   '     If REC1.EOF = False Then
   '         tmpRecStatus = True
   '     Else
   '         tmpRecStatus = False
   '     End If
   '
   '     REC1.Close
   '     Set REC1 = Nothing
   ' End If
    
    'If txtReceiptNo = "" Then
    '    MsgBox "Please Enter Receipt Number! ", vbCritical
    'If txtMedixRecNo = "" Then
    '    MsgBox "Please Enter Medix Receipt Number! ", vbCritical
    'ElseIf tmpRecStatus = True Then
     '   MsgBox "Receipt Number Has Been Uploaded! ", vbCritical
    'Else
    
    
        update = MsgBox("Do you want to update the record ?", vbYesNo + vbExclamation)
        
        If update = vbYes Then
        
            For i = 1 To lstReceiptListing.ListItems.Count
                tmpCaseNumber = ""
                tmpReceiptAmt = 0
                tmnpReceiptAmtExcel = 0
                
                tmpCaseNumber = lstReceiptListing.ListItems(i).SubItems(3)
                tmpBordxNo = lstReceiptListing.ListItems(i).SubItems(4)
                tmpReceiptAmt = lstReceiptListing.ListItems(i).SubItems(2)
                'tmnpReceiptAmtExcel = lstReceiptListing.ListItems(i).SubItems(11)
                tmpUpdClientPVNo = lstReceiptListing.ListItems(i).SubItems(12)
                tmpRecNoExcelUpd = lstReceiptListing.ListItems(i).SubItems(13)
                tmpMedixRecNoExcelUpd = lstReceiptListing.ListItems(i).SubItems(14)
                tmpRecAmtExcelUpd = lstReceiptListing.ListItems(i).SubItems(15)
                tmpChqExcelUpd = lstReceiptListing.ListItems(i).SubItems(16)
                tmpRecDateExcelUpd = Format(lstReceiptListing.ListItems(i).SubItems(17), "dd/mm/yyyy")
                tmpAccPerMthExcelUpd = Mid(tmpRecDateExcelUpd, 4, 2)
                tmpAccPerYrExcelUpd = Mid(tmpRecDateExcelUpd, 7, 4)
                
            If lstReceiptListing.ListItems(i).SubItems(6) = "Matched" Then
            
                strsqlREC1 = " Select * from ReceiptfrPAY " & _
                             " Where CASNumber = '" & Trim(tmpCaseNumber) & "'"
                
                REC1.Open strsqlREC1, MediConnCISPayment, adOpenKeyset, adLockOptimistic
                
                strsqlCLM = " Select CASNumber, BordxRefNo, PVRecPayorStatus, CLMBordxAmt " & _
                            " From Claim " & _
                            " Where CASNumber = '" & Trim(tmpCaseNumber) & "'"
                CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
                tmpPayAmt = 0
                
                If REC1.RecordCount = 0 Then
                    REC1.AddNew
                End If
                    REC1.AddNew
                    With REC1
                        tmpPayAmt = IIf(IsNumeric(tmpReceiptAmt), tmpReceiptAmt, 0)
                        !CASNumber = tmpCaseNumber
                        !ReceiptNo = tmpRecNoExcelUpd
                        !ReceiptDate = Format(txtReceiptDate, "dd/mm/yyyy")
                        !ReceiptAmt = IIf(IsNumeric(txtReceiptAmt), txtReceiptAmt, 0)
                        !ChequeNo = ChkStr(txtChequeNo)
                        '!Remark = ChkStr(TxtRemark)
                        !BordxNo = ChkStr(tmpBordxNo)
                        !PAYLastUpdateDate = Date
                        !PAYLastUpdateUser = Logon
                        !RECPAYAmt = IIf(IsNumeric(tmpReceiptAmt), tmpReceiptAmt, 0)
                        !CheckAmt = IIf(IsNumeric(tmpReceiptAmt), tmpReceiptAmt, 0)
                        
                        If cboAccPeriodMth <> "" Then
                            !RECAccPeriodMth = cboAccPeriodMth
                        End If
                        If cboAccPeriodYR <> "" Then
                            !RECAccPeriodYR = cboAccPeriodYR
                        End If
                        !ClientPVNo = tmpUpdClientPVNo
                        !MedixReceiptNo = tmpMedixRecNoExcelUpd
                        REC1.update
                    End With
                    
                    With CLM
                        !PVRecPayorStatus = "F"
                    End With
                    CLM.update
                    REC1.Close
                    Set REC1 = Nothing
                    CLM.Close
                    Set CLM = Nothing
            End If
            
            Next i

                
            MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
    '    End If
    End If
    
End Function

Private Sub cmpUpload_Click()
Dim AccNo As String
Dim TransRef As String
Dim AppAmt As String
Dim AccLine As String
Dim tmpJournalNo As String
Dim commitTrans As Boolean
Dim Val1 As Integer
Dim error As Boolean
Dim Record As ListItem
Dim lSize As Long
Dim FileExists As Integer
Dim SendDate As String
Dim tmpDebitAmt As Double
Dim tmpCreditAmt As Double
Dim objItem As ListItem
Dim dAcctNo As String
Dim dTransRef As String
Dim dAppAmt As String
Dim dAccLine As String
Dim dJournalNo As String

Dim CASSql As String
Dim CISCAS1 As New ADODB.Recordset
Dim tmpClaimNoExcel As String
Dim tmpBordxNoExcel As String
Dim tmpReceiptAmtExcel As Double
Dim tmpClaimNoCIS As String
Dim tmpBordxNoCIS As String
Dim tmpReceiptAmtCIS As Double
Dim tmpCaseStatus As Boolean
Dim tmpBordxStatus As Boolean
Dim tmpAmountDiff As String
Dim tmpReceiptstatus As String
Dim tmpReceiptNo As String
Dim tmpReceiptDate As String
Dim tmpReceiptAmt As String
Dim RecSql As String
Dim REC As New ADODB.Recordset
Dim tmpClientPVNo As String
Dim tmpRecNoExcel As String
Dim tmpMedixRecNoExcel As String

    Screen.MousePointer = vbHourglass
    If txtImportFilename <> "" Then
            Val1 = 2
            error = False
            cmpUpload.Enabled = False
            FileExists = 0
        
            lSize = -1
            lSize = FileLen(CommonDialog1.Filename)
            If lSize <= 0 Then
                FileExists = 0
            ElseIf lSize > 0 Then
                FileExists = 1
            End If
            lblDebit = 0
            lblCrebit = 0
            'FileLen (dlgOpen.Filename)
            lstReceiptListing.ListItems.Clear
            txtImportFilename = CommonDialog1.Filename
            If FileExists = 1 Then
                Set xlwbook = xl.Workbooks.Open(txtImportFilename)
                Set xlsheet = xlwbook.Sheets.Item(1)
                
                Do Until Val1 = 60000
                    If xlsheet.Cells(Val1, 2) <> "" Then
                        error = False
                        tmpBordxNoExcel = xlsheet.Cells(Val1, 2)
                        tmpClaimNoExcel = xlsheet.Cells(Val1, 3)
                        'tmpClaimNoExcel = xlsheet.Cells(Val1, 1)
                        'tmpBordxNoExcel = xlsheet.Cells(Val1, 2)
                        tmpReceiptAmtExcel = Format(xlsheet.Cells(Val1, 6), "#,##0.00")
                        tmpReceiptAmtExcel = Format(Trim(tmpReceiptAmtExcel), "#,###0.00")
                        tmpClientPVNo = xlsheet.Cells(Val1, 7)
                        tmpRecNoExcel = xlsheet.Cells(Val1, 8)
                        tmpMedixRecNoExcel = xlsheet.Cells(Val1, 9)
                        
                        CASSql = "SELECT CASNumber,BordxRefNo,CLMBordxAmt FROM Claim WHERE CASNumber='" & Trim(tmpClaimNoExcel) & "'"
                        CISCAS1.Open CASSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic
                       
                            tmpClaimNoCIS = ""
                            tmpBordxNoCIS = ""
                            tmpReceiptAmtCIS = 0
                            
                        If CISCAS1.EOF = False Then
                            tmpClaimNoCIS = CISCAS1!CASNumber
                            tmpBordxNoCIS = CISCAS1!Bordxrefno
                            tmpReceiptAmtCIS = CISCAS1!CLMBordxAmt
                       
                            CISCAS1.Close
                            Set CISCAS1 = Nothing
                        End If
                        
                        ' adOpenForwardOnly, adLockReadOnly
                        
                         ' RecSql = " Select * from ReceiptfrPAY " & _
                         '    " Where CASNumber = '" & Trim(tmpClaimNoExcel) & "'"
               '
               '             REC.Open RecSql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
               '
               '                 tmpReceiptstatus = ""
               '                 tmpReceiptNo = ""
               '                 tmpReceiptDate = ""
               '                 tmpReceiptAmt = 0
               '
               '             If REC.EOF = False Then
               '
               '                 tmpReceiptstatus = IIf(IsNull(REC!RECStatus), "F", REC!RECStatus)
               '                 tmpReceiptNo = IIf(Len(REC!ReceiptNo), REC!ReceiptNo, "")
               '                 tmpReceiptDate = IIf(Len(REC!ReceiptDate), REC!ReceiptDate, "")
               '                 tmpReceiptAmt = IIf(Len(REC!ReceiptAmt), REC!ReceiptAmt, "0")
                            'REC.MoveNext
                            
                            'Loop
                            'Else

                            'End If
                            'REC.Close
                            'Set REC = Nothing
                            
                    '    Else
                    '        tmpClaimNoCIS = ""
                    '        tmpBordxNoCIS = ""
                    '        tmpReceiptAmtCIS = "0"
                    '        tmpCaseStatus = False
                    '        tmpBordxStatus = False
                    '        tmpAmountDiff = tmpReceiptAmtExcel
                    '    End If
                        

                        
                                If tmpClaimNoCIS = "" Then
                                    tmpCaseStatus = False
                                    tmpBordxStatus = False
                                    tmpAmountDiff = tmpReceiptAmtExcel
                                Else
                                    If tmpClaimNoCIS = tmpClaimNoExcel Then
                                        tmpCaseStatus = True
                                    Else
                                        tmpCaseStatus = False
                                    End If
                                    
                                    If tmpBordxNoCIS = tmpBordxNoCIS Then
                                        tmpBordxStatus = True
                                    Else
                                        tmpBordxStatus = False
                                    End If
                                    
                                    If tmpReceiptAmtCIS <> tmpReceiptAmtExcel Then
                                        tmpAmountDiff = tmpReceiptAmtCIS - tmpReceiptAmtExcel
                                    Else
                                        tmpAmountDiff = 0
                                    End If
                                End If
                                
                                
                                
                                                    
                                Set Record = lstReceiptListing.ListItems.Add(, , tmpClaimNoExcel)
                                    cmdUpdate.Enabled = True
                                    Record.SubItems(1) = tmpBordxNoExcel
                                    Record.SubItems(2) = Format(tmpReceiptAmtExcel, "#,##0.00")
                                    Record.SubItems(3) = tmpClaimNoCIS
                                    Record.SubItems(4) = tmpBordxNoCIS
                                    Record.SubItems(5) = Format(tmpReceiptAmtCIS, "#,##0.00")
                                    If tmpCaseStatus = True Then
                                        Record.SubItems(6) = "Matched"
                                    Else
                                        Record.SubItems(6) = "Not Matched"
                                    End If
                                    If tmpBordxStatus = True Then
                                        Record.SubItems(7) = "Matched"
                                    Else
                                        Record.SubItems(7) = "Not Matched"
                                    End If
                                    
                                    Record.SubItems(8) = tmpAmountDiff
                                    'If tmpCaseStatus = True Then
                                        'Record.SubItems(9) = tmpReceiptstatus
                                        Record.SubItems(9) = ""
                                    'End If
                                    'If tmpCaseStatus = True Then
                                        'Record.SubItems(10) = tmpReceiptNo
                                        Record.SubItems(10) = ""
                                    'End If
                                    'If tmpCaseStatus = True Then
                                        'Record.SubItems(11) = tmpReceiptDate
                                        Record.SubItems(11) = ""
                                    'End If
                                    'If tmpCaseStatus = True Then
                                        Record.SubItems(12) = tmpClientPVNo
                                    'End If
                                        Record.SubItems(13) = tmpRecNoExcel
                                        Record.SubItems(14) = tmpMedixRecNoExcel
                    Else
                        
                        Screen.MousePointer = vbDefault
                        cmpUpload.Enabled = True
                        Exit Sub
                    End If
                    
                    
                    'End If
                    Val1 = Val1 + 1
                    Loop
            End If
        Else
            Screen.MousePointer = vbDefault
            cmpUpload.Enabled = True
            MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "frmReceiptUpMultiple")
    txtDB.Text = "UHC"
End Sub

