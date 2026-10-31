VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPaymentMemberAuto08 
   Caption         =   "Payment Member"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   66
      Top             =   240
      Width           =   11655
      Begin VB.CheckBox checkall 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   113
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox ChkBatchNo 
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
         Left            =   4800
         TabIndex        =   77
         Top             =   495
         Width           =   375
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
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
         Left            =   3000
         TabIndex        =   19
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox TxtWSNoTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   3360
         MaxLength       =   15
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1330
      End
      Begin VB.CheckBox ChkPVNo 
         Height          =   255
         Left            =   11040
         TabIndex        =   76
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtWSNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox txtPVNoTo 
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
         Left            =   9600
         MaxLength       =   20
         TabIndex        =   8
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox txtPVNoFr 
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
         MaxLength       =   20
         TabIndex        =   7
         Top             =   180
         Width           =   1335
      End
      Begin VB.CheckBox ChkPeriodYr 
         Height          =   255
         Left            =   9840
         TabIndex        =   75
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkChequeNo 
         Height          =   255
         Left            =   11040
         TabIndex        =   74
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox chkCRNo 
         Height          =   255
         Left            =   11040
         TabIndex        =   73
         Top             =   975
         Width           =   375
      End
      Begin VB.CheckBox chkPeriodMth 
         Height          =   255
         Left            =   9840
         TabIndex        =   72
         Top             =   2205
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox ChkPVDate 
         Height          =   255
         Left            =   11040
         TabIndex        =   71
         Top             =   480
         Width           =   375
      End
      Begin VB.ComboBox cboACPeriodMthFr 
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
         ItemData        =   "p.frx":0000
         Left            =   6600
         List            =   "p.frx":0028
         TabIndex        =   70
         Top             =   2175
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodMthTo 
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
         ItemData        =   "p.frx":0053
         Left            =   8400
         List            =   "p.frx":007B
         TabIndex        =   69
         Top             =   2175
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodYRFr 
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
         ItemData        =   "p.frx":00A6
         Left            =   6600
         List            =   "p.frx":012E
         TabIndex        =   68
         Top             =   2460
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodYRTo 
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
         ItemData        =   "p.frx":023A
         Left            =   8400
         List            =   "p.frx":02C2
         TabIndex        =   67
         Top             =   2460
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtOPBatchNo1 
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
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   2
         Top             =   510
         Width           =   1200
      End
      Begin VB.TextBox txtOPBatchNo2 
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
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   3
         Top             =   510
         Width           =   1335
      End
      Begin VB.TextBox TxtMBMNo 
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
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   4
         Top             =   765
         Width           =   3015
      End
      Begin VB.TextBox TxtMBMName 
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
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1020
         Width           =   3015
      End
      Begin VB.ComboBox CboPayor 
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
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1275
         Width           =   3015
      End
      Begin MSMask.MaskEdBox txtPVDateTo 
         Height          =   315
         Left            =   9600
         TabIndex        =   10
         Top             =   435
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox txtPVDateFr 
         Height          =   315
         Left            =   7800
         TabIndex        =   9
         Top             =   435
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
      Begin VB.TextBox txtCHENoFr 
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
         MaxLength       =   20
         TabIndex        =   11
         Top             =   705
         Width           =   1335
      End
      Begin VB.TextBox txtCHENoTo 
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
         Left            =   9600
         MaxLength       =   20
         TabIndex        =   12
         Top             =   705
         Width           =   1335
      End
      Begin VB.TextBox txtCRNoFr 
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
         MaxLength       =   20
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtCRNoTo 
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
         Left            =   9600
         MaxLength       =   20
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox CboPayment 
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1560
         Width           =   3015
      End
      Begin VB.ComboBox CboBordxStatus 
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
         ItemData        =   "p.frx":03CE
         Left            =   7800
         List            =   "p.frx":03D0
         TabIndex        =   16
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtBordxFr 
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
         MaxLength       =   20
         TabIndex        =   104
         Top             =   1485
         Width           =   1335
      End
      Begin VB.TextBox txtBordxTo 
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
         Left            =   9600
         MaxLength       =   20
         TabIndex        =   105
         Top             =   1485
         Width           =   1335
      End
      Begin VB.Label lblheckall 
         Caption         =   "Select All"
         Height          =   255
         Left            =   480
         TabIndex        =   114
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Bordx No from"
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
         Left            =   6120
         TabIndex        =   107
         Top             =   1515
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "To"
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
         Left            =   9240
         TabIndex        =   106
         Top             =   1545
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx Status"
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
         Left            =   6120
         TabIndex        =   99
         Top             =   1260
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "To"
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
         Left            =   3000
         TabIndex        =   98
         Top             =   525
         Width           =   375
      End
      Begin VB.Label Label16 
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
         Left            =   120
         TabIndex        =   97
         Top             =   525
         Width           =   1215
      End
      Begin VB.Label Label37 
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
         Left            =   3000
         TabIndex        =   96
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Member No"
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
         TabIndex        =   95
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Member Name"
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
         TabIndex        =   94
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Payment Status"
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
         Left            =   120
         TabIndex        =   93
         Top             =   1605
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Group Company"
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
         TabIndex        =   92
         Top             =   1290
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "PV No from"
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
         Left            =   6120
         TabIndex        =   91
         Top             =   220
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "To"
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
         Left            =   9240
         TabIndex        =   90
         Top             =   200
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   11040
         TabIndex        =   89
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label36 
         Caption         =   "Claim No"
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
         TabIndex        =   88
         Top             =   230
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "To"
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
         Left            =   9240
         TabIndex        =   87
         Top             =   490
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "PV Date from"
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
         Left            =   6120
         TabIndex        =   86
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label25 
         Caption         =   "To"
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
         Left            =   9240
         TabIndex        =   85
         Top             =   990
         Width           =   375
      End
      Begin VB.Label Label26 
         Caption         =   "Credit Note From"
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
         Left            =   6120
         TabIndex        =   84
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "To"
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
         Left            =   9240
         TabIndex        =   83
         Top             =   735
         Width           =   375
      End
      Begin VB.Label Label29 
         Caption         =   "Cheq. No from"
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
         Left            =   6120
         TabIndex        =   82
         Top             =   735
         Width           =   1455
      End
      Begin VB.Label Label30 
         Caption         =   "A/C Period Yr Fr"
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
         Left            =   4920
         TabIndex        =   81
         Top             =   2520
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label31 
         Caption         =   "To"
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
         Left            =   8040
         TabIndex        =   80
         Top             =   2550
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label32 
         Caption         =   "To"
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
         Left            =   8040
         TabIndex        =   79
         Top             =   2250
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label34 
         Caption         =   "A/C Period Mth Fr"
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
         Left            =   4920
         TabIndex        =   78
         Top             =   2235
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   120
      TabIndex        =   47
      Top             =   7560
      Width           =   11655
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   5280
         TabIndex        =   110
         Top             =   895
         Width           =   525
      End
      Begin VB.TextBox lblSelectedAmt 
         Height          =   285
         Left            =   120
         TabIndex        =   109
         Top             =   480
         Width           =   1485
      End
      Begin VB.TextBox lblCurRecord 
         Height          =   285
         Left            =   120
         TabIndex        =   108
         Top             =   240
         Width           =   1485
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
         TabIndex        =   62
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton CmdPreviewPV 
         Caption         =   "Previ&w PV"
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
         Left            =   3480
         TabIndex        =   61
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdCheck 
         Caption         =   "&Check All"
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
         Left            =   10320
         TabIndex        =   60
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
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
         Left            =   10320
         TabIndex        =   59
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
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
         Left            =   10920
         TabIndex        =   58
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
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
         Left            =   9000
         TabIndex        =   57
         Top             =   150
         Width           =   735
      End
      Begin VB.TextBox txtPVStatus 
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
         Left            =   4440
         TabIndex        =   56
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
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
         TabIndex        =   55
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
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
         Left            =   9000
         TabIndex        =   54
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrintPV 
         Caption         =   "Print P&V"
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
         Left            =   7320
         TabIndex        =   53
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
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
         Left            =   9720
         TabIndex        =   52
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Left            =   7320
         TabIndex        =   51
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   330
         Left            =   7320
         TabIndex        =   50
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   8160
         TabIndex        =   49
         Top             =   150
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
         Left            =   8640
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label15 
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
         Left            =   1680
         TabIndex        =   65
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
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
         Left            =   1680
         TabIndex        =   64
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "of"
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
         Left            =   4920
         TabIndex        =   63
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Frame FramePV 
      Enabled         =   0   'False
      Height          =   1755
      Left            =   120
      TabIndex        =   20
      Top             =   5760
      Width           =   11655
      Begin VB.ComboBox Cbovoidpayment 
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
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   112
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TxtPayTo 
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
         Left            =   960
         MaxLength       =   60
         TabIndex        =   18
         Top             =   160
         Width           =   5175
      End
      Begin VB.TextBox txtChequeNo 
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
         Left            =   960
         MaxLength       =   20
         TabIndex        =   17
         Top             =   420
         Width           =   1095
      End
      Begin VB.ComboBox CboACPeriodMth 
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
         ItemData        =   "p.frx":03D2
         Left            =   7080
         List            =   "p.frx":03D4
         TabIndex        =   32
         Top             =   160
         Width           =   800
      End
      Begin VB.ComboBox CboACPeriodYr 
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
         ItemData        =   "p.frx":03D6
         Left            =   8520
         List            =   "p.frx":03D8
         TabIndex        =   31
         Top             =   160
         Width           =   855
      End
      Begin VB.TextBox txtChequeAmt 
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
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   30
         Top             =   420
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
         Left            =   960
         MaxLength       =   500
         TabIndex        =   29
         Top             =   680
         Width           =   5175
      End
      Begin VB.TextBox txtPVNo 
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
         Left            =   960
         MaxLength       =   15
         TabIndex        =   28
         Top             =   930
         Width           =   1095
      End
      Begin VB.ComboBox CboBank 
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
         ItemData        =   "p.frx":03DA
         Left            =   4920
         List            =   "p.frx":03DC
         TabIndex        =   27
         Text            =   "CB - CITIBANK"
         Top             =   930
         Width           =   1215
      End
      Begin VB.TextBox TxtAccountNo 
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
         Left            =   7080
         MaxLength       =   30
         TabIndex        =   26
         Top             =   455
         Width           =   4425
      End
      Begin VB.TextBox TxtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   7080
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   710
         Width           =   4425
      End
      Begin VB.TextBox txtPayIndex 
         Height          =   285
         Left            =   9720
         TabIndex        =   24
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtCaseNo 
         Height          =   285
         Left            =   10320
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txttmpCents 
         Height          =   285
         Left            =   9240
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txttmpDollar 
         Height          =   285
         Left            =   8400
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSMask.MaskEdBox txtChequeDate 
         Height          =   285
         Left            =   3000
         TabIndex        =   33
         Top             =   420
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox txtPVDate 
         Height          =   285
         Left            =   3000
         TabIndex        =   34
         Top             =   930
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label lblvoidstatus 
         Caption         =   "Payment Void"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   111
         Top             =   1320
         Width           =   1095
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
         TabIndex        =   46
         Top             =   700
         Width           =   735
      End
      Begin VB.Label Label35 
         Caption         =   "Chq. Date"
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
         Left            =   2160
         TabIndex        =   45
         Top             =   455
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Bank"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4200
         TabIndex        =   44
         Top             =   980
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "A/C No"
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
         Left            =   6360
         TabIndex        =   43
         Top             =   450
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "PV No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   42
         Top             =   980
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "PV Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2160
         TabIndex        =   41
         Top             =   980
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Chq. Amt"
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
         Left            =   4200
         TabIndex        =   40
         Top             =   455
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Remark"
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
         Left            =   6360
         TabIndex        =   39
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "A/C Mth"
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
         Left            =   6360
         TabIndex        =   38
         Top             =   220
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "A/C Yr"
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
         Left            =   7980
         TabIndex        =   37
         Top             =   220
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Chq. No"
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
         TabIndex        =   36
         Top             =   455
         Width           =   975
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
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   200
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lstPaymentMBM 
      Height          =   3255
      Left            =   120
      TabIndex        =   100
      Top             =   2535
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5741
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
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   37
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1236
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ref. No"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Pymt Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "MBM Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Claim No"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Amt Paid"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "OS Amt"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "PostageFee"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Amt Paying"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Pay To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "PV No"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "PV MBM Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "PV Date"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Cheq. No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "Cheq. Date"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Chq amt"
         Object.Width           =   1764
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
         Text            =   "Account"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Remark"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Batch No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Bordx Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Amt Wording"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Principal Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Bill Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Group Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "Subsidiary"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "Group Category"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "Member No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "Cover ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "BusinessChanel"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "MBMEmployeeNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "CASTreatReason1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "CASTreatReason2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "Date Visit"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label39 
      Caption         =   "* - Search Empty Data"
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
      Left            =   2400
      TabIndex        =   103
      Top             =   8520
      Width           =   2295
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
      TabIndex        =   102
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Payment Member"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   101
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmPaymentMemberAuto08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer
Dim intMBMAccount As Integer
Dim intExit As Integer
Dim EditFlag As Boolean
Dim ScanCode As String
Dim listloop1 As Integer
Dim listloop2 As Integer
Dim check As Boolean
Dim RecordSelected As Integer
Dim PVNumber As String
Dim tmpACCNumber As String
Private m_blnDirection(1 To 37) As Boolean

Private Sub CboACPeriodMth_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboACPeriodMth_KeyPress"
On Error GoTo debugs

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
        
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub CboACPeriodYr_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboACPeriodYr_KeyPress"
On Error GoTo debugs

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
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub Cbovoidpayment_LostFocus()
     If Mid(Cbovoidpayment, 1, 1) = 1 Then
        TxtRemark.Enabled = True
    End If
End Sub


Private Sub checkall_Click()
Const METHOD_NAME = "CheckAll_Click"
On Error GoTo debugs
Dim LstCnt As Integer
    Dim Item As String
    Item = ""
    If lstPaymentMBM.ListItems.Count <> 0 And checkall <> 0 Then
        LstCnt = lstPaymentMBM.ListItems.Count
        For LstCnt = 1 To LstCnt
                Trim (lstPaymentMBM.ListItems(LstCnt).SubItems(4)) <> ""
                lstPaymentMBM.ListItems(LstCnt).Checked = True
                Call ProcessCheckedItem(lstPaymentMBM.ListItems(LstCnt))
        Next LstCnt
    ElseIf lstPaymentMBM.ListItems.Count <> 0 And checkall = 0 Then
         LstCnt = lstPaymentMBM.ListItems.Count
         For LstCnt = 1 To LstCnt
         lstPaymentMBM.ListItems(LstCnt).Checked = False
         lstPaymentMBM.ListItems(LstCnt).SubItems(2) = ""
         Next LstCnt
    ElseIf checkall = 1 Then
        MsgBox "Please search a record.", vbOKOnly + vbCritical
        checkall = 0
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdCheck_Click()
Const METHOD_NAME = "CmdCheck_Click"
On Error GoTo debugs

    If lstPaymentMBM.ListItems.Count <> 0 Then
        SetCheckAllItems True
    Else
        MsgBox "Please select a record.", vbInformation
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub SetCheckAllItems(bState As Boolean)
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim r As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstPaymentMBM.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstPaymentMBM.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
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

Private Sub ChkScan_Click()
    lstPaymentMBM.ListItems.Clear
    If ChkScan.Value = 1 Then
        Label37.Visible = False
       TxtWSNoTo.Visible = False
        ChkScan.Left = 3000
        TxtWSNo = ""
        TxtWSNo.SetFocus
    Else
        Label37.Visible = True
        TxtWSNoTo.Visible = True
        ChkScan.Left = 4800
        TxtWSNo.SetFocus
    End If
End Sub

Private Sub CmdAdd_Click()
Const METHOD_NAME = "CmdAdd_Click"
On Error GoTo debugs

Dim ZZ As Integer
Dim ItemChk As Boolean
    ItemChk = False
    For ZZ = 1 To lstPaymentMBM.ListItems.Count
        If lstPaymentMBM.ListItems(ZZ).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next ZZ
    CmdUpdate.Visible = False
    CmdSave.Visible = False
    If ItemChk Then
        Call ClearFramePV
        Call EnableFramePV(False)
        txtChequeNo.Enabled = True
        cmdAdd.Visible = False
        CmdSave.Visible = True
        CmdSave.Enabled = True
        FramePV.Enabled = True
        CboBank.Enabled = True
        'TxtPayTo.SetFocus
    Else
        MsgBox "Please checked at least 1 Claim  ! ", vbCritical + vbOKOnly, DialogBoxTitle
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdPreviewPV_Click()
Const METHOD_NAME = "CmdPreviewPV_Click"
On Error GoTo debugs

Dim ItemChecked As Boolean
Dim II As Integer
    ItemChecked = False
    For II = 1 To lstPaymentMBM.ListItems.Count
        If lstPaymentMBM.ListItems(II).Checked = True Then
            ItemChecked = True
        Exit For
        End If
    Next II
    
    If ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    Else
        Call PreviewPV
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub cmdPrintChq_Click()
Const METHOD_NAME = "cmdPrintChq_Click"
On Error GoTo debugs

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

        If txtPVNo = "" Then
            MsgBox "Please Enter or Generate PV No", vbOKOnly + vbCritical
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
                objExcel.Visible = False
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\PaymentCheque.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
            
        On Error GoTo HANDLE_ERROR
        ' Don't allow user to affect workbook
    '        strsql = "select CASNumber, ClaimVersion, BordxRefNo, MBMNumber, MBMPatientCovID, " & _
    '                " HLTCode, ClaimCloseDate, PVNo, PVChequeNo, PVDate, MBMPaidAmt, CLMPayableByMedix2M, " & _
    '                " PVBankCharges from View_Payment2MBMNew where PVNo = '" & TxtPVNo & "' AND PVStatus='A'"
    
            
        strsql = "select PVChequeNo, PVChequeDate, PVCLNChequeAmt, " & _
        " PVWordingOne,PVWordingTwo, PVPayto from VIEW_Payment2MBM_Chq where PVChequeNo = '" & Trim(txtChequeNo) & "'"
        chq.Open strsql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
        
        Do While Not chq.EOF
                objWorksheet.Cells(3, "A") = chq!PVPayto
                objWorksheet.Cells(5, "A") = chq!PVWordingOne
                objWorksheet.Cells(7, "A") = chq!PVWordingTwo
                objWorksheet.Cells(5, "E") = Format(chq!PVCLNChequeAmt, "#,##0.00")
                tmpDD = Format(chq!PVChequeDate, "DD")
                tmpMM = Format(chq!PVChequeDate, "MM")
                tmpYY = Format(chq!PVChequeDate, "YY")
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
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            GoTo ExitFunction
    End If
    Screen.MousePointer = vbDefault
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdPrintChqRange_Click()
    frmPrintChequeRange.chqMode = ""
    Load frmPrintChequeRange
    frmPrintChequeRange.chqMode = "MedixPayMBM"
    frmPrintChequeRange.Show vbModal
End Sub

Private Sub CmdPrintPV_Click()
    frmPrintPVRange.chqMode = ""
    Load frmPrintPVRange
    frmPrintPVRange.chqMode = "MedixPayMBM"
    frmPrintPVRange.Show vbModal
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
    Call ComboListing
    CboPayment.ListIndex = 0
    Cbovoidpayment.ListIndex = 0
    CboBordxStatus.ListIndex = 1
    EditFlag = False
    intExit = 0
    intMBMAccount = 1
    lblCurRecord = 0
    lblTotal = 0
    lblSelectedAmt = 0
End Sub

'=============== Click Command Button ============================
Private Sub CmdEdit_Click()
Dim ZZ As Integer
Dim ItemChk As Boolean

    EditFlag = True
    ItemChk = False
    For ZZ = 1 To lstPaymentMBM.ListItems.Count
        If lstPaymentMBM.ListItems(ZZ).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next ZZ

    
    CmdUpdate.Visible = False
    If ItemChk = False And Len(Trim(txtPVNo)) Then
        CmdSave.Enabled = False
        CmdEdit.Visible = False
        CmdUpdate.Visible = True
        FramePV.Enabled = True
        EditFlag = True
        'If Len(Trim(lstPaymentMBM.SelectedItem.SubItems(1))) Then
            Call EnableFramePV(True)
        'End If
        'TxtPayTo.SetFocus
    End If
End Sub

Private Sub CmdUpdate_Click()
Const METHOD_NAME = "CmdUpdate_Click"
On Error GoTo debugs

Dim RecordSaved
Dim PayMBM As New ADODB.Recordset
Dim strsql As String
Dim PaymentAmt As String
Dim PVSql As String
Dim TEmpStatus As String
Dim TempPayAmt As String
Dim ZZ As Integer
Dim ItemChk As Boolean
Dim tmpAmt As String

Dim strUpdPVsql As String
Dim UpdPV As New ADODB.Recordset
Dim CaseNo As String
Dim PymtIndex As String

Dim tmptxtAmtWording As String
tmptxtAmtWording = ""
Dim wordArray() As String
Erase wordArray
Dim y As Integer
Dim lenword As Integer
lenword = 0
Dim tmpUpdWording1 As String
Dim tmpUpdWording2 As String


    ItemChk = False
    For ZZ = 1 To lstPaymentMBM.ListItems.Count
        If lstPaymentMBM.ListItems(ZZ).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next ZZ
    If ItemChk = True Then

        ElseIf lstPaymentMBM.ListItems.Count = 0 Then
            MsgBox "Please select the record", vbOKOnly + vbCritical
        ElseIf Trim(txtPayIndex) = "" Then
            MsgBox "Please select the record", vbOKOnly + vbCritical
        ElseIf TxtPayTo = "" Then
            MsgBox "Please Enter Pay to! ", vbCritical
            TxtPayTo.SetFocus
        ElseIf txtChequeNo = "" Then
            MsgBox "Please Enter Cheque Number!", vbCritical
            txtChequeNo.SetFocus
        ElseIf txtChequeDate = "__/__/____" Then
            MsgBox "Please Enter Cheque Date!", vbCritical
            txtChequeDate.SetFocus
        ElseIf txtChequeAmt = "" Then
            MsgBox "Please Enter Cheque Amt!", vbCritical
            txtChequeAmt.SetFocus
        ElseIf Trim(CboBank) = "" Then
            MsgBox "Please Enter Bank !", vbCritical
            CboBank.SetFocus
        ElseIf Trim(CboACPeriodMth) = "" Then
            MsgBox "Please keyin Account Period Month ! ", vbCritical
            CboACPeriodMth.SetFocus
        ElseIf Trim(CboACPeriodYr) = "" Then
            MsgBox "Please keyin Account Period Year ! ", vbCritical
            CboACPeriodYr.SetFocus
        ElseIf Trim(txtAmtWording) = "" Then
            MsgBox "Please Enter Wording Amount! ", vbCritical
            
        Else
            RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                strUpdPVsql = ""
                strUpdPVsql = " SELECT PVMBMIndex, CASNumber, PVCLNChequeAmt " & _
                         " from Payment2MBM WHERE PVNo ='" & Trim(txtPVNo) & "'"

                UpdPV.Open strUpdPVsql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
                
                If IsNumeric(txtChequeAmt) Then
                    tmpAmt = ""
                    txttmpDollar = ""
                    txttmpCents = ""
                    tmpAmt = Trim(ConvertCurrencyToEnglish(txtChequeAmt.Text, "MedixPayMBM"))
                    If Mid(tmpAmt, 1, 3) = "And" Then
                        tmpAmt = Mid(tmpAmt, 5, Len(tmpAmt))
                    End If
                    txtAmtWording = tmpAmt & " Only"
                End If
                
                tmptxtAmtWording = Trim(txtAmtWording)
                If Mid(tmptxtAmtWording, 1, 3) = "And" Then
                    tmptxtAmtWording = Trim(Mid(tmptxtAmtWording, 5, Len(tmptxtAmtWording)))
                End If
                
                tmpUpdWording1 = ""
                tmpUpdWording2 = ""
                
                wordArray() = Split(tmptxtAmtWording, " ")
                For y = LBound(wordArray) To UBound(wordArray)
                    lenword = Len(tmpUpdWording1) + Len(tmpUpdWording2) + Len(wordArray(y))

                    'lenword = lenword + Len(wordArray(y))
                    
                    If lenword <= 50 Then
                        If Trim(tmpUpdWording1) = "" Then
                            tmpUpdWording1 = wordArray(y)
                        Else
                            tmpUpdWording1 = tmpUpdWording1 + " " + wordArray(y)
                        End If
                        tmpUpdWording1 = Trim(tmpUpdWording1)
                    Else
                        If Trim(tmpUpdWording2) = "" Then
                            tmpUpdWording2 = wordArray(y)
                        Else
                            tmpUpdWording2 = tmpUpdWording2 + " " + wordArray(y)
                        End If
                        tmpUpdWording2 = Trim(tmpUpdWording2)
                    End If
                Next y
                
                If UpdPV.RecordCount Then
                    Do Until UpdPV.EOF
                        CaseNo = ""
                        CaseNo = Trim(UpdPV!CasNumber)
                        PymtIndex = ""
                        PymtIndex = UpdPV!PVMBMIndex
                        If Trim(CaseNo) <> "" And Trim(PymtIndex) <> "" Then
                            Call UpdatePVPaymentMBM(CaseNo, PymtIndex, tmpUpdWording1, tmpUpdWording2)
                        Else
                            MsgBox "Record Update Failed ! ", vbOKOnly
                            Exit Sub
                        End If
                        UpdPV.MoveNext
                    Loop
                End If
                UpdPV.Close
                Set UpdPV = Nothing
                
                CmdSave.Enabled = True
                EditFlag = False
                PVSql = " and PVNo = '" & Trim(txtPVNo) & "'"
                Call ClearFramePV
                lstPaymentMBM.ListItems.Clear
                Call Payment2MBMListing(PVSql)
                CmdUpdate.Visible = False
                CmdEdit.Visible = True
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                MsgBox "Record Updated ! ", vbOKOnly
            End If
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelAnalysis1(lstPaymentMBM)
    Screen.MousePointer = vbDefault
End Sub

Public Sub CmdSearch_Click()
Const METHOD_NAME = "CmdSearch_Click"
On Error GoTo debugs

Dim listrecord As Integer
    
    Screen.MousePointer = vbHourglass
    ScanCode = TxtWSNo
    CmdEdit.Visible = True
    cmdAdd.Visible = True
    CmdSave.Enabled = False
    CmdUpdate.Visible = False
    
    If Mid(ScanCode, 1, 2) = "LO" Then
        TxtWSNo = ""
        Call SearchForm(ScanCode)
    Else
    lblSelectedAmt = 0
    EditFlag = False

        If ChkScan.Value = 1 Then
            If TxtWSNo = "" Then
                MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
                TxtWSNo.SetFocus
            Else
            
                Screen.MousePointer = vbHourglass
                
                listrecord = lstPaymentMBM.ListItems.Count
                If listrecord <> 0 Then
                    
                    For listloop1 = 1 To listrecord
                        Call CaseNoCheck
                    Next listloop1
                End If
                If check = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                ElseIf check = False Then
                    Call SearchRecord
                End If
                check = False
                Screen.MousePointer = Default
                TxtWSNo = ""
                TxtWSNo.SetFocus
            End If
            TxtWSNo.SetFocus
        Else
            Call SearchRecord
        End If
    End If
    Screen.MousePointer = Default
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    CboPayment.ListIndex = 0
    CmdSave.Visible = False
    FramePV.Enabled = False
    lstPaymentMBM.ListItems.Clear
    lblCurRecord = 0
    lblTotal = 0
    lblSelectedAmt = 0
    EditFlag = False
    ChkScan.Value = 1
    cmdAdd.Visible = True
    CmdEdit.Visible = True
    CmdUpdate.Visible = False
    Cbovoidpayment.ListIndex = 0
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
End Sub

Private Sub cmdExit_Click()
        Unload Me
End Sub
'=============== Click Command Button ============================

'************************Start ComboListing**************************

Private Sub ComboListing()
Const METHOD_NAME = "ComboListing"
On Error GoTo debugs

Dim Company As New ADODB.Recordset
Dim monthstart
Dim yearstart
Dim HLT As New ADODB.Recordset
Dim BankRec As New ADODB.Recordset
        BankRec.Open "Banks", medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
        Do Until BankRec.EOF
            CboBank.AddItem ChkStr(BankRec!BankCode) & " - " & ChkStr(BankRec!BankDesc)
            BankRec.MoveNext
        Loop
        BankRec.Close
        Set BankRec = Nothing
    
        Company.Open "Select GRPCoName from GRPCompany", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        
        Do Until Company.EOF
            CboPayor.AddItem Trim(Company!GrpCoName)
            Company.MoveNext
        Loop
        Company.Close
        Set Company = Nothing
        
    With CboBordxStatus
        .AddItem "0 - Not Bordx"
        .AddItem "1 - Bordx"
    End With
    
    With CboPayment
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
        .AddItem "3 - Partial Paid"
        .AddItem "4 - Paid & Partial Paid"
        .AddItem "5 - Unpaid & Partial Paid"
    End With
    
    With Cbovoidpayment
        .AddItem "0 - Select"
        .AddItem "1 - VOID"
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
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub
'************************Start ComboListing**************************


'**********************Start Search Record*******************************
Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim strsql As String
Dim StrAppend As String
Dim sql As String
Dim GrpCompanyName As String
    
    
    StrAppend = ""
    GrpCompanyName = ""
    
    If ChkScan.Value = 1 Then
        StrAppend = StrAppend + " AND CISWorksheet.dbo.ClaimInvoice.CASNumber = '" & Trim(TxtWSNo) & "'"
    Else
        If Len(TxtWSNo) Then
            StrAppend = StrAppend + " AND CISWorksheet.dbo.ClaimInvoice.CASNumber >= '" & Trim(TxtWSNo) & "'"
        End If
    
        If Len(TxtWSNoTo) Then
            StrAppend = StrAppend + " AND CISWorksheet.dbo.ClaimInvoice.CASNumber <= '" & Trim(TxtWSNoTo) & "'"
        End If
    End If
 
    StrAppend = StrAppend + " And CISWorksheet.dbo.ClaimInvoice.INVPayTo = 'M'"
    
    If Len(Trim(TxtMBMNo)) Then
        StrAppend = StrAppend + " And CISDB.dbo.CISCas.MBMnumber = '" & ChkStr(TxtMBMNo) & "'"
    End If
    
    If Len(Trim(TxtMBMName)) Then
        StrAppend = StrAppend + " And MBMName like '%" & ChkStr(TxtMBMName) & "%'"
    End If
    
    If Len(Trim(CboPayor)) Then
         GrpCompanyName = Replace(CboPayor, "'", "''")
         StrAppend = StrAppend + " AND MBMGroupCompany = '" & Trim(GrpCompanyName) & "'"
    End If
    
    If Mid(CboPayment, 1, 1) = 1 Then
      StrAppend = StrAppend + " and PVPayMBMStatus = 'F'"
    End If
    
    If Mid(CboPayment, 1, 1) = 2 Then
      StrAppend = StrAppend + " and (PVPayMBMStatus = '' or PVPayMBMStatus is null) "
    End If
    
    If Mid(CboPayment, 1, 1) = 3 Then
      StrAppend = StrAppend + " and PVPayMBMStatus = 'P'"
    End If
    
    If Mid(CboPayment, 1, 1) = 4 Then
      StrAppend = StrAppend + " and (PVPayMBMStatus = 'P' or PVPayMBMStatus = 'F') "
    End If
    
    If Mid(CboPayment, 1, 1) = 5 Then
      StrAppend = StrAppend + " and (PVPayMBMStatus = '' or PVPayMBMStatus is null or PVPayMBMStatus = 'P') "
    End If

    If Len(txtPVNoFr) Then
        StrAppend = StrAppend + " and PVNo >= '" & Trim(txtPVNoFr) & "'"
    End If
        
    If Len(txtPVNoTo) Then
        StrAppend = StrAppend + " and PVNo <= '" & Trim(txtPVNoTo) & "'"
    End If
    
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

    If ChkPVNo.Value = 1 Then
        sql = ""
        sql = "AND (PVNo Is null OR PVNo = '') " & ""
        StrAppend = StrAppend + sql
    End If
    
    If (txtPVDateFr) <> "__/__/____" And IsDate(txtPVDateFr) = True _
        And (txtPVDateTo) <> "__/__/____" And IsDate(txtPVDateTo) = True Then
        sql = ""
        sql = " AND PVDate >= '" & Format(Trim(txtPVDateFr), "mm/dd/yyyy") & _
              "' And  PVDate <= '" & Format(Trim(txtPVDateTo), "mm/dd/yyyy") & "'"
        StrAppend = StrAppend + sql
    End If
    
    If Len(Trim(txtPVDateFr)) > 0 And IsDate(txtPVDateFr) = True Then
        StrAppend = StrAppend + " And PVDate >= '" & Format(txtPVDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPVDateTo)) > 0 And IsDate(txtPVDateTo) = True Then
        StrAppend = StrAppend + " And PVDate <= '" & Format(txtPVDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If ChkPVDate.Value = 1 Then
        sql = ""
        sql = "AND (PVDate Is null OR PVDate = '') " & ""
        StrAppend = StrAppend + sql
    End If
    
    If Len(txtCHENoFr) Then
        StrAppend = StrAppend + " and PVChequeNo >= '" & Trim(txtCHENoFr) & "'"
    End If
        
    If Len(txtCHENoTo) Then
        StrAppend = StrAppend + " and PVChequeNo <= '" & Trim(txtCHENoTo) & "'"
    End If
    
    If chkChequeNo.Value = 1 Then
        sql = ""
        sql = "AND (PVChequeNo Is null OR PVChequeNo = '') " & ""
        StrAppend = StrAppend + sql
    End If
    
    If Len(txtCRNoFr) Then
        StrAppend = StrAppend + " and DebitCreditRefNo >= '" & Trim(txtCRNoFr) & "'"
    End If
        
    If Len(txtCRNoTo) Then
        StrAppend = StrAppend + " and DebitCreditRefNo <= '" & Trim(txtCRNoTo) & "'"
    End If
    
    If chkCRNo.Value = 1 Then
        sql = ""
        sql = "AND (DebitCreditRefNo Is null OR DebitCreditRefNo = '') " & ""
        StrAppend = StrAppend + sql
    End If
    If Len(cboACPeriodMthFr) Then
        StrAppend = StrAppend + " and PAYAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
        
    If Len(cboACPeriodMthTo) Then
        StrAppend = StrAppend + " and PAYAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If chkPeriodMth.Value = 1 Then
        sql = ""
        sql = " AND PAYAccPeriodMth Is null " & ""
        StrAppend = StrAppend + sql
    End If
    
    If Len(cboACPeriodYRTo) Then
        StrAppend = StrAppend + " and PAYAccPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
    End If
        
    If Len(cboACPeriodYRTo) Then
        StrAppend = StrAppend + " and PAYAccPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
    
    If ChkPeriodYr.Value = 1 Then
        sql = ""
        sql = " AND PAYAccPeriodYr Is null " & ""
        StrAppend = StrAppend + sql
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 0 Then
        StrAppend = StrAppend + " and ClaimBordxStatus = '" & 0 & "'"
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 1 Then
        StrAppend = StrAppend + " and ClaimBordxStatus = " & 1 & ""
    End If
    
    If Len(txtBordxFr) Then
        StrAppend = StrAppend + " AND CISWorksheet.dbo.Claim.BordxRefNo >= '" & Trim(txtBordxFr) & "'"
    End If

    If Len(txtBordxTo) Then
        StrAppend = StrAppend + " AND CISWorksheet.dbo.Claim.BordxRefNo <= '" & Trim(txtBordxTo) & "'"
    End If
    
        
    Call Payment2MBMListing(StrAppend)
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
   
End Function
'**********************End Search Record*******************************

'**********************Start List Record*******************************

Private Sub Payment2MBMListing(Optional StrAppend As String)

Const METHOD_NAME = "Payment2MBMListing"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim strsqlPay As String
Dim PAY As New ADODB.Recordset
Dim INVAmt As Variant
Dim CasNumber As String
Dim WSVersion As String
Dim Record As ListItem
Dim sql As String
Dim Total As String
Dim Current As String
Dim check As Boolean
Dim RECFull As Boolean
Dim RECPartial As String
Dim PatientDB As New ADODB.Recordset
Dim PatientSql As String
Dim tmpPay2MBM As Double
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter

       Total = 0
       Current = 0
        If ChkScan.Value = 0 Then
            lstPaymentMBM.ListItems.Clear
        End If
        Set cmd.ActiveConnection = MediConnCISPayment
        cmd.CommandText = "SP_CISPayment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("PaymentType", adVarChar, adParamInput, 10, "Pay2MBM")
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
        If rst2.EOF = True Then
            MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
        
        Do
            Do Until rst2.EOF
            'Total = rst2.RecordCount
            lblTotal = Total
            txtPVStatus = ""
            If Len(rst2!PVPayMBMStatus) Then
                txtPVStatus = rst2!PVPayMBMStatus
            End If
                
                With rst2
                    Set Record = lstPaymentMBM.ListItems.Add(, , "")
                    
                    If Len(rst2!DebitCreditRefNo) Then
                        Record.SubItems(1) = rst2!DebitCreditRefNo
                    End If
                    
                    If Trim(txtPVStatus) = "F" Then
                        Record.SubItems(2) = "Full"
                    ElseIf Trim(txtPVStatus) = "P" Then
                        Record.SubItems(2) = "Partial"
                    Else
                        Record.SubItems(2) = ""
                    End If
                    If rst2!coverid = "00" Then
                        PatientSql = "Select MBMName as PatientName from MBM where MBMNumber='" & rst2!MBMNumber & "'"
                    Else
                        PatientSql = "Select MBMCName as PatientName from MBMCoveredPersons where MBMCNumber='" & rst2!MBMNumber & "' AND MBMCCoverID ='" & rst2!coverid & "'"
                    End If
                    
                    PatientDB.Open PatientSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                        Do Until PatientDB.EOF
                            If Len(PatientDB!PatientName) Then
                                Record.SubItems(3) = Trim(PatientDB!PatientName)
                            End If
                        PatientDB.MoveNext
                        Loop
                    PatientDB.Close
                    Set PatientDB = Nothing
                    
                    If Len(rst2!CasNumber) Then
                        Record.SubItems(4) = (rst2!CasNumber)
                    End If
                    
                    CasNumber = rst2!CasNumber
                    tmpPay2MBM = IIf(IsNumeric(rst2!CLMPayableByMedix2M), rst2!CLMPayableByMedix2M, 0)
                    Record.SubItems(5) = Format(tmpPay2MBM, "#,##0.00")
    
                    Record.SubItems(6) = IIf(IsNumeric(rst2!MBMPaidAmt), Format(rst2!MBMPaidAmt, "#,##0.00"), 0)
                    Record.SubItems(8) = IIf(IsNumeric(rst2!CLMPostageFee), Format(rst2!CLMPostageFee, "#,##0.00"), 0)
                    
                    If Len(rst2!PVPayto) Then
                        Record.SubItems(10) = rst2!PVPayto
                    End If
                        
                    If Len(rst2!PVNo) Then
                        Record.SubItems(11) = Trim(rst2!PVNo)
                    End If
                    
                    If Len(rst2!PVMBMIndex) Then
                        Record.SubItems(12) = rst2!PVMBMIndex
                    End If
                    
                    If Len(rst2!PVDate) Then
                        Record.SubItems(13) = rst2!PVDate
                    End If
                    
                    If Len(rst2!PVChequeNo) Then
                        Record.SubItems(14) = Trim(rst2!PVChequeNo)
                    End If
                                    
                    If Len(!PVChequeDate) Then
                        Record.SubItems(15) = Trim(rst2!PVChequeDate)
                    End If
                        
                    If Len(rst2!PVCLNChequeAmt) Then
                        Record.SubItems(16) = Format(rst2!PVCLNChequeAmt, "#,##0.00")
                    End If
                    
                    If Len(rst2!PAYAccPeriodMth & rst2!PAYAccPeriodYr) Then
                        Record.SubItems(17) = Format(Trim(rst2!PAYAccPeriodMth & "/" & rst2!PAYAccPeriodYr), "mm/yyyy")
                    End If
                    If Len(rst2!PVPayeeBank) Then
                        Record.SubItems(18) = Trim(rst2!PVPayeeBank)
                    End If
                    If Len(rst2!MBMBankACNo) Then
                        Record.SubItems(19) = Trim(rst2!MBMBankACNo)
                    End If
                   
                    If Len(rst2!Remark) Then
                        Record.SubItems(20) = Trim(rst2!Remark)
                    End If
                    
                    If Len(rst2!OPBatchNo) Then
                        Record.SubItems(21) = Trim(rst2!OPBatchNo)
                    End If
                    
                    If Len(rst2!ClaimBordxStatus) Then
                        If (rst2!ClaimBordxStatus) = "0" Then
                            Record.SubItems(22) = "No"
                        Else
                            Record.SubItems(22) = "Yes"
                        End If
                    End If
                    
                    strsqlPay = "Select * from Payment2MBM where CASNumber = '" & Trim(CasNumber) & "'"
                    PAY.Open strsqlPay, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
                    INVAmt = 0
                    If Trim(txtPVStatus) = "P" Then
                        Do Until PAY.EOF
                            INVAmt = INVAmt + PAY!MBMPaidAmt
                        PAY.MoveNext
                        Loop
                        If Len(INVAmt) Then
                            Record.SubItems(6) = INVAmt
                            Record.SubItems(7) = rst2!CLMPayableByMedix2M - INVAmt
                        End If
                    ElseIf txtPVStatus = "" Then
                        If Len(rst2!CLMPayableByMedix2M) Then
                            Record.SubItems(7) = Record.SubItems(5) - Record.SubItems(6)
                        End If
                    End If
                    
                    If Len(rst2!PVWordingOne) Then
                        Record.SubItems(23) = Trim(rst2!PVWordingOne)
                    End If
                    
                    If Len(rst2!PVWordingTwo) Then
                        Record.SubItems(23) = Record.SubItems(23) & " " & rst2!PVWordingTwo
                    End If
                    
                    If Len(rst2!MBMName) Then
                        Record.SubItems(24) = rst2!MBMName
                    End If
                    
                    If Len(rst2!CLMTotalActual) Then
                        Record.SubItems(25) = Format(rst2!CLMTotalActual, "#,##0.00")
                    End If
                    
                    If Len(rst2!MBMGroupCompany) Then
                        Record.SubItems(26) = rst2!MBMGroupCompany
                    End If
    
                    If Len(rst2!Subsidiary) Then
                        Record.SubItems(27) = rst2!Subsidiary
                    End If
    
                    If Len(rst2!MBMGroupCategory) Then
                        Record.SubItems(28) = rst2!MBMGroupCategory
                    End If
    
    
                    If Len(rst2!MBMNumber) Then
                        Record.SubItems(29) = rst2!MBMNumber
                    End If
    
                    If Len(rst2!coverid) Then
                        Record.SubItems(30) = rst2!coverid
                    End If
    
                    If Len(rst2!BordxRefNo) Then
                        Record.SubItems(31) = rst2!BordxRefNo
                    End If
                    
                    If Len(rst2!BusinessChanel) Then
                        Record.SubItems(32) = rst2!BusinessChanel + " - " + rst2!Descr
                    End If
                    
                    
                    If Len(rst2!MBMEmployeeNo) Then
                        Record.SubItems(33) = rst2!MBMEmployeeNo
                    End If
                    
                    If Len(rst2!CASTreatReason1) Then
                        Record.SubItems(34) = rst2!CASTreatReason1
                    End If
                    
                    If Len(rst2!CASTreatReason2) Then
                       Record.SubItems(35) = rst2!CASTreatReason2
                    End If
                    
                     If Len(rst2!DateVisit) Then
                       Record.SubItems(36) = rst2!DateVisit
                    End If
                    
                    Current = Current + 1
                    lblCurRecord = Current
                End With
                PAY.Close
                Set PAY = Nothing
                rst2.MoveNext
                
                DoEvents
                    If intExit = 1 Then
                       check = False
                       Exit Do
                    End If
            Loop
        Loop Until check = False
        End If
        intExit = 0
        rst2.Close
        Set rst2 = Nothing
        TxtWSNo.SetFocus
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub
'**********************End List Record*******************************
Private Sub lstPaymentMBM_Click()
Const METHOD_NAME = "lstPaymentMBM_Click"
On Error GoTo debugs

    Call ClearFramePV
    If lstPaymentMBM.ListItems.Count Then
        If Trim(lstPaymentMBM.SelectedItem.SubItems(12)) <> "" Then
            CmdSave.Enabled = False
            FramePV.Enabled = False
            CboBank.Enabled = False
            CmdEdit.Visible = True
            cmdAdd.Visible = True
            CmdSave.Enabled = False
            CmdUpdate.Visible = False
            txtPayIndex = lstPaymentMBM.SelectedItem.SubItems(12)
            txtCaseNo = lstPaymentMBM.SelectedItem.SubItems(4)
            txtPVNo = lstPaymentMBM.SelectedItem.SubItems(11)
            If Len(lstPaymentMBM.SelectedItem.SubItems(10)) Then
                TxtPayTo = lstPaymentMBM.SelectedItem.SubItems(10)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(13)) Then
                txtPVDate = lstPaymentMBM.SelectedItem.SubItems(13)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(14)) Then
                txtChequeNo = lstPaymentMBM.SelectedItem.SubItems(14)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(15)) Then
                txtChequeDate = Format(lstPaymentMBM.SelectedItem.SubItems(15), "dd/MM/yyyy")
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(16)) Then
                txtChequeAmt = lstPaymentMBM.SelectedItem.SubItems(16)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(17)) Then
                CboACPeriodMth = Mid(lstPaymentMBM.SelectedItem.SubItems(17), 1, 2)
                CboACPeriodYr = Mid(lstPaymentMBM.SelectedItem.SubItems(17), 4, 4)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(18)) Then
                CboBank = lstPaymentMBM.SelectedItem.SubItems(18)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(19)) Then
                tmpACCNumber = lstPaymentMBM.SelectedItem.SubItems(19)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(20)) Then
                TxtRemark = lstPaymentMBM.SelectedItem.SubItems(20)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(23)) Then
                txtAmtWording = lstPaymentMBM.SelectedItem.SubItems(23)
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
Private Sub lstPaymentMBM_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Const METHOD_NAME = "lstPaymentMBM_ItemCheck"
On Error GoTo debugs


    EditFlag = False
    CmdSave.Enabled = True
    Call ProcessCheckedItem(Item)
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
Const METHOD_NAME = "ProcessCheckedItem"
On Error GoTo debugs
Dim i  As Integer
Dim Found As Boolean
Dim temp
Dim ListLoop As Integer
Dim ListCount As Integer
Dim PAY As New ADODB.Recordset
Dim CaseNo As String
Dim InvoiceNo As String
Dim Pay2MBM As New ADODB.Recordset
Dim AA, CASNum As String
Dim strsqlPay  As String



If Mid(Item.SubItems(32), 1, 1) = "B" Then
    MsgBox ("Please check Payment towards borker")
End If

    CaseNo = Item.SubItems(4)
                                   
    strsqlPay = " Select PVMBMIndex,PVPayMBMStatus From Payment2MBM where CASNumber = '" & Trim(CaseNo) & "'"
    PAY.Open strsqlPay, MediConnCISPayment, adOpenKeyset, adLockPessimistic
            
    If Item.Checked = True Then

        If PAY.RecordCount <> 0 Then
            If Mid(PAY!PVPayMBMStatus, 1, 1) = "F" Then
                MsgBox "Payment has been made!", vbCritical
                Item.Checked = False
                checkall = False
            Else
                Item.SubItems(9) = CDbl(Item.SubItems(7))
                Item.SubItems(2) = "Full"
                lblSelectedAmt = lblSelectedAmt + CDbl(Format(Item.SubItems(9), "#,##0.00"))
            End If
        Else
            If Partial = False Then
                Item.SubItems(9) = IIf(Trim(Item.SubItems(6)) = "", CDbl(Item.SubItems(5)), CDbl(Item.SubItems(7)))
                Item.SubItems(9) = Item.SubItems(9) - IIf(IsNumeric(Item.SubItems(8)), CDbl(Item.SubItems(8)), 0)
                Item.SubItems(2) = "Full"
            End If
            lblSelectedAmt = lblSelectedAmt + CDbl(Format(Item.SubItems(9), "#,##0.00"))
        End If
    Else
        lblSelectedAmt = lblSelectedAmt - CDbl(Format(Item.SubItems(9), "#,##0.00"))
        Item.SubItems(9) = ""
        If PAY.RecordCount Then
            If Trim(PAY!PVPayMBMStatus) = "P" Then
                Item.SubItems(2) = "Partial"
            End If
        Else
            Item.SubItems(2) = ""
        End If
        
    End If
    PAY.Close
    Set PAY = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Sub CmdSave_Click()
Const METHOD_NAME = "CmdSave_Click"
On Error GoTo debugs


Dim CASNo As String
Dim PVStatus As String
Dim PVIndexNo As String
Dim PaymentAmt As String
Dim Update
Dim ItemChecked As Boolean
Dim i As Integer
Dim ZZ As String
Dim TempBank As String
Dim PVSql As String
Dim BankRec As New ADODB.Recordset
Dim BankSql As String
Dim TmpPVCode As String
Dim PYMTSTATUS As Boolean
Dim pymtcnt As Integer
Dim chqcode As String
Dim tmpchqno As String
Dim tmpPVNoFr As String
Dim tmpPVNoto As String
Dim tmppvcnt As Integer
Dim strupdatepymt As String
Dim tmpUpdWording1 As String
Dim tmpUpdWording2 As String

Dim tmptxtAmtWording As String
tmptxtAmtWording = ""
Dim wordArray() As String
Erase wordArray
Dim y As Integer
Dim lenword As Integer

lenword = 0

    ItemChecked = False
    CASNo = ""
    PVStatus = ""
    PaymentAmt = ""
    TmpPVCode = ""
    tmppvcnt = 1
    ZZ = InStr(1, CboBank, "-")
    TempBank = Trim(Left(CboBank, IIf(ZZ - 1 <= 0, Len(Trim(CboBank)), ZZ - 1)))
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        BankSql = "SELECT PVCode from Banks WHERE BankCode='" & Trim(TempBank) & "'"
        BankRec.Open BankSql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
        If BankRec.RecordCount Then
            TmpPVCode = Trim(BankRec!PVCode)
        End If
        BankRec.Close
        Set BankRec = Nothing
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
    
    For i = 1 To lstPaymentMBM.ListItems.Count
        If lstPaymentMBM.ListItems(i).Checked = True Then
            ItemChecked = True
        Exit For
        End If
    Next i
    
    If lstPaymentMBM.ListItems.Count = 0 Then
        MsgBox "Please select record", vbOKOnly + vbCritical
    ElseIf ItemChecked = False Then
        MsgBox "Please select record", vbOKOnly + vbCritical
  '  ElseIf TxtPayTo = "" Then
  '      MsgBox "Please Enter Pay To!", vbCritical
'        TxtPayTo.SetFocus
 '   ElseIf txtPVDate = "__/__/____" Then
 '       MsgBox "Please Enter PV Date!", vbCritical
 '       txtPVDate.SetFocus
 '   ElseIf IsDate(txtPVDate) = False Then
 '       MsgBox "Invalid PV Date!", vbCritical
 '       txtPVDate.SetFocus
    ElseIf Trim(txtChequeNo) = "" And Mid(Cbovoidpayment, 1, 1) <> 1 Then
        MsgBox "Please Enter Cheque Number!", vbCritical
        txtChequeNo.SetFocus
 '   ElseIf txtChequeDate = "__/__/____" Then
 '       MsgBox "Please Enter Cheque Date!", vbCritical
 '       txtChequeDate.SetFocus
 '   ElseIf txtChequeAmt = "" Then
 '       MsgBox "Please Enter Cheque Amt!", vbCritical
 '       txtChequeAmt.SetFocus
 '   ElseIf txtChequeAmt <> lblSelectedAmt And TxtRemark = "" Then
 '       MsgBox "Amount to paid is incorrect" & vbNewLine & "Please key in remark for reason.", vbCritical + vbOKOnly
 '       TxtRemark.SetFocus
 '   ElseIf txtAmtWording = "" Then
 '       MsgBox "Please Enter Wording Amount!", vbCritical
    ElseIf Trim(CboBank) = "" Then
        MsgBox "Please Enter Bank !", vbCritical
        CboBank.SetFocus
 '   ElseIf Trim(CboACPeriodMth) = "" Then
 '       MsgBox "Please keyin Account Period Month ! ", vbCritical
 '       CboACPeriodMth.SetFocus
  '  ElseIf Trim(CboACPeriodYr) = "" Then
  '      MsgBox "Please keyin Account Period Year ! ", vbCritical
  '      CboACPeriodYr.SetFocus
    ElseIf Trim(TmpPVCode) = "" Then
        MsgBox "Please keyin Bank - PV Code ! ", vbCritical
    Else
        Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
        If Update = vbYes Then

            'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
            'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
            
            txtChequeAmt = 0
            tmpPVNoFr = ""
            tmpPVNoto = ""
            pymtcnt = 0
            PYMTSTATUS = False
            If Mid(Cbovoidpayment, 1, 1) = 1 Then
                PVNumber = "VOID"
                txtPVNo = PVNumber
            Else
                PVNumber = GenerateSPVNo(TempBank, TmpPVCode)
                txtPVNo = ""
                txtPVNo = PVNumber
            End If
            'pymtcnt = pymtcnt + 1
            'If CDbl(pymtcnt) = "1" Then
                txtChequeNo = txtChequeNo
                tmpPVNoFr = PVNumber
                tmpPVNoto = PVNumber
            'Else
                'chqcode = Mid(txtChequeNo, 1, 2)
                'tmpchqno = Mid(txtChequeNo, 3, 8)
                'tmpchqno = Format(tmpchqno + 1, "000000")
                'txtChequeNo = chqcode & tmpchqno
                'tmpPVNoto = PVNumber
            'End If

            'If ItemChecked = True Then
                For i = 1 To lstPaymentMBM.ListItems.Count
                    If lstPaymentMBM.ListItems(i).Checked = True Then
                        'If tmppvcnt = 1 Then
                        'End If
                        tmpACCNumber = ""
                        CASNo = Trim(lstPaymentMBM.ListItems(i).SubItems(4))
                        txtChequeAmt = CDbl(txtChequeAmt) + CDbl(lstPaymentMBM.ListItems(i).SubItems(5))
                        PVStatus = Trim(Mid(lstPaymentMBM.ListItems(i).SubItems(2), 1, 1))
                        PaymentAmt = Trim(lstPaymentMBM.ListItems(i).SubItems(9))
                        PVIndexNo = Trim(lstPaymentMBM.ListItems(i).SubItems(12))
                        TxtPayTo = Trim(lstPaymentMBM.ListItems(i).SubItems(24))
                        txtChequeDate = Format(Date, "DD/MM/YYYY")
                        txtPVDate = Format(Date, "DD/MM/YYYY")
                        tmpACCNumber = Trim(lstPaymentMBM.ListItems(i).SubItems(19))
                        Call UpdateRecord(CASNo, PVStatus, PaymentAmt, PVIndexNo)
                                         
                        CASNo = ""
                        PVStatus = ""
                        PaymentAmt = ""
                  End If
                Next i
                
               ' If Len(txtAmtWording) > 55 Then
               '    If Mid(txttmpDollar, 1, 3) = "And" Then
               '        ' !PVWordingOne = Mid(txttmpDollar, 5, 500)
               '         tmpUpdWording1 = Mid(txttmpDollar, 5, 500)
               '     Else
               '         '!PVWordingOne = txttmpDollar
               '         tmpUpdWording1 = Mid(txttmpDollar, 5, 500)
               '     End If
               '     '!PVWordingTwo = IIf(Len(Trim(txttmpCents)), txttmpCents & " Only", "")
               ' Else
               '     '!PVWordingOne = IIf(Len(Trim(txtAmtWording)), Trim(txtAmtWording), "")
               '     tmpUpdWording1 = txtAmtWording
               '     '!PVWordingTwo = ""
               '     tmpUpdWording2 = ""
               ' End If
                
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If Mid(Cbovoidpayment, 1, 1) <> 1 Then
                tmptxtAmtWording = Trim(txtAmtWording)
                If Mid(tmptxtAmtWording, 1, 3) = "And" Then
                    tmptxtAmtWording = Trim(Mid(tmptxtAmtWording, 5, Len(tmptxtAmtWording)))
                End If
                
                tmpUpdWording1 = ""
                tmpUpdWording2 = ""
                

                wordArray() = Split(tmptxtAmtWording, " ")
                For y = LBound(wordArray) To UBound(wordArray)
                    lenword = Len(tmpUpdWording1) + Len(tmpUpdWording2) + Len(wordArray(y))

                    'lenword = lenword + Len(wordArray(y))
                    
                    If lenword <= 50 Then
                        If Trim(tmpUpdWording1) = "" Then
                            tmpUpdWording1 = wordArray(y)
                        Else
                            tmpUpdWording1 = tmpUpdWording1 + " " + wordArray(y)
                        End If
                        tmpUpdWording1 = Trim(tmpUpdWording1)
                    Else
                        If Trim(tmpUpdWording2) = "" Then
                            tmpUpdWording2 = wordArray(y)
                        Else
                            tmpUpdWording2 = tmpUpdWording2 + " " + wordArray(y)
                        End If
                        tmpUpdWording2 = Trim(tmpUpdWording2)
                    End If
                Next y
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    strupdatepymt = ""
                    strupdatepymt = " update Payment2MBM set " & _
                                  " PVChequeDate = '" & Format(txtChequeDate, "yyyy/mm/dd") & "', " & _
                                  " PVChequeNo = '" & txtChequeNo & "', " & _
                                  " PVCLNChequeAmt = " & txtChequeAmt & ", " & _
                                  " PVPayto = '" & TxtPayTo & "', " & _
                                  " PVWordingOne = '" & tmpUpdWording1 & "', " & _
                                  " PVWordingTwo = '" & tmpUpdWording2 & "' " & _
                                  " Where PVNo = '" & Trim(txtPVNo) & "'"
                    MediConnCISPayment.Execute strupdatepymt
        
    
                    lstPaymentMBM.ListItems.Clear
                'End If
                PVSql = " and PVNo >= '" & Trim(tmpPVNoFr) & "' and PVNo <= '" & Trim(tmpPVNoto) & "'"
                Call ClearFramePV
                lstPaymentMBM.ListItems.Clear
                Call Payment2MBMListing(PVSql)
            End If
            'MediConnCISDB.Close
            'Set MediConnCISDB = Nothing
            'medixmasterconn.Close
            'Set medixmasterconn = Nothing
            'MediConnCISWS.Close
            'Set MediConnCISWS = Nothing
            'MediConnCISPayment.Close
            'Set MediConnCISPayment = Nothing
            
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            Cbovoidpayment.ListIndex = 0
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub UpdateRecord(CASNo As String, status As String, PaymentAmt As Variant, PVIndexNo As String)
Const METHOD_NAME = "UpdateRecord"
On Error GoTo debugs

Dim Update
Dim strsql As String
Dim PAY As New ADODB.Recordset
Dim strsqlPay As String
Dim ZZ As String
Dim TempBank As String
    ZZ = InStr(1, CboBank, "-")
    TempBank = Trim(Left(CboBank, IIf(ZZ - 1 <= 0, Len(Trim(CboBank)), ZZ - 1)))
    strsqlPay = " Select PVMBMIndex, * from Payment2MBM " & _
                " Where PVMBMIndex = '" & Trim(PVIndexNo) & "'"
    PAY.Open strsqlPay, MediConnCISPayment, adOpenKeyset, adLockOptimistic
    
    If PAY.RecordCount > 0 Then
        If EditFlag = False Then
            If Trim(PAY!PVPayMBMStatus) = "P" Then
                PAY!PVPayMBMStatus = Mid(status, 1, 1)
                PAY.Update
                PAY.AddNew
            End If
        End If
    Else
        PAY.AddNew
    End If
        
        With PAY
            !CasNumber = ChkStr(CASNo)
            !PVNo = txtPVNo
            !PVDate = IIf(txtPVDate <> "__/__/____", Format(txtPVDate, "dd/mm/yyyy"), "")
            !PVChequeDate = IIf(txtChequeDate <> "__/__/____", Format(txtChequeDate, "dd/mm/yyyy"), "")
            If Mid(Cbovoidpayment, 1, 1) = 1 Then
                !PVChequeNo = "VOID"
            Else
                !PVChequeNo = IIf(Trim(txtChequeNo) <> "", txtChequeNo, "")
            End If
            
            
            '!PVCLNChequeAmt = IIf(IsNumeric(txtChequeAmt), Format(txtChequeAmt, "#,##0.00"), 0)
            !PVPayto = ChkStr(TxtPayTo)
            !PVPayeeBank = IIf(TempBank <> "", TempBank, "")
            !PVPayeeACNo = IIf(tmpACCNumber <> "", tmpACCNumber, "")
            !Remark = IIf(Trim(TxtRemark) <> "", TxtRemark, "")
            !MBMPaidAmt = IIf(IsNumeric(PaymentAmt), PaymentAmt, 0)
            !Pay2MBMLastUpdateDate = Date
            !Pay2MBMLastUpdateUser = Logon
            !PAYAccPeriodMth = IIf(CboACPeriodMth <> "", Trim(CboACPeriodMth), "")
            !PAYAccPeriodYr = IIf(CboACPeriodYr <> "", Trim(CboACPeriodYr), "")
            !PVPayMBMStatus = Mid(status, 1, 1)

        End With
    
    
    PAY.Update
    PAY.Close
    Set PAY = Nothing
    lblSelectedAmt = 0
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub UpdatePVPaymentMBM(CASNo As String, PVIndexNo As String, tmpUpdWording1 As String, tmpUpdWording2 As String)
Const METHOD_NAME = "UpdatePVPaymentMBM"
On Error GoTo debugs

Dim Update
Dim strsql As String
Dim PAY As New ADODB.Recordset
Dim strsqlPay As String

    strsqlPay = " Select PVMBMIndex, * from Payment2MBM " & _
                " Where PVMBMIndex = '" & Trim(PVIndexNo) & "'"
    PAY.Open strsqlPay, MediConnCISPayment, adOpenKeyset, adLockOptimistic
    
    If PAY.RecordCount > 0 Then
        PAY.Update
            With PAY
                !CasNumber = ChkStr(CASNo)
                '!PVNo = txtPVNo
                !PVDate = IIf(txtPVDate <> "__/__/____", Format(txtPVDate, "dd/mm/yyyy"), "")
                !PVChequeDate = IIf(txtChequeDate <> "__/__/____", Format(txtChequeDate, "dd/mm/yyyy"), "")
                !PVChequeNo = IIf(Trim(txtChequeNo) <> "", txtChequeNo, "")
                !PVCLNChequeAmt = IIf(IsNumeric(txtChequeAmt), Format(txtChequeAmt, "#,##0.00"), 0)
                !PVPayto = ChkStr(TxtPayTo)
                '!PVPayeeBank = IIf(TempBank <> "", TempBank, "")
                '!PVPayeeACNo = IIf(TxtAccountNo <> "", TxtAccountNo, "")
                !Remark = IIf(Trim(TxtRemark) <> "", TxtRemark, "")
                '!MBMPaidAmt = IIf(IsNumeric(PaymentAmt), PaymentAmt, 0)
                '!Pay2MBMLastUpdateDate = Date
                '!Pay2MBMLastUpdateUser = Logon
                '!PAYAccPeriodMth = IIf(CboACPeriodMth <> "", Trim(CboACPeriodMth), "")
                '!PAYAccPeriodYr = IIf(CboACPeriodYr <> "", Trim(CboACPeriodYr), "")
                '!PVPayMBMStatus = Mid(status, 1, 1)
                !PVWordingOne = tmpUpdWording1
                !PVWordingTwo = tmpUpdWording2
            End With
        PAY.Update
        PAY.Close
        Set PAY = Nothing
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
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

Private Sub cboBank_LostFocus()
    CboBank = ChkStr(CboBank)
End Sub

Private Sub lstPaymentMBM_KeyDown(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "lstPaymentMBM_KeyDown"
On Error GoTo debugs


    Call ClearFramePV
    If lstPaymentMBM.ListItems.Count Then
        If Trim(lstPaymentMBM.SelectedItem.SubItems(12)) <> "" Then
            CmdSave.Enabled = False
            FramePV.Enabled = False
            CboBank.Enabled = False
            CmdEdit.Visible = True
            cmdAdd.Visible = True
            CmdSave.Enabled = False
            CmdUpdate.Visible = False
            txtPayIndex = lstPaymentMBM.SelectedItem.SubItems(12)
            txtCaseNo = lstPaymentMBM.SelectedItem.SubItems(4)
            txtPVNo = lstPaymentMBM.SelectedItem.SubItems(11)
            If Len(lstPaymentMBM.SelectedItem.SubItems(10)) Then
                TxtPayTo = lstPaymentMBM.SelectedItem.SubItems(10)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(13)) Then
                txtPVDate = lstPaymentMBM.SelectedItem.SubItems(13)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(14)) Then
                txtChequeNo = lstPaymentMBM.SelectedItem.SubItems(14)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(15)) Then
                txtChequeDate = lstPaymentMBM.SelectedItem.SubItems(15)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(16)) Then
                txtChequeAmt = lstPaymentMBM.SelectedItem.SubItems(16)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(17)) Then
                CboACPeriodMth = Mid(lstPaymentMBM.SelectedItem.SubItems(17), 1, 2)
                CboACPeriodYr = Mid(lstPaymentMBM.SelectedItem.SubItems(17), 4, 4)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(18)) Then
                CboBank = lstPaymentMBM.SelectedItem.SubItems(18)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(19)) Then
                tmpACCNumber = lstPaymentMBM.SelectedItem.SubItems(19)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(20)) Then
                TxtRemark = lstPaymentMBM.SelectedItem.SubItems(20)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(23)) Then
                txtAmtWording = lstPaymentMBM.SelectedItem.SubItems(23)
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

Private Sub lstPaymentMBM_KeyUp(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "lstPaymentMBM_KeyUp"
On Error GoTo debugs


    Call ClearFramePV
    If lstPaymentMBM.ListItems.Count Then
        If Trim(lstPaymentMBM.SelectedItem.SubItems(12)) <> "" Then
            CmdSave.Enabled = False
            FramePV.Enabled = False
            CboBank.Enabled = False
            CmdEdit.Visible = True
            cmdAdd.Visible = True
            CmdSave.Enabled = False
            CmdUpdate.Visible = False
            txtPayIndex = lstPaymentMBM.SelectedItem.SubItems(12)
            txtCaseNo = lstPaymentMBM.SelectedItem.SubItems(4)
            txtPVNo = lstPaymentMBM.SelectedItem.SubItems(11)
            If Len(lstPaymentMBM.SelectedItem.SubItems(10)) Then
                TxtPayTo = lstPaymentMBM.SelectedItem.SubItems(10)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(13)) Then
                txtPVDate = lstPaymentMBM.SelectedItem.SubItems(13)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(14)) Then
                txtChequeNo = lstPaymentMBM.SelectedItem.SubItems(14)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(15)) Then
                txtChequeDate = lstPaymentMBM.SelectedItem.SubItems(15)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(16)) Then
                txtChequeAmt = lstPaymentMBM.SelectedItem.SubItems(16)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(17)) Then
                CboACPeriodMth = Mid(lstPaymentMBM.SelectedItem.SubItems(17), 1, 2)
                CboACPeriodYr = Mid(lstPaymentMBM.SelectedItem.SubItems(17), 4, 4)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(18)) Then
                CboBank = lstPaymentMBM.SelectedItem.SubItems(18)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(19)) Then
                tmpACCNumber = lstPaymentMBM.SelectedItem.SubItems(19)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(20)) Then
                TxtRemark = lstPaymentMBM.SelectedItem.SubItems(20)
            End If
            If Len(lstPaymentMBM.SelectedItem.SubItems(23)) Then
                txtAmtWording = lstPaymentMBM.SelectedItem.SubItems(23)
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

Private Sub txtChequeAmt_Change()
Dim tmpAmt As String
    If IsNumeric(txtChequeAmt) Then
        tmpAmt = ""
        txttmpDollar = ""
        txttmpCents = ""
        tmpAmt = ConvertCurrencyToEnglish(txtChequeAmt.Text, "MedixPayMBM")
        If Mid(tmpAmt, 1, 3) = "And" Then
            tmpAmt = Mid(tmpAmt, 5, 500)
        End If
        txtAmtWording = tmpAmt & " Only"
    End If
End Sub

Private Sub txtChequeAmt_LostFocus()
Dim tmpAmt As String
    If IsNumeric(txtChequeAmt) Then
        tmpAmt = ""
        txttmpDollar = ""
        txttmpCents = ""
        tmpAmt = ConvertCurrencyToEnglish(txtChequeAmt.Text, "MedixPayMBM")
        If Mid(tmpAmt, 1, 3) = "And" Then
            tmpAmt = Mid(tmpAmt, 5, 500)
        End If
        txtAmtWording = tmpAmt & " Only"
    End If
End Sub

Private Sub txtChequeNo_LostFocus()
    txtChequeNo = ChkStr(txtChequeNo)
End Sub


Private Sub txtOPBatchNo1_LostFocus()
    txtOPBatchNo1 = ChkStr(txtOPBatchNo1)
End Sub

Private Sub txtOPBatchNo2_LostFocus()
    txtOPBatchNo2 = ChkStr(txtOPBatchNo2)
End Sub

Private Sub TxtPayTo_LostFocus()
    TxtPayTo = ChkStr(TxtPayTo)
End Sub


Private Sub TxtRemark_lostfocus()
   TxtRemark = ChkStr(TxtRemark)
End Sub

Private Sub TxtWSNo_LostFocus()
    TxtWSNo = ChkStr(TxtWSNo)
End Sub

'*************************** Start Diable Non Integer Value
Private Sub TxtChequeAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub lstPaymentMBM_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        'SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtDateTime, m_blnDirection(14)
    Case 15
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtDateTime, m_blnDirection(16)
    Case 17
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 19
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(19)
    Case 20
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(21)
    Case 22
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(22)
    Case 23
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(23)
    Case 24
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(24)
    Case 25
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(25)
    Case 26
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
    Case 27
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(27)
    Case 28
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(28)
    Case 29
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(29)
    Case 30
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(30)
    Case 31
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(31)
    Case 32
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(32)
    Case 33
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(33)
    Case 34
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(34)
    Case 35
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(35)
    Case 36
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(36)
    Case 37
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(37)
    
    End Select
End Sub

'***********************Start EntriesEnable***************************
Private Sub EntriesEnable(status As Boolean)

Dim ctl As Control

    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = status
        End If
    Next ctl
    
End Sub
'***********************Start EntriesEnable***************************

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

Private Sub txtChequeDate_LostFocus()
    If IsDate(txtChequeDate) Then
    Else
        If txtChequeDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtChequeDate.SetFocus
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

Private Sub txtPVDateFr_LostFocus()
    If IsDate(txtPVDateFr) Then
    Else
        If txtPVDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDateFr.SetFocus
        End If
    End If
End Sub

Private Sub txtPVDateTo_LostFocus()
    If IsDate(txtPVDateTo) Then
    Else
        If txtPVDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDateTo.SetFocus
        End If
    End If
End Sub

Private Function CaseNoCheck()
'ChkStr (lstWorksheetTrackList.listitems(listloop1).SubItems(2))
    If Trim(ChkStr(TxtWSNo)) = Trim(Mid(lstPaymentMBM.ListItems(listloop1).SubItems(4), 1, IIf(InStr(1, lstPaymentMBM.ListItems(listloop1).SubItems(4), "-") = 0, 10, InStr(1, lstPaymentMBM.ListItems(listloop1).SubItems(4), "-") - 1))) Then
        check = True
    End If
End Function

Private Sub ClearFramePV()
    TxtPayTo = ""
    txtPVDate = "__/__/____"
    txtChequeDate.Text = "__/__/____"
    txtChequeNo.Text = ""
    CboBank.Text = "MB - MAYBANK"
    TxtAccountNo = ""
    txtChequeAmt = ""
    CboACPeriodMth = ""
    CboACPeriodYr = ""
    TxtRemark = ""
    txtCaseNo = ""
    txtPayIndex = ""
    txtPVNo = ""
    txtAmtWording = ""
End Sub
Private Sub EnableFramePV(status As Boolean)
    TxtPayTo.Enabled = status
    txtPVDate.Enabled = status
    txtChequeDate.Enabled = status
    txtChequeNo.Enabled = status
    CboBank.Enabled = status
    TxtAccountNo.Enabled = status
    txtChequeAmt.Enabled = status
    CboACPeriodMth.Enabled = status
    CboACPeriodYr.Enabled = status
    TxtRemark.Enabled = status
    txtCaseNo.Enabled = status
    txtPayIndex.Enabled = status
    txtPVNo.Enabled = status
End Sub
Private Sub PreviewPV()
Const METHOD_NAME = "PreviewPV"
On Error GoTo debugs

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim strsql, MBMSql, RECSql As String
Dim PV As New ADODB.Recordset
Dim tmpCase, tmpIndex As String
Dim RecCnt As Integer
Dim MBM As New ADODB.Recordset
Dim REC As New ADODB.Recordset
Dim ReceiptNo As String
Dim totAmtPaid, totPostage, totPayeble2M As Double
    totAmtPaid = 0
    totPostage = 0
    totPayeble2M = 0
    RecCnt = 0
    Screen.MousePointer = vbHourglass
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If
    
    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    objExcel.Interactive = False
    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If

    objExcel.WindowState = xlMaximized

    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\PV-ClinicMBM.xlt")
    Set objWorksheet = objWorkbook.Sheets(1)
    objExcel.DisplayAlerts = False
    Set objWorksheet = objWorkbook.Worksheets.Item(1)
    For i = 1 To lstPaymentMBM.ListItems.Count
         If lstPaymentMBM.ListItems(i).Checked = True Then
            tmpCase = Trim(lstPaymentMBM.ListItems(i).SubItems(4))
            strsql = "select CASNumber,BordxRefNo, MBMPaidAmt, MBMNumber, CLMPayableByMedix2M, " & _
                " ClaimCloseDate, PVNo, PVChequeNo, pvpAYtO, PVDate, CLMPostageFee " & _
                "from View_Payment2MBM where CASNumber = '" & tmpCase & "'"
            PV.Open strsql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
            If PV.RecordCount Then
                RecCnt = RecCnt + 1
                
                MBMSql = "SELECT MBMAddress1,MBMAddress2,MBMAddress3, CTYCode, MBMPostCode, STACode" & _
                         " FROM MBM WHERE MBMNumber='" & PV!MBMNumber & "'"
                MBM.Open MBMSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                
                objWorksheet.Cells(8, "C") = TxtPayTo
                objWorksheet.Cells(9, "C") = MBM!MBMAddress1
                objWorksheet.Cells(10, "C") = Trim(MBM!MBMAddress2 & " " & MBM!MBMAddress3)
                objWorksheet.Cells(11, "C") = Trim(MBM!MBMPostCode & " " & Trim(MBM!CTYCode) & " " & Trim(MBM!STACode))
                objWorksheet.Cells(8, "I") = txtPVNo '& "/" & Format(PV!PVDate, "MM/YY")
                objWorksheet.Cells(9, "I") = PV!PVChequeNo
                objWorksheet.Cells(10, "I") = PV!PVDate
                objWorksheet.Cells(36, "C") = Logon
                MBM.Close
                Set MBM = Nothing
                
                ReceiptNo = ""
                RECSql = "select ReceiptNo from ReceiptfrMNRB where CASNumber = '" & tmpCase & "'"
                REC.Open RECSql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
                    If REC.RecordCount Then
                        ReceiptNo = REC!ReceiptNo
                    End If
                REC.Close
                Set REC = Nothing
                totAmtPaid = totAmtPaid + IIf(IsNumeric(PV!CLMPayableByMedix2M), PV!CLMPayableByMedix2M, 0)
                totPostage = totPostage + IIf(IsNumeric(PV!CLMPostageFee), PV!CLMPostageFee, 0)
                totPayeble2M = totPayeble2M + IIf(IsNumeric(PV!CLMPayableByMedix2M), PV!CLMPayableByMedix2M, 0)
                objWorksheet.Cells(15 + CDbl(RecCnt), "A") = i
                objWorksheet.Cells(15 + CDbl(RecCnt), "B") = PV!CasNumber
                objWorksheet.Cells(15 + CDbl(RecCnt), "C") = PV!ClaimCloseDate
                objWorksheet.Cells(15 + CDbl(RecCnt), "D") = PV!BordxRefNo
                'objWorksheet.Cells(15 + CDbl(i), "E") = ReceiptNo
                objWorksheet.Cells(15 + CDbl(RecCnt), "F") = PV!PVPayto
                objWorksheet.Cells(15 + CDbl(RecCnt), "H") = PV!CLMPayableByMedix2M
                objWorksheet.Cells(15 + CDbl(RecCnt), "I") = PV!CLMPostageFee
                objWorksheet.Cells(15 + CDbl(RecCnt), "J") = Format(PV!CLMPayableByMedix2M - PV!CLMPostageFee, "#,##0.00")
                PV.MoveNext
            End If
            PV.Close
            Set PV = Nothing
        End If
    Next i
    objWorksheet.Cells(33, "J") = CDbl(totAmtPaid) - CDbl(totPostage)
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    Screen.MousePointer = vbDefault
    Exit Sub
    
HANDLE_ERROR:
    MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
    Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    If Err.Description <> "" Then
        Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
    End If
    
ExitFunction:
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    objExcel.Quit
    Screen.MousePointer = vbDefault
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub





