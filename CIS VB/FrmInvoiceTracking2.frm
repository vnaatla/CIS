VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInvoiceTracking2 
   Caption         =   "Invoice Tracking - Jupiter"
   ClientHeight    =   8592
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8592
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   104
      Top             =   7080
      Width           =   11775
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   7680
         TabIndex        =   58
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10440
         TabIndex        =   62
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   11040
         TabIndex        =   63
         Top             =   200
         Width           =   615
      End
      Begin VB.TextBox txtNoofDay 
         Height          =   285
         Left            =   120
         TabIndex        =   105
         Text            =   "Text1"
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9480
         TabIndex        =   61
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   330
         Left            =   6960
         TabIndex        =   57
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   330
         Left            =   8280
         TabIndex        =   59
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   330
         Left            =   8880
         TabIndex        =   60
         Top             =   200
         Width           =   615
      End
      Begin VB.Label Label29 
         Caption         =   "Total App Amt"
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
         Left            =   2280
         TabIndex        =   118
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label LblAppAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Left            =   3360
         TabIndex        =   117
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1320
         TabIndex        =   111
         Top             =   195
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1440
         TabIndex        =   110
         Top             =   600
         Width           =   975
      End
      Begin VB.Label LblTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   5760
         TabIndex        =   109
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "Total Act Amt"
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   108
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   1200
         TabIndex        =   107
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   195
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3435
      Left            =   0
      TabIndex        =   64
      Top             =   240
      Width           =   11775
      Begin VB.CheckBox ChkPayor2 
         Height          =   255
         Left            =   5760
         TabIndex        =   115
         Top             =   1120
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom 
         Height          =   255
         Left            =   5475
         TabIndex        =   11
         Top             =   1380
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom2 
         Height          =   255
         Left            =   5760
         TabIndex        =   12
         Top             =   1380
         Width           =   255
      End
      Begin VB.CheckBox ChkClinic2 
         Height          =   255
         Left            =   5760
         TabIndex        =   7
         Top             =   850
         Width           =   255
      End
      Begin VB.TextBox TxtWSNo2 
         Height          =   310
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   1
         Top             =   255
         Width           =   1575
      End
      Begin VB.TextBox TxtWSNo1 
         Height          =   310
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   0
         Top             =   255
         Width           =   1605
      End
      Begin VB.CheckBox ChkICNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   11280
         TabIndex        =   32
         Top             =   520
         Width           =   375
      End
      Begin VB.TextBox TxtPlan2 
         Height          =   315
         Left            =   9720
         MaxLength       =   4
         TabIndex        =   28
         Top             =   255
         Width           =   1455
      End
      Begin VB.TextBox TxtPlan1 
         Height          =   315
         Left            =   7560
         MaxLength       =   4
         TabIndex        =   27
         Top             =   255
         Width           =   1455
      End
      Begin VB.CheckBox ChkPlan 
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
         Height          =   280
         Left            =   11280
         TabIndex        =   29
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox ChkName 
         Height          =   255
         Left            =   5480
         TabIndex        =   14
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox ChkBordxNo 
         Height          =   255
         Left            =   5480
         TabIndex        =   23
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox ChkCaseNo 
         Height          =   255
         Left            =   5480
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox ChkInvNo 
         Height          =   255
         Left            =   5480
         TabIndex        =   20
         Top             =   2235
         Width           =   255
      End
      Begin VB.CheckBox ChkMbmNo 
         Height          =   255
         Left            =   5480
         TabIndex        =   17
         Top             =   1970
         Width           =   255
      End
      Begin VB.CheckBox ChkPayor 
         Height          =   255
         Left            =   5480
         TabIndex        =   9
         Top             =   1120
         Width           =   255
      End
      Begin VB.CheckBox ChkClinic 
         Height          =   255
         Left            =   5480
         TabIndex        =   6
         Top             =   850
         Width           =   255
      End
      Begin VB.CheckBox ChkPayTo 
         Height          =   255
         Left            =   5480
         TabIndex        =   4
         Top             =   580
         Width           =   255
      End
      Begin VB.CheckBox ChkInvAmt 
         Height          =   255
         Left            =   11280
         TabIndex        =   47
         Top             =   1850
         Width           =   375
      End
      Begin VB.CheckBox chkDOV 
         Height          =   255
         Left            =   11280
         TabIndex        =   35
         Top             =   800
         Width           =   375
      End
      Begin VB.CheckBox chkDOR 
         Height          =   255
         Left            =   11280
         TabIndex        =   38
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox chkInvDate 
         Height          =   255
         Left            =   11280
         TabIndex        =   41
         Top             =   1360
         Width           =   375
      End
      Begin VB.CheckBox chkInvReg 
         Height          =   255
         Left            =   11280
         TabIndex        =   44
         Top             =   1600
         Width           =   375
      End
      Begin VB.CheckBox ChkMonth 
         Height          =   255
         Left            =   11280
         TabIndex        =   50
         Top             =   2100
         Width           =   375
      End
      Begin VB.CheckBox ChkYear 
         Height          =   255
         Left            =   11280
         TabIndex        =   53
         Top             =   2400
         Width           =   375
      End
      Begin VB.CheckBox ChkOPBatchNo 
         Height          =   255
         Left            =   11280
         TabIndex        =   56
         Top             =   2680
         Width           =   375
      End
      Begin VB.ComboBox CboPayTo 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Text            =   "C - Clinic"
         Top             =   540
         Width           =   4220
      End
      Begin VB.ComboBox CboClinic 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   830
         Width           =   4220
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   1080
         Width           =   4220
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1365
         Width           =   4220
      End
      Begin VB.TextBox TxtPatientName 
         Height          =   310
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1635
         Width           =   4220
      End
      Begin VB.TextBox TxtMBMNo1 
         Height          =   310
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   15
         Top             =   1920
         Width           =   1570
      End
      Begin VB.TextBox TxtInvNo1 
         Height          =   310
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   18
         Top             =   2205
         Width           =   1570
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   310
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   21
         Top             =   2490
         Width           =   1570
      End
      Begin VB.ComboBox CboCaseStatus 
         Height          =   315
         Left            =   1200
         TabIndex        =   24
         Top             =   2775
         Width           =   1570
      End
      Begin VB.ComboBox CboClinicType 
         Height          =   315
         Left            =   1200
         TabIndex        =   26
         Top             =   3060
         Width           =   1570
      End
      Begin VB.TextBox TxtMBMNo2 
         Height          =   310
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   16
         Top             =   1920
         Width           =   1570
      End
      Begin VB.TextBox TxtInvNo2 
         Height          =   310
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   19
         Top             =   2205
         Width           =   1570
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   310
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   22
         Top             =   2490
         Width           =   1570
      End
      Begin VB.ComboBox CboClaimability 
         Height          =   315
         Left            =   3840
         TabIndex        =   25
         Top             =   2775
         Width           =   1570
      End
      Begin VB.TextBox TxtICNo1 
         Height          =   300
         Left            =   7560
         MaxLength       =   20
         TabIndex        =   30
         Top             =   540
         Width           =   1455
      End
      Begin VB.TextBox TxtICNo2 
         Height          =   300
         Left            =   9720
         MaxLength       =   20
         TabIndex        =   31
         Top             =   540
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtDOVDate1 
         Height          =   320
         Left            =   7560
         TabIndex        =   33
         Top             =   795
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   572
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDORDate1 
         Height          =   315
         Left            =   7560
         TabIndex        =   36
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDate1 
         Height          =   315
         Left            =   7560
         TabIndex        =   39
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtIRDDate1 
         Height          =   315
         Left            =   7560
         TabIndex        =   42
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtINVAmtFr 
         Height          =   285
         Left            =   7560
         TabIndex        =   45
         Top             =   1850
         Width           =   1455
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking2.frx":0000
         Left            =   7560
         List            =   "FrmInvoiceTracking2.frx":0002
         TabIndex        =   48
         Top             =   2100
         Width           =   1455
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking2.frx":0004
         Left            =   7560
         List            =   "FrmInvoiceTracking2.frx":0006
         TabIndex        =   51
         Top             =   2390
         Width           =   1455
      End
      Begin VB.TextBox TxtOPBatchNo1 
         Height          =   285
         Left            =   7560
         MaxLength       =   8
         TabIndex        =   54
         Top             =   2670
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtDOVDate2 
         Height          =   320
         Left            =   9720
         TabIndex        =   34
         Top             =   795
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   572
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDORDate2 
         Height          =   315
         Left            =   9720
         TabIndex        =   37
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDate2 
         Height          =   315
         Left            =   9720
         TabIndex        =   40
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtIRDDate2 
         Height          =   315
         Left            =   9720
         TabIndex        =   43
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtINVAmtTo 
         Height          =   285
         Left            =   9720
         TabIndex        =   46
         Top             =   1850
         Width           =   1455
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking2.frx":0008
         Left            =   9720
         List            =   "FrmInvoiceTracking2.frx":000A
         TabIndex        =   49
         Top             =   2100
         Width           =   1455
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking2.frx":000C
         Left            =   9720
         List            =   "FrmInvoiceTracking2.frx":000E
         TabIndex        =   52
         Top             =   2390
         Width           =   1455
      End
      Begin VB.TextBox TxtOPBatchNo2 
         Height          =   285
         Left            =   9720
         MaxLength       =   8
         TabIndex        =   55
         Top             =   2670
         Width           =   1455
      End
      Begin VB.ComboBox cboInfoFFS 
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
         Left            =   3840
         Sorted          =   -1  'True
         TabIndex        =   121
         Top             =   3060
         Width           =   1570
      End
      Begin VB.Label Label31 
         Caption         =   "FFS Status"
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
         Left            =   2880
         TabIndex        =   120
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label49 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   1490
         Width           =   975
      End
      Begin VB.Label Label33 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   135
         Left            =   5760
         TabIndex        =   102
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label52 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   135
         Left            =   5520
         TabIndex        =   101
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label85 
         Caption         =   "I/C No"
         Height          =   255
         Left            =   6360
         TabIndex        =   100
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label84 
         Caption         =   "To"
         Height          =   255
         Left            =   9240
         TabIndex        =   99
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label47 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   98
         Top             =   2550
         Width           =   255
      End
      Begin VB.Label Label46 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   97
         Top             =   1995
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   96
         Top             =   2235
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Case Status"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "Inv Reg Date"
         Height          =   255
         Left            =   6360
         TabIndex        =   94
         Top             =   1600
         Width           =   1140
      End
      Begin VB.Label Label23 
         Caption         =   "To"
         Height          =   255
         Left            =   9240
         TabIndex        =   93
         Top             =   1600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   1190
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Clinic"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "DOV"
         Height          =   255
         Left            =   6360
         TabIndex        =   90
         ToolTipText     =   "Date of Visit"
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "DOR"
         Height          =   255
         Left            =   6360
         TabIndex        =   89
         ToolTipText     =   "Date of Register"
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   9240
         TabIndex        =   88
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   9240
         TabIndex        =   87
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   6360
         TabIndex        =   86
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "To"
         Height          =   255
         Left            =   9240
         TabIndex        =   85
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "Inv Date"
         Height          =   255
         Left            =   6360
         TabIndex        =   84
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Pay To"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "Claim No"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   280
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "Inv No"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   2550
         Width           =   1215
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   79
         Top             =   285
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   1770
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Member No"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "To"
         Height          =   255
         Left            =   9240
         TabIndex        =   76
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   2880
         TabIndex        =   75
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label40 
         Caption         =   "Inv Amount"
         Height          =   255
         Left            =   6360
         TabIndex        =   74
         Top             =   1850
         Width           =   855
      End
      Begin VB.Label Label41 
         Caption         =   "To"
         Height          =   255
         Left            =   9240
         TabIndex        =   73
         Top             =   1850
         Width           =   375
      End
      Begin VB.Label Label42 
         Caption         =   "A/C Period Yr"
         Height          =   255
         Left            =   6360
         TabIndex        =   72
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label43 
         Caption         =   "To"
         Height          =   255
         Left            =   9240
         TabIndex        =   71
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label44 
         Caption         =   "To"
         Height          =   255
         Left            =   9240
         TabIndex        =   70
         Top             =   2120
         Width           =   375
      End
      Begin VB.Label Label45 
         Caption         =   "A/C Period Mth"
         Height          =   255
         Left            =   6360
         TabIndex        =   69
         Top             =   2120
         Width           =   1215
      End
      Begin VB.Label Label18 
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
         Left            =   11280
         TabIndex        =   68
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "To"
         Height          =   255
         Left            =   9240
         TabIndex        =   67
         Top             =   2680
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "OP Batch No"
         Height          =   255
         Left            =   6360
         TabIndex        =   66
         Top             =   2680
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Clinic Type"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   3150
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstInvoiceTrackList 
      Height          =   3345
      Left            =   0
      TabIndex        =   116
      Top             =   3720
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   5906
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   39
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Worksheet No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Bordx Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Inv. No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Inv. Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Inv. Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Excess Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Net Amt (Inv. Amt - Excess Amt)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "CoyPayment"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Approved Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Non GP Fee"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "FFS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "DOR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "DOV"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "No of Days (DOR - DOV)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Payee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Clinic Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Text            =   "Clinic Type"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Membership No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   20
         Text            =   "Plan Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   21
         Text            =   "Insured Type"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Group Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   23
         Text            =   "Inv. Reg"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Inv. Entry By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "OP Batch No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   26
         Text            =   "Case Status"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   27
         Text            =   "Claimability"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   28
         Text            =   "Bordx Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "Reason for Treatment"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "A/C Period"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "MC Days"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "Reason for Treatment 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "Reason for Treatment 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "Reason for Treatment 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "Reason for Treatment 4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "Reason for Treatment 5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Text            =   "OP Enter By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Text            =   "FFS"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label txtFFSStatus 
      Caption         =   "FFS Status"
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
      Left            =   2880
      TabIndex        =   119
      Top             =   3405
      Width           =   975
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Invoice  Tracking - Jupiter"
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
      TabIndex        =   114
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label Label34 
      Caption         =   "* - Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   113
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label48 
      Caption         =   "* - Empty Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   112
      Top             =   7680
      Width           =   1335
   End
End
Attribute VB_Name = "FrmInvoiceTracking2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Private m_blnDirection(1 To 39) As Boolean

Private Sub ChkPayor2_Click()
Const METHOD_NAME = "ChkPayor2_Click"
On Error GoTo debugs

Dim PAY As New ADODB.Recordset
Dim strsql As String

    'PAY.Open "Select PAYCode, PAYCompanyName from PAY", MediNeptuneConnCISDB, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
    CboPayor.Clear
    If ChkPayor2.Value = 1 Then
       ' MediNeptuneConnCISWS.Open "File Name=" & BOSPath & "Conn\MediNeptuneConnCISWS.UDL"
        strsql = "EXEC SearchPayor"
        PAY.Open strsql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
        
        Do Until PAY.EOF
            CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
            PAY.MoveNext
        Loop
        PAY.Close
        Set PAY = Nothing
        'MediNeptuneConnCISWS.Close
        'Set MediNeptuneConnCISWS = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
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
    Call ActionLogUpdate(12, "FrmInvoiceTracking2")
    intExit = 0
    lblCurrent = 0
    LblTotalAmt = 0
    Call ComboListing
End Sub

'************************Start Command Button**************************
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    lstInvoiceTrackList.ListItems.Clear
    lblCurrent = 0
    LblTotalAmt = 0
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdSearch.Enabled = True
    CmdClear.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
End Sub

Public Sub cmdSearch_Click()
On Error GoTo errRun

    CmdSearch.Enabled = False
    CmdClear.Enabled = False
    CmdPrintFile.Enabled = False
    CmdExit.Enabled = False
    
    Screen.MousePointer = vbHourglass
        'MediNeptuneConnCISWS.Open "File Name=" & BOSPath & "Conn\MediNeptuneConnCISWS.UDL"
        Call SearchRecord
       ' MediNeptuneConnCISWS.Close
       ' Set MediNeptuneConnCISWS = Nothing
    Screen.MousePointer = Default
    
    Exit Sub
    
errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lstInvoiceTrackList.ListItems.Clear
    lblCurrent = ""
    LblTotalAmt = ""
    
    CmdSearch.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdPrintFile.Enabled = True
    MediNeptuneConnCISWS.Close
    Set MediNeptuneConnCISWS = Nothing
    
End Sub

Private Sub cmdBack_Click()
    lstInvoiceTrackList.Height = 3585
    lstInvoiceTrackList.Left = 0
    lstInvoiceTrackList.Top = 3480
    lstInvoiceTrackList.Width = 11775
    Frame1.Visible = True
End Sub

Private Sub CmdPrintFile_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelInvoice(lstInvoiceTrackList)
    Screen.MousePointer = Default
End Sub

Private Sub CmdView_Click()
    lstInvoiceTrackList.Height = 6705
    lstInvoiceTrackList.Left = 0
    lstInvoiceTrackList.Top = 360
    lstInvoiceTrackList.Width = 11775
    Frame1.Visible = False
End Sub
'************************End Command Button**************************


'************************Start sorted lstInvoiceTrackList**************************

Private Sub lstInvoiceTrackList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
    Case 6
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
    Case 11
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    Case 13
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 19
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(19)
    Case 20
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(21)
    Case 22
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(22)
    Case 23
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(23)
    Case 24
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(24)
    Case 25
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(25)
    Case 26
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(26)
    Case 27
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(27)
    Case 28
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
    
    Case 29
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(29)
    Case 30
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(30)
    Case 31
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(31)
    Case 32
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(32)
    Case 33
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(33)
    Case 34
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(34)
    Case 35
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(35)
    Case 36
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(36)
    Case 37
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(37)
    Case 38
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(38)
    Case 39
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(39)
    
    End Select

End Sub
'************************End sorted lstInvoiceTrackList**************************

'************************Start ComboListing**************************
Private Sub ComboListing()
Dim sql As String
Dim yearstart
Dim monthstart

    With CboPayTo
        .AddItem "C - Clinic"
        .AddItem "M - Member"
    End With
    
    With CboClaimability
        .AddItem "A - Accepted"
        .AddItem "R - Rejected"
    End With
    
    With CboCaseStatus
        .AddItem "O - Open"
        .AddItem "C - Close"
        .AddItem "D - Delete"
    End With
    
    For monthstart = 1 To 12
        With cboACPeriodMthFr
            .AddItem monthstart
        End With
        With cboACPeriodMthTo
            .AddItem monthstart
        End With
    Next monthstart
    
    For yearstart = 1999 To 3000
        With cboACPeriodYRFr
            .AddItem yearstart
        End With
        With cboACPeriodYRTo
            .AddItem yearstart
        End With
    Next yearstart
    
    CboClinicType.AddItem "0 - Specialised"
    CboClinicType.AddItem "1 - General Practisional"
    CboClinicType.AddItem "2 - Others"
        
    cboInfoFFS.AddItem "Y - Yes"
    cboInfoFFS.AddItem "N - No"

End Sub
'************************Start ComboListing**************************
Private Sub cboInfoFFS_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboInfoFFS.Text, cboInfoFFS.SelStart) + Chr(KeyAscii) + Mid(cboInfoFFS.Text, cboInfoFFS.SelStart + cboInfoFFS.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboInfoFFS.ListCount - 1
 
        If NewText = LCase(Left(cboInfoFFS.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboInfoFFS.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboInfoFFS.Text = ValidValue
               cboInfoFFS.SelStart = 0
               cboInfoFFS.SelLength = Len(ValidValue)
             End If
        End If
End Sub
'************************Start Search Record**************************

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim strsqlHOS As String
Dim CLN As New ADODB.Recordset
Dim sql As String
Dim sql2 As String
Dim Bordx As Integer
Dim CheckDate As New ADODB.Recordset
    
    
    strAppend = ""
    Clinic = ""
    GRPCompany = ""
    
    If Len(Trim(TxtWSNo1)) Then
        strAppend = strAppend + " AND CASNumber >= '" & Trim(TxtWSNo1) & "'"
    End If
    
    If Len(Trim(TxtWSNo2)) Then
        strAppend = strAppend + " AND CASNumber <= '" & Trim(TxtWSNo2) & "'"
    End If
    
    If ChkCaseNo.Value = 1 Then
        sql = ""
        sql = "AND (CASNumber Is null OR CASNumber = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(CboPayTo)) Then
        strAppend = strAppend + " AND INVPayTo = '" & Trim(Mid(CboPayTo, 1, IIf(InStr(1, CboPayTo, "-") = 0, 4, InStr(1, CboPayTo, "-") - 1))) & "'"
    End If
    
    If ChkPayTo.Value = 1 Then
        sql = ""
        sql = "AND (INVPayTo Is null OR INVPayTo = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(CboClinic)) Then
        
        strsqlHOS = "select CLNCode, CLNName from CLN where CLNCode = '" & Mid(CboClinic, 1, 4) & "'"
        CLN.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        strAppend = strAppend + " AND INVPayee = '" & Trim(CLN!CLNCode) & "'"
        
        CLN.Close
        Set CLN = Nothing
    End If
    
    If ChkClinic.Value = 1 Then
        sql = ""
        sql = "AND (INVPayee Is null OR INVPayee = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If ChkPayor.Value = 1 Then
        sql = ""
        sql = "AND (PAYCode Is null OR PAYCode = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(CboGrpCom)) Then
       strAppend = strAppend + " AND MBMGroupCompany = '" & Trim(RemStr(CboGrpCom)) & "'"
    End If
    
    If ChkGrpCom.Value = 1 Then
       strAppend = strAppend + " AND (MBMGroupCompany IS NULL OR MBMGroupCompany = '')"
    End If
    
    If Len(Trim(TxtPatientName)) Then
       strAppend = strAppend + " And PatientName = '" & Trim(TxtPatientName) & "'"
    End If
    
    If ChkName.Value = 1 Then
        sql = ""
        sql = "AND (PatientName Is null OR PatientName = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtMBMNo1)) Then
        strAppend = strAppend + " AND MBMNumber >= '" & Trim(TxtMBMNo1) & "'"
    End If
    
    If Len(Trim(TxtMBMNo2)) Then
        strAppend = strAppend + " AND MBMNumber <= '" & Trim(TxtMBMNo2) & "'"
    End If
    
    If ChkMbmNo.Value = 1 Then
        sql = ""
        sql = "AND (MBMNumber Is null OR MBMNumber = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtInvNo1)) Then
        strAppend = strAppend + " AND BillNo >= '" & Trim(TxtInvNo1) & "'"
    End If

    If Len(Trim(TxtInvNo2)) Then
        strAppend = strAppend + " AND BillNo <= '" & Trim(TxtInvNo2) & "'"
    End If

    If ChkInvNo.Value = 1 Then
        sql = ""
        sql = "AND (BillNo Is null OR BillNo = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtBordx1)) Then
        strAppend = strAppend + " AND BordxRefNo >= '" & Trim(TxtBordx1) & "'"
    End If
    
    If Len(Trim(TxtBordx2)) Then
        strAppend = strAppend + " AND BordxRefNo <= '" & Trim(TxtBordx2) & "'"
    End If
           
    If ChkBordxNo.Value = 1 Then
        sql = ""
        sql = "AND (BordxRefNo Is null OR BordxRefNo = '')" & ""
        strAppend = strAppend + sql
    End If
        
    If Mid(CboCaseStatus, 1, 1) = "C" Then
        strAppend = strAppend + " AND CASStatus = 'C'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "O" Then
        strAppend = strAppend + " AND CASStatus = 'O'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "D" Then
        strAppend = strAppend + " AND CASStatus = 'D'"
    End If
    
    If Mid(CboClaimability, 1, 1) = "R" Then
        strAppend = strAppend + " Or CASClaimability = 'R'"
    ElseIf Mid(CboClaimability, 1, 1) = "A" Then
        strAppend = strAppend + " Or CASClaimability = 'A'"
    End If
    
    If Len(Trim(TxtPlan1)) Then
        strAppend = strAppend + " AND PLNCode >= '" & Trim(TxtPlan1) & "'"
    End If
    
    If Len(Trim(TxtPlan2)) Then
        strAppend = strAppend + " AND PLNCode <= '" & Trim(TxtPlan2) & "'"
    End If
    
    If ChkPlan.Value = 1 Then
        sql = ""
        sql = "AND (PLNCode Is null OR PLNCode = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtICNo1)) Then
        strAppend = strAppend + " AND ICNo >= '" & Trim(TxtICNo1) & "'"
    End If
    
    If Len(Trim(TxtICNo2)) Then
        strAppend = strAppend + " AND ICNo <= '" & Trim(TxtICNo2) & "'"
    End If
    
    If ChkICNo.Value = 1 Then
      strAppend = strAppend + " AND (ICNo Is null)"
    End If
    
    If Len(Trim(txtDOVDate1)) > 0 And IsDate(txtDOVDate1) = True Then
        strAppend = strAppend + " And DateVisit >= '" & Format(txtDOVDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOVDate2)) > 0 And IsDate(txtDOVDate2) = True Then
        strAppend = strAppend + " And DateVisit <= '" & Format(txtDOVDate2, "mm/dd/yyyy") & "'"
    End If
    
    If chkDOV.Value = 1 Then
        sql = ""
        sql = "AND (DateVisit Is null OR DateVisit = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDORDate1)) > 0 And IsDate(txtDORDate1) = True Then
        strAppend = strAppend + " And RegisterDate >= '" & Format(txtDORDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDORDate2)) > 0 And IsDate(txtDORDate2) = True Then
        strAppend = strAppend + " And RegisterDate <= '" & Format(txtDORDate2, "mm/dd/yyyy") & "'"
    End If
    
    If chkDOR.Value = 1 Then
        sql = ""
        sql = "AND (RegisterDate Is null OR RegisterDate = '')" & ""
        strAppend = strAppend + sql
    End If
        
    If Len(Trim(txtInvDate1)) > 0 And IsDate(txtInvDate1) = True Then
        strAppend = strAppend + " And BillDate >= '" & Format(txtInvDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtInvDate2)) > 0 And IsDate(txtInvDate2) = True Then
        strAppend = strAppend + " And BillDate <= '" & Format(txtInvDate2, "mm/dd/yyyy") & "'"
    End If
    
    If chkInvDate.Value = 1 Then
        sql = ""
        sql = "AND (BillDate Is null OR BillDate = '')" & ""
        strAppend = strAppend + sql
    End If
        
    If Len(Trim(txtIRDDate1)) > 0 And IsDate(txtIRDDate1) = True Then
        strAppend = strAppend + " And INVEnteredDate >= '" & Format(txtIRDDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtIRDDate2)) > 0 And IsDate(txtIRDDate2) = True Then
        strAppend = strAppend + " And INVEnteredDate <= '" & Format(txtIRDDate2, "mm/dd/yyyy") & "'"
    End If
    
    If chkInvReg.Value = 1 Then
        sql = ""
        sql = "AND (INVEnteredDate Is null OR INVEnteredDate = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtINVAmtFr)) Then
        strAppend = strAppend + " AND INVAmount >= " & Trim(txtINVAmtFr) & ""
    End If
    
    If Len(Trim(txtINVAmtTo)) Then
        strAppend = strAppend + " AND INVAmount <= " & Trim(txtINVAmtTo) & ""
    End If
    
    If ChkInvAmt.Value = 1 Then
        sql = ""
        sql = "AND (INVAmount Is null)" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(cboACPeriodMthFr)) Then
        strAppend = strAppend + " AND AccPeriodMonth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
    
    If Len(Trim(cboACPeriodMthTo)) Then
        strAppend = strAppend + " AND AccPeriodMonth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If ChkMonth.Value = 1 Then
        sql = ""
        sql = "AND AccPeriodMonth Is null " & ""
        strAppend = strAppend + sql
    End If

    If Len(Trim(cboACPeriodYRFr)) Then
        strAppend = strAppend + " AND AccPeriodYear >= '" & Trim(cboACPeriodYRFr) & "'"
    End If

    If Len(Trim(cboACPeriodYRTo)) Then
        strAppend = strAppend + " AND AccPeriodYear <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
    
    If ChkYear.Value = 1 Then
        sql = ""
        sql = "AND AccPeriodYear Is null " & ""
        strAppend = strAppend + sql
    End If
    
    If Mid(CboClinicType, 1, 1) = 0 Then
        strAppend = strAppend & " AND CLNType = 'S'"
    ElseIf Mid(CboClinicType, 1, 1) = 1 Then
        strAppend = strAppend & " AND CLNType = 'G'"
    ElseIf Mid(CboClinicType, 1, 1) = 2 Then
        strAppend = strAppend & " AND (CLNType = '' OR CLNType IS NULL)"
    End If
    
    If Len(Trim(TxtOPBatchNo1)) Then
        strAppend = strAppend + " AND OPBatchNo >= '" & Trim(TxtOPBatchNo1) & "'"
    End If

    If Len(Trim(TxtOPBatchNo2)) Then
        strAppend = strAppend + " AND OPBatchNo <= '" & Trim(TxtOPBatchNo2) & "'"
    End If

    If ChkOPBatchNo.Value = 1 Then
        sql = ""
        sql = "AND (OPBatchNo Is null OR OPBatchNo = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Mid(cboInfoFFS, 1, 1) = "Y" Then
        strAppend = strAppend & " AND CLNFFSStatus = 'Y'"
    ElseIf Mid(cboInfoFFS, 1, 1) = "N" Then
        strAppend = strAppend & " AND CLNFFSStatus = 'N' OR (CLNFFSStatus = '' OR CLNFFSStatus IS NULL)"
    End If
    
    Call InvoiceTrackingListing(strAppend)

End Function
'************************End Search Record**************************

'************************Start List Record**************************

Private Sub InvoiceTrackingListing(Optional strAppend As String)
Const METHOD_NAME = "InvoiceTrackingListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim Total As String
Dim Current As String
Dim TotAmt As Variant
Dim TotalInvNo As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
    
    Current = 0
    TotAmt = 0
    LblTotalAmt = 0
    LblAppAmt = 0
    lstInvoiceTrackList.ListItems.Clear
    
    strAppendCase = "1"
            
    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "ClaimBordxModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    
    'Sql = " Select CASNumber, BordxRefNo, BordxDate, BillNo, BillDate, " & _
          " INVAmount, INVPaidByMem, NetAmt, CLMFinalApprove, RegisterDate, " & _
          " DateVisit, NoofDays, INVPayee, CLNName, MBMNumber, PatientName, " & _
          " PLNCode, INSCode, MBMGroupCompany, INVEnteredDate, " & _
          " INVEnteredBy, CASStatus, ClaimBordxStatus, ReasonDesc, " & _
          " AccPeriodMonth, AccPeriodYear, INVPayTo, ICNo, Claimability, " & _
          " OPBatchNo, CLNType, CLMNonGpFee From VIEW_Invoice_Tracking where 0 = 0 "
          
    
    'If Len(Trim(strAppend)) Then
        'Sql = Sql + strAppend
    'End If
    
    'rst2.Open Sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    
         
    Do
        Do Until rst2.EOF
            
            'lblTotal = rst2.RecordCount
            
            With rst2
                
                Set Record = lstInvoiceTrackList.ListItems.Add(, , rst2!CasNumber)
                
                If Len(rst2!Bordxrefno) Then
                    Record.SubItems(1) = Trim(rst2!Bordxrefno)
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(2) = Trim(rst2!BordxDate)
                End If
                If Len(rst2!BillNo) Then
                    Record.SubItems(3) = Trim(rst2!BillNo)
                End If
                If Len(rst2!BillDate) Then
                   Record.SubItems(4) = Format(rst2!BillDate, "dd/mm/yyyy")
                End If
                If Len(rst2!INVAmount) Then
                    Record.SubItems(5) = Format(rst2!INVAmount, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                If Len(rst2!INVPaidByMem) Then
                    Record.SubItems(6) = Format(rst2!INVPaidByMem, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                
                Record.SubItems(7) = Format(rst2!NetAmt, "#,##0.00")
                
                If Len(rst2!CLMCoPayCLNColl) Then
                    Record.SubItems(8) = Format(rst2!CLMCoPayCLNColl, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(8) = "0.00"
                End If

                If Len(rst2!CLMFinalApprove) Then
                    Record.SubItems(9) = Format(rst2!CLMFinalApprove, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(9) = "0.00"
                End If
                
                If Len(rst2!ClmNonGPFee) Then
                    Record.SubItems(10) = Format(rst2!ClmNonGPFee, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(10) = "0.00"
                End If
                
                If Len(rst2!CLMMCOProcFee) Then
                    Record.SubItems(11) = Format(rst2!CLMMCOProcFee, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(11) = "0.00"
                End If
                
                If Len(rst2!RegisterDate) Then
                    Record.SubItems(12) = Format(rst2!RegisterDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!DateVisit) Then
                    Record.SubItems(13) = Format(rst2!DateVisit, "dd/mm/yyyy")
                End If
                
                If rst2!NoofDays = "0" Then
                    Record.SubItems(14) = "1"
                ElseIf Len(rst2!NoofDays) Then
                    Record.SubItems(14) = rst2!NoofDays
                End If
                
                If Len(rst2!INVPayee) Then
                    Record.SubItems(15) = rst2!INVPayee
                End If
                                                                       
                If Trim(rst2!INVPayTo) = "C" Then
                    If Len(rst2!CLNName) Then
                        Record.SubItems(16) = Trim(rst2!CLNName)
                    End If
                Else
                
                    If Len(rst2!NPClinic) Then
                        Record.SubItems(16) = Trim(rst2!NPClinic)
                    End If
                End If
                
                If Len(rst2!CLNType) Then
                    Record.SubItems(17) = Trim(rst2!CLNType)
                End If
                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(18) = Trim(rst2!MBMNumber)
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(19) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!PLNCode) Then
                    Record.SubItems(20) = Trim(rst2!PLNCode)
                End If
                
                If Len(rst2!INSCode) Then
                    Record.SubItems(21) = Trim(rst2!INSCode)
                End If
                
                If Len(rst2!MBMGroupCompany) Then
                    Record.SubItems(22) = Trim(rst2!MBMGroupCompany)
                End If
                
                If Len(rst2!INVEnteredDate) Then
                    Record.SubItems(23) = Format(rst2!INVEnteredDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!INVEnteredBy) Then
                    Record.SubItems(24) = Trim(rst2!INVEnteredBy)
                End If
                
                If Len(rst2!OPBatchNo) Then
                    Record.SubItems(25) = Trim(rst2!OPBatchNo)
                End If
                
                If Len(rst2!CASStatus) Then
                    Record.SubItems(26) = Trim(rst2!CASStatus)
                End If
                
                If Len(rst2!Claimability) Then
                    Record.SubItems(27) = Trim(rst2!Claimability)
                End If
                
                If Len(rst2!ClaimBordxStatus) Then
                    Record.SubItems(28) = Trim(rst2!ClaimBordxStatus)
                End If
                
                'If Len(rst2!ReasonDesc) Then
                 '   Record.SubItems(29) = Trim(rst2!ReasonDesc)
                'End If
                
                If Len(rst2!AccPeriodMonth & "/" & rst2!AccPeriodYear) Then
                    Record.SubItems(30) = rst2!AccPeriodMonth & "/" & rst2!AccPeriodYear
                End If
                
                If Len(rst2!MCDays) Then
                    Record.SubItems(31) = Trim(rst2!MCDays)
                End If
                
                If Len(rst2!ReasonDesc1) Then
                    Record.SubItems(32) = Trim(rst2!ReasonDesc1)
                End If
                
                If Len(rst2!ReasonDesc2) Then
                    Record.SubItems(33) = Trim(rst2!ReasonDesc2)
                End If
                
                If Len(rst2!ReasonDesc3) Then
                    Record.SubItems(34) = Trim(rst2!ReasonDesc3)
                End If
                
                If Len(rst2!ReasonDesc4) Then
                    Record.SubItems(35) = Trim(rst2!ReasonDesc4)
                End If
                
                If Len(rst2!CASTreatReason5) Then
                    If Trim(rst2!CASTreatReason5) <> "199" And Trim(rst2!CASTreatReason5) <> "299" Then
                        Record.SubItems(36) = Trim(rst2!ReasonDesc5)
                    Else
                        Record.SubItems(36) = Trim(rst2!CASDiagnosis5)
                    End If
                End If
                
                Record.SubItems(37) = IIf(Len(rst2!OPNoEnterUser), rst2!OPNoEnterUser, "")
                
                Record.SubItems(38) = IIf(Len(rst2!CLNFFSStatus), rst2!CLNFFSStatus, "")
                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!INVAmount) Then
                    TotAmt = TotAmt + rst2!INVAmount
                End If
                LblTotalAmt = Format(TotAmt, "#,##0.00")
                
                If IsNumeric(rst2!CLMFinalApprove) Then
                    LblAppAmt = CDbl(LblAppAmt) + CDbl(rst2!CLMFinalApprove)
                    LblAppAmt = Format(LblAppAmt, "#,##0.00")
                End If
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If lstInvoiceTrackList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdPrintFile.Enabled = True
        CmdExit.Enabled = True
        CmdSearch.Enabled = True
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If


End Sub
'************************End List Record**************************

'************************Start ComboBox Change/KeyPress**************************

Private Sub CboPayTo_Click()
    If InStr(CboPayTo.Text, "null") Then
       CboPayTo.Enabled = False
    End If
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

If CboPayor.Text = "" Then
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
End If
        
End Sub

Private Sub CboClaimability_Change()
    If InStr(CboClaimability.Text, "null") Then
       CboClaimability.Enabled = False
    End If
End Sub

Private Sub CboClaimability_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClaimability.Text, CboClaimability.SelStart) + Chr(KeyAscii) + Mid(CboClaimability.Text, CboClaimability.SelStart + CboClaimability.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClaimability.ListCount - 1
 
        If NewText = LCase(Left(CboClaimability.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClaimability.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClaimability.Text = ValidValue
                CboClaimability.SelStart = 0
                CboClaimability.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboClinic_Change()
    If InStr(CboClinic.Text, "null") Then
       CboClinic.Enabled = False
    End If
End Sub

Private Sub CboClinic_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClinic.Text, CboClinic.SelStart) + Chr(KeyAscii) + Mid(CboClinic.Text, CboClinic.SelStart + CboClinic.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClinic.ListCount - 1
 
        If NewText = LCase(Left(CboClinic.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClinic.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClinic.Text = ValidValue
                CboClinic.SelStart = 0
                CboClinic.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboCaseStatus_Change()
    If InStr(CboCaseStatus.Text, "null") Then
       CboCaseStatus.Enabled = False
    End If
End Sub

Private Sub CboCaseStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCaseStatus.Text, CboCaseStatus.SelStart) + Chr(KeyAscii) + Mid(CboCaseStatus.Text, CboCaseStatus.SelStart + CboCaseStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCaseStatus.ListCount - 1
 
        If NewText = LCase(Left(CboCaseStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCaseStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCaseStatus.Text = ValidValue
                CboCaseStatus.SelStart = 0
                CboCaseStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboClinicType_Change()
    If InStr(CboClinicType.Text, "null") Then
       CboClinicType.Enabled = False
    End If
End Sub

Private Sub CboClinicType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClinicType.Text, CboClinicType.SelStart) + Chr(KeyAscii) + Mid(CboClinicType.Text, CboClinicType.SelStart + CboClinicType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClinicType.ListCount - 1
 
        If NewText = LCase(Left(CboClinicType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClinicType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClinicType.Text = ValidValue
                CboClinicType.SelStart = 0
                CboClinicType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboPayTo_Change()
    If InStr(CboPayTo.Text, "null") Then
       CboPayTo.Enabled = False
    End If
End Sub

Private Sub CboPayTo_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayTo.Text, CboPayTo.SelStart) + Chr(KeyAscii) + Mid(CboPayTo.Text, CboPayTo.SelStart + CboPayTo.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayTo.ListCount - 1
 
        If NewText = LCase(Left(CboPayTo.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayTo.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPayTo.Text = ValidValue
                CboPayTo.SelStart = 0
                CboPayTo.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboACPeriodMthFr_Change()
If cboACPeriodMthFr.Text = "" Then
    If InStr(cboACPeriodMthFr.Text, "null") Then
       cboACPeriodMthFr.Enabled = False
    End If
End If
End Sub

Private Sub cboACPeriodMthFr_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

If cboACPeriodMthFr.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboACPeriodMthFr.Text, cboACPeriodMthFr.SelStart) + Chr(KeyAscii) + Mid(cboACPeriodMthFr.Text, cboACPeriodMthFr.SelStart + cboACPeriodMthFr.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboACPeriodMthFr.ListCount - 1
 
        If NewText = LCase(Left(cboACPeriodMthFr.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboACPeriodMthFr.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboACPeriodMthFr.Text = ValidValue
                cboACPeriodMthFr.SelStart = 0
                cboACPeriodMthFr.SelLength = Len(ValidValue)
            End If
        End If
End If
End Sub

Private Sub cboACPeriodMthTo_Change()
If cboACPeriodMthTo.Text = "" Then
    If InStr(cboACPeriodMthTo.Text, "null") Then
       cboACPeriodMthTo.Enabled = False
    End If
End If
End Sub


Private Sub cboACPeriodMthTo_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

If cboACPeriodMthTo.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboACPeriodMthTo.Text, cboACPeriodMthTo.SelStart) + Chr(KeyAscii) + Mid(cboACPeriodMthTo.Text, cboACPeriodMthTo.SelStart + cboACPeriodMthTo.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboACPeriodMthTo.ListCount - 1
 
        If NewText = LCase(Left(cboACPeriodMthTo.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboACPeriodMthTo.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboACPeriodMthTo.Text = ValidValue
                cboACPeriodMthTo.SelStart = 0
                cboACPeriodMthTo.SelLength = Len(ValidValue)
            End If
        End If
End If
End Sub

Private Sub cboACPeriodYRFr_Change()
If cboACPeriodYRFr.Text = "" Then
    If InStr(cboACPeriodYRFr.Text, "null") Then
       cboACPeriodYRFr.Enabled = False
    End If
End If
End Sub

Private Sub cboACPeriodYRFr_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

If cboACPeriodYRFr.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboACPeriodYRFr.Text, cboACPeriodYRFr.SelStart) + Chr(KeyAscii) + Mid(cboACPeriodYRFr.Text, cboACPeriodYRFr.SelStart + cboACPeriodYRFr.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboACPeriodYRFr.ListCount - 1
 
        If NewText = LCase(Left(cboACPeriodYRFr.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboACPeriodYRFr.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboACPeriodYRFr.Text = ValidValue
                cboACPeriodYRFr.SelStart = 0
                cboACPeriodYRFr.SelLength = Len(ValidValue)
            End If
        End If
End If
End Sub

Private Sub cboACPeriodYRTo_Change()
If cboACPeriodYRTo.Text = "" Then
    If InStr(cboACPeriodYRTo.Text, "null") Then
       cboACPeriodYRTo.Enabled = False
    End If
End If
End Sub

Private Sub cboACPeriodYRTo_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

If cboACPeriodYRTo.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboACPeriodYRTo.Text, cboACPeriodYRTo.SelStart) + Chr(KeyAscii) + Mid(cboACPeriodYRTo.Text, cboACPeriodYRTo.SelStart + cboACPeriodYRTo.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboACPeriodYRTo.ListCount - 1
 
        If NewText = LCase(Left(cboACPeriodYRTo.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboACPeriodYRTo.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboACPeriodYRTo.Text = ValidValue
                cboACPeriodYRTo.SelStart = 0
                cboACPeriodYRTo.SelLength = Len(ValidValue)
            End If
        End If
End If
End Sub
'************************End ComboBox Change/KeyPress**************************

Private Sub txtDOVDate1_LostFocus()
    If IsDate(txtDOVDate1) Then
    Else
        If txtDOVDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOVDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDOVDate2_LostFocus()
    If IsDate(txtDOVDate2) Then
    Else
        If txtDOVDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOVDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtDORDate1_LostFocus()
    If IsDate(txtDORDate1) Then
    Else
        If txtDORDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDORDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDORDate2_LostFocus()
    If IsDate(txtDORDate2) Then
    Else
        If txtDORDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDORDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtINVAmtFr_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtINVAmtTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
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

Private Sub txtIRDDate1_LostFocus()
    If IsDate(txtIRDDate1) Then
    Else
        If txtIRDDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtIRDDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtIRDDate2_LostFocus()
    If IsDate(txtIRDDate2) Then
    Else
        If txtIRDDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtIRDDate2.SetFocus
        End If
    End If
End Sub

Private Sub ChkClinic2_Click()
Dim CLN As New ADODB.Recordset
Dim strsql As String

    CboClinic.Clear
    If ChkClinic2.Value = 1 Then
    
       ' MediNeptuneConnCISWS.Open "File Name=" & BOSPath & "Conn\MediNeptuneConnCISWS.UDL"
        
        strsql = "EXEC SearchCLN"
        CLN.Open strsql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    
        Do Until CLN.EOF
            CboClinic.AddItem CLN!CLNCode & " - " & CLN!CLNName
            CLN.MoveNext
        Loop
        
        CLN.Close
        Set CLN = Nothing
       ' MediNeptuneConnCISWS.Close
       ' Set MediNeptuneConnCISWS = Nothing
    End If

End Sub

Private Sub ChkGrpCom2_Click()
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset

CboGrpCom.Clear
If ChkGrpCom2.Value = 1 Then
    
    'MediNeptuneConnCISWS.Open "File Name=" & BOSPath & "Conn\MediNeptuneConnCISWS.UDL"
    strsql = "EXEC SearchPayGrp '" & Mid(CboPayor, 1, 2) & "'"
    PAYGRP.Open strsql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    
    If Not PAYGRP.EOF Then
        Do Until PAYGRP.EOF
           CboGrpCom.AddItem Trim(PAYGRP!GrpCoName)
           PAYGRP.MoveNext
        Loop
    End If
    PAYGRP.Close
    Set PAYGRP = Nothing
   ' MediNeptuneConnCISWS.Close
   ' Set MediNeptuneConnCISWS = Nothing
End If
End Sub

Private Sub CboGrpCom_Change()
    If InStr(CboGrpCom.Text, "null") Then
       CboGrpCom.Enabled = False
    End If
End Sub

Private Sub CboGrpCom_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboGrpCom.Text, CboGrpCom.SelStart) + Chr(KeyAscii) + Mid(CboGrpCom.Text, CboGrpCom.SelStart + CboGrpCom.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboGrpCom.ListCount - 1
 
        If NewText = LCase(Left(CboGrpCom.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboGrpCom.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboGrpCom.Text = ValidValue
                CboGrpCom.SelStart = 0
                CboGrpCom.SelLength = Len(ValidValue)
            End If
        End If
End Sub
