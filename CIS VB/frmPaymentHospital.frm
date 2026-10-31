VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPaymentClinic 
   Caption         =   "Payment to Clinic"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   0
      TabIndex        =   36
      Top             =   360
      Width           =   11775
      Begin VB.CheckBox ChkPeriodYr 
         Height          =   255
         Left            =   11280
         TabIndex        =   46
         Top             =   1545
         Width           =   375
      End
      Begin VB.CheckBox chkInvAmt 
         Height          =   255
         Left            =   11280
         TabIndex        =   45
         Top             =   420
         Width           =   375
      End
      Begin VB.CheckBox chkPVNo 
         Height          =   255
         Left            =   11280
         TabIndex        =   44
         Top             =   690
         Width           =   375
      End
      Begin VB.CheckBox chkPVDate 
         Height          =   255
         Left            =   11280
         TabIndex        =   43
         Top             =   975
         Width           =   375
      End
      Begin VB.CheckBox chkPeriodMth 
         Height          =   255
         Left            =   11280
         TabIndex        =   42
         Top             =   1230
         Width           =   375
      End
      Begin VB.CheckBox ChkInvNo 
         Height          =   255
         Left            =   5280
         TabIndex        =   41
         Top             =   1920
         Width           =   375
      End
      Begin VB.CheckBox ChkInvDate 
         Height          =   255
         Left            =   11280
         TabIndex        =   40
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox TxtWSNoTo 
         Height          =   310
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   38
         Top             =   160
         Width           =   1560
      End
      Begin VB.TextBox TxtWSNo 
         Height          =   310
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   39
         Top             =   160
         Width           =   1560
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   59
         Top             =   440
         Width           =   3615
      End
      Begin VB.ComboBox CboCLN 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   60
         Top             =   690
         Width           =   3615
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1560
         TabIndex        =   61
         Top             =   975
         Width           =   3615
      End
      Begin VB.ComboBox CboPayment 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   1260
         Width           =   3615
      End
      Begin VB.TextBox TxtWSVersion 
         Height          =   310
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   63
         Top             =   1545
         Width           =   1560
      End
      Begin VB.TextBox TxtWSVersionTo 
         Height          =   310
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   64
         Top             =   1545
         Width           =   1560
      End
      Begin VB.TextBox TxtInvNo 
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   65
         Top             =   1830
         Width           =   3615
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   3240
         TabIndex        =   37
         Top             =   180
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtInvDate1 
         Height          =   315
         Left            =   8160
         TabIndex        =   47
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtInvAmt1 
         Height          =   285
         Left            =   8160
         MaxLength       =   8
         TabIndex        =   49
         Top             =   405
         Width           =   1215
      End
      Begin VB.TextBox txtPVNo1 
         Height          =   285
         Left            =   8160
         MaxLength       =   15
         TabIndex        =   51
         Top             =   660
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtPVDate1 
         Height          =   315
         Left            =   8160
         TabIndex        =   56
         Top             =   915
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDate2 
         Height          =   315
         Left            =   9840
         TabIndex        =   48
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtInvAmt2 
         Height          =   285
         Left            =   9840
         MaxLength       =   8
         TabIndex        =   50
         Top             =   405
         Width           =   1335
      End
      Begin VB.TextBox txtPVNo2 
         Height          =   285
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   52
         Top             =   660
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         Height          =   315
         ItemData        =   "frmPaymentHospital.frx":0000
         Left            =   8160
         List            =   "frmPaymentHospital.frx":0002
         TabIndex        =   57
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         Height          =   315
         ItemData        =   "frmPaymentHospital.frx":0004
         Left            =   8160
         List            =   "frmPaymentHospital.frx":0006
         TabIndex        =   58
         Top             =   1485
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtPVDate2 
         Height          =   315
         Left            =   9840
         TabIndex        =   53
         Top             =   915
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         Height          =   315
         ItemData        =   "frmPaymentHospital.frx":0008
         Left            =   9840
         List            =   "frmPaymentHospital.frx":000A
         TabIndex        =   54
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         Height          =   315
         ItemData        =   "frmPaymentHospital.frx":000C
         Left            =   9840
         List            =   "frmPaymentHospital.frx":000E
         TabIndex        =   55
         Top             =   1485
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Clinic"
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
         Left            =   120
         TabIndex        =   87
         Top             =   735
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Payor"
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
         Left            =   120
         TabIndex        =   86
         Top             =   1035
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Inv No"
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
         Left            =   120
         TabIndex        =   85
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Pymt Status"
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
         Left            =   120
         TabIndex        =   84
         Top             =   1305
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Inv Amt"
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
         Left            =   6720
         TabIndex        =   83
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label14 
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
         Left            =   9480
         TabIndex        =   82
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "PV No"
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
         Left            =   6720
         TabIndex        =   81
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label16 
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
         Left            =   9480
         TabIndex        =   80
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "PV Date"
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
         Left            =   6720
         TabIndex        =   79
         Top             =   1005
         Width           =   855
      End
      Begin VB.Label Label19 
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
         Left            =   9480
         TabIndex        =   78
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "Inv Date"
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
         Left            =   6720
         TabIndex        =   77
         Top             =   195
         Width           =   975
      End
      Begin VB.Label Label21 
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
         Left            =   9480
         TabIndex        =   76
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label34 
         Caption         =   "A/C Period Mth"
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
         Left            =   6720
         TabIndex        =   75
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label32 
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
         Left            =   9480
         TabIndex        =   74
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label31 
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
         Left            =   9480
         TabIndex        =   73
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label30 
         Caption         =   "A/C Period Yr"
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
         Left            =   6720
         TabIndex        =   72
         Top             =   1590
         Width           =   1815
      End
      Begin VB.Label Label12 
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
         Left            =   5280
         TabIndex        =   71
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label25 
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
         Left            =   3240
         TabIndex        =   70
         Top             =   160
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Ver No"
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
         Left            =   120
         TabIndex        =   69
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label27 
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
         Left            =   3240
         TabIndex        =   68
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Clm No"
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
         Left            =   120
         TabIndex        =   67
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label35 
         Caption         =   "Health Type"
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
         Left            =   120
         TabIndex        =   66
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Payment to Clinic"
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
         TabIndex        =   35
         Top             =   90
         Width           =   3285
      End
   End
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   0
      TabIndex        =   17
      Top             =   7490
      Width           =   11775
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   330
         Left            =   7440
         TabIndex        =   27
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   8160
         TabIndex        =   26
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10440
         TabIndex        =   25
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   11040
         TabIndex        =   24
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Height          =   330
         Left            =   6120
         TabIndex        =   23
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   6720
         TabIndex        =   22
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9480
         TabIndex        =   21
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdPrintPV 
         Caption         =   "Print P&V"
         Height          =   330
         Left            =   8760
         TabIndex        =   20
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   330
         Left            =   4800
         TabIndex        =   19
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   330
         Left            =   5520
         TabIndex        =   18
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label29 
         Caption         =   "Checked Amt"
         Height          =   255
         Left            =   2520
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   840
         TabIndex        =   32
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   210
         Width           =   645
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   1800
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   210
         Width           =   645
      End
      Begin VB.Label LblAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3600
         TabIndex        =   28
         Top             =   210
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   6480
      Width           =   11775
      Begin VB.TextBox TxtPVNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   4
         Top             =   200
         Width           =   1215
      End
      Begin VB.TextBox txtChequeNo 
         Height          =   285
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   3
         Top             =   200
         Width           =   1215
      End
      Begin VB.ComboBox CboACPeriodMth 
         Height          =   315
         ItemData        =   "frmPaymentHospital.frx":0010
         Left            =   5880
         List            =   "frmPaymentHospital.frx":0012
         TabIndex        =   2
         Top             =   200
         Width           =   1215
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
         Height          =   840
         Left            =   7920
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   200
         Width           =   3705
      End
      Begin MSMask.MaskEdBox txtPVDate 
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   450
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtChequeDate 
         Height          =   315
         Left            =   3360
         TabIndex        =   7
         Top             =   435
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtChequeAmt 
         Height          =   285
         Left            =   3360
         MaxLength       =   8
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox CboACPeriodYr 
         Height          =   315
         ItemData        =   "frmPaymentHospital.frx":0014
         Left            =   5880
         List            =   "frmPaymentHospital.frx":0016
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "PV No"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "PV Date"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Cheque No"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   225
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Cheque Amt"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "A/C Period Yr"
         Height          =   255
         Left            =   4680
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "A/C Period Mth"
         Height          =   255
         Left            =   4680
         TabIndex        =   11
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Remark"
         Height          =   255
         Left            =   7200
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "Cheque Date"
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstPaymentCLN 
      Height          =   3855
      Left            =   0
      TabIndex        =   88
      Top             =   2640
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
      NumItems        =   19
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1764
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
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Pay2HOS"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amt Paid"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "OS Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Amt Paying"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Patient Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Clm No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Clm Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Date Visit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Payment Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Entry By"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "PV HOS Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "PV No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Cheque No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Text            =   "Cheque Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Text            =   "A/C Period"
         Object.Width           =   2540
      EndProperty
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
      TabIndex        =   90
      Top             =   8040
      Width           =   1575
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
      Left            =   2520
      TabIndex        =   89
      Top             =   8040
      Width           =   2295
   End
End
Attribute VB_Name = "frmPaymentClinic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Source As Integer
Public intMBMAccount As Integer
Public intExit As Integer
Dim CaseNoArray() As String
Dim one As Boolean
Dim searchCriteria As String
Private CINIC As String
Dim ScanCode As String
Dim Check As Boolean
Dim listloop1 As Integer
Dim PVNumber As String
Private m_blnDirection(1 To 22) As Boolean

Private Sub cboHLTCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHLTCode.Text, cboHLTCode.SelStart) + Chr(KeyAscii) + Mid(cboHLTCode.Text, cboHLTCode.SelStart + cboHLTCode.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboHLTCode.ListCount - 1
       
        If NewText = LCase(Left(cboHLTCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHLTCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboHLTCode.Text = ValidValue
               cboHLTCode.SelStart = 0
               cboHLTCode.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub ChkScan_Click()
    lstPaymentCLN.ListItems.Clear
    If ChkScan.value = 1 Then
        Label25.Visible = False
        TxtWSNoTo.Visible = False
        ChkScan.Left = 3240
        TxtWSNo = ""
        TxtWSNo.SetFocus
    Else
        Label25.Visible = True
        TxtWSNoTo.Visible = True
        ChkScan.Left = 5280
        TxtWSNo.SetFocus
    End If
End Sub

Private Sub CmdAdd_Click()
    CmdSave.Enabled = True
    CmdUpdate.Enabled = False
End Sub

Private Sub cmdPrintPV_Click()
'On Error Resume Next
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim i As Integer
    Dim strsql As String
    Dim PV As New ADODB.Recordset
    
    Screen.MousePointer = vbHourglass

    If TxtPVNo = "" Then
    
        MsgBox "Please Enter or Generate PV No", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
    Else
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
        End If
        
        If Err.Number > 0 Then
            MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
            GoTo ExitFunction
        End If
        
        On Error GoTo HANDLE_ERROR
        ' Don't allow user to affect workbook
        objExcel.Interactive = False
    
        strsql = "select CASNumber,BordxRefNo,BordxPAYCode, CLMTotalApproved, " & _
                "InvNo, InvDate, InvAmount, PVCLNNo, PVCLNChequeNo, PVCLNDate, " & _
                "CLNName, CLNAdd1, CLNAdd2,CLNAdd3,CLMPayableByMedix2H " & _
                "from View_Payment2CLN where PVCLNNo = '" & TxtPVNo & "'"
        PV.Open strsql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
        'Set objExcel = New excel.Application
        
        If PV.RecordCount Then
            RecordSelected = PV.RecordCount
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
        
            objExcel.WindowState = xlMaximized
    
        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\PV-CLNpital.xlt")
        Set objWorksheet = objWorkbook.Sheets(1)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
        'Do While objWorkbook.Worksheets.Count > 1
         '   Set objWorksheet = objWorkbook.Worksheets.Item(objWorkbook.Worksheets.Count)
          '  objWorksheet.Delete
        'Loop
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
            
            objWorksheet.Cells(9, "C") = PV!CLNName
            objWorksheet.Cells(10, "C") = PV!CLNAdd1
            objWorksheet.Cells(11, "C") = PV!CLNAdd2
            objWorksheet.Cells(12, "C") = PV!CLNAdd3
            objWorksheet.Cells(9, "I") = TxtPVNo & "/" & Format(PV!PVCLNDate, "MM/YY")
            objWorksheet.Cells(10, "I") = PV!PVCLNChequeNo
            objWorksheet.Cells(11, "I") = PV!PVCLNDate
            objWorksheet.Cells(38, "C") = Logon
            For i = 1 To RecordSelected
            
                objWorksheet.Cells(16 + CDbl(i), "A") = i
                objWorksheet.Cells(16 + CDbl(i), "B") = PV!INVNo
                objWorksheet.Cells(16 + CDbl(i), "C") = PV!InvDate
                objWorksheet.Cells(16 + CDbl(i), "D") = PV!CASNumber
                objWorksheet.Cells(16 + CDbl(i), "E") = PV!BordxRefNo
                objWorksheet.Cells(16 + CDbl(i), "I") = PV!CLMPayableByMedix2H
            PV.MoveNext
            
            Next i
            
            End If
            'cmdCredit.Enabled = False
            objExcel.DisplayAlerts = True
            'objWorksheet.Columns.AutoFit
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
            GoTo ExitFunction
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

Private Sub cmdprint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelAnalysis1(lstPaymentCLN)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEdit_Click()
    one = True
    CmdSave.Enabled = False
    CmdUpdate.Enabled = True

End Sub

Public Sub cmdSearch_Click()
Dim listrecord As Integer

'    ScanCode = TxtWSNo
'    If Mid(ScanCode, 1, 2) = "LO" Then
'        TxtWSNo = ""
'        Call SearchForm(ScanCode)
'    Else
        
        Screen.MousePointer = vbHourglass
        LblAmt = 0
        one = False
        Erase CaseNoArray
        ReDim CaseNoArray(0)
'        If ChkScan.value = 1 Then
'            If TxtWSNo = "" Then
'                MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
'                TxtWSNo.SetFocus
'            Else
'                listrecord = lstPaymentCLN.ListItems.count
'                If listrecord <> 0 Then
                    
'                    For listloop1 = 1 To listrecord
'                        Call CaseNoCheck
'                    Next listloop1
'                End If
'                    If Check = True Then
'                        MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
'                    ElseIf Check = False Then
'                        Call SearchRecord
'                    End If
'                    Check = False
'                Screen.MousePointer = Default
'                TxtWSNo = ""
'                TxtWSNo.SetFocus
'            End If
'            TxtWSNo.SetFocus
'        Else
            Call SearchRecord
 '       End If
 '   End If
    Screen.MousePointer = Default
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    Check = False
    TxtWSNo.SetFocus
    CmdExit.Enabled = True
End Sub

Private Sub CmdUpdate_Click()
Dim RecordSaved
Dim PayCLN As New ADODB.Recordset
Dim UpdatePayCLN As New ADODB.Recordset
Dim strsql As String
    
    If one = False Then
    ElseIf lstPaymentCLN.ListItems.count = 0 Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
           
        ElseIf TxtPVNo = "" Then
            MsgBox "Please Confirm PV No!", vbCritical
        ElseIf IsDate(txtPVDate) = False Then
            MsgBox "Invalid PV Date!", vbCritical
            txtPVDate.SetFocus
        ElseIf txtPVDate = "__/__/____" Then
            MsgBox "Please Enter PV Date! ", vbCritical
            txtPVDate.SetFocus
        ElseIf txtChequeAmt = "" Then
            MsgBox "Please Enter Cheque Amt!", vbCritical
            txtChequeAmt.SetFocus
        Else
            strsql = "Select PVCLNIndex,* From PAyment2CLN Where PVCLNIndex = '" & lstPaymentCLN.SelectedItem.SubItems(17) & "'"
                     
            PayCLN.Open strsql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
            RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If PayCLN.RecordCount = 0 Then
                    UpdatePayCLN.AddNew
                End If
                
                With PayCLN
                    !PVCLNNo = Trim(TxtPVNo)
                    !PVCLNDate = txtPVDate
                    If txtChequeAmt <> "" Then
                        !CLNPaidAmt = Trim(txtChequeAmt)
                    Else
                        !CLNPaidAmt = ""
                    End If
                    If txtChequeNo <> "" Then
                        !PVCLNChequeNo = Trim(txtChequeNo)
                    Else
                        !PVCLNChequeNo = ""
                    End If
                    If txtChequeDate <> "__/__/____" Then
                        !PVCLNChequeDate = Format(txtChequeDate, "dd/mm/yyyy")
                    Else
                        !PVCLNChequeDate = Null
                    End If
                    !CLNRemark = Trim(TxtRemark)
                    !Pay2CLNLastUpdateDate = Date
                    !Pay2CLNLastUpdateUser = Logon
                    If CboACPeriodMth <> "" Then
                        !CLNAccPeriodMth = Trim(CboACPeriodMth)
                    Else
                         !CLNAccPeriodMth = ""
                    End If
                    If CboACPeriodYr <> "" Then
                        !CLNAccPeriodYr = Trim(CboACPeriodYr)
                    Else
                        !CLNAccPeriodYr = ""
                    End If
                    .Update
                End With
                MsgBox "Record Updated...", vbOKOnly + vbInformation, "Information"
          
                CmdSave.Enabled = True
                one = False
            End If
            PayCLN.Close
            Set PayCLN = Nothing
            
            TxtPVNo.Text = ""
            txtPVDate = "__/__/____"
            txtChequeAmt.Text = ""
            txtChequeNo.Text = ""
            txtChequeDate.Text = "__/__/____"
            TxtRemark.Text = ""
            CboACPeriodMth.Text = ""
            CboACPeriodYr.Text = ""
    End If
End Sub

Private Sub Form_Load()
    ReDim CaseNoArray(0)
    Call ComboListing
    CboPayment.ListIndex = 0
    intMBMAccount = 1
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblAmt = 0
'    Label25.Visible = False
'    TxtWSNoTo.Visible = False
    Check = False
    one = False
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    CmdSave.Enabled = True
    one = False
    ReDim CaseNoArray(0)
    lstPaymentCLN.ListItems.Clear
    CboPayment.ListIndex = 0
    lblCurrent = 0
    lblTotal = 0
    LblAmt = 0
    Check = False
    ChkScan.value = 1
    TxtWSNo.SetFocus
End Sub


Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

'************************Start ComboListing**************************

Private Sub ComboListing()
Dim CLN As New ADODB.Recordset
Dim PAY As New ADODB.Recordset
Dim MonthStart
Dim YearStart
Dim HLT As New ADODB.Recordset

    PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
    CLN.Open "Select CLNName from CLN", medixmasterconn, adOpenKeyset, adLockOptimistic

    HLT.Open "HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until HLT.EOF
        cboHLTCode.AddItem ChkStr(HLT!HLTCode) & " - " & ChkStr(HLT!HLTDescription)
        HLT.MoveNext
    Loop
    HLT.Close
    Set HLT = Nothing

    Do Until PAY.EOF
        CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
    Do Until CLN.EOF
        CboCLN.AddItem Trim(CLN!CLNName)
        CLN.MoveNext
    Loop
    CLN.Close
    Set CLN = Nothing
    
    With CboPayment
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
        .AddItem "3 - Partial Paid"
        .AddItem "4 - Paid & Partial Paid"
        .AddItem "5 - Unpaid & Partial Paid"
    End With
    
    For MonthStart = 1 To 12
        With CboACPeriodMth
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
        With CboACPeriodYr
            .AddItem YearStart
        End With
        With cboACPeriodYRFr
            .AddItem YearStart
        End With
        With cboACPeriodYRTo
            .AddItem YearStart
        End With
    Next YearStart
    
End Sub
'************************Start ComboListing**************************

Private Function SearchRecord()
Dim strsql As String
Dim StrAppend As String
Dim Sql As String
    
    StrAppend = ""
    CINIC = ""
            
    If Len(Trim(CboCLN)) Then
        CINIC = Replace(CboCLN, "'", "''")
        StrAppend = StrAppend + " AND CLNName = '" & Trim(CINIC) & "'"
    End If
    
    If Len(Trim(CboPayor)) Then
        StrAppend = StrAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
        
    If Mid(CboPayment, 1, 1) = 1 Then
        StrAppend = StrAppend + " and PVPayCLNStatus = 'F'"
    End If
 
    If Mid(CboPayment, 1, 1) = 2 Then
        StrAppend = StrAppend + " and (PVPayCLNStatus = '' or PVPayCLNStatus is null) "
    End If
 
    If Mid(CboPayment, 1, 1) = 3 Then
        StrAppend = StrAppend + " and PVPayCLNStatus = 'P'"
    End If
 
    If Mid(CboPayment, 1, 1) = 4 Then
        StrAppend = StrAppend + " and (PVPayCLNStatus = 'P' or PVPayCLNStatus = 'F') "
    End If
            
    If ChkScan.value = 1 Then
        StrAppend = StrAppend + " AND CASNumber = '" & Trim(TxtWSNo) & "'"
    Else
        If Len(TxtWSNo) Then
            StrAppend = StrAppend + " AND CASNumber >= '" & Trim(TxtWSNo) & "'"
        End If
    
        If Len(TxtWSNoTo) Then
            StrAppend = StrAppend + " AND CASNumber <= '" & Trim(TxtWSNoTo) & "'"
        End If
    End If
    
    If Len(TxtWSVersion) Then
        StrAppend = StrAppend + " and ClaimVersion >= '" & Trim(TxtWSVersion) & "'"
    End If
    
    If Len(TxtWSVersionTo) Then
        StrAppend = StrAppend + " and ClaimVersion <= '" & Trim(TxtWSVersionTo) & "'"
    End If
    
    If Len(TxtInvNo) Then
        StrAppend = StrAppend + " and InvNo = '" & Trim(TxtInvNo) & "'"
    End If
    
    If ChkInvNo.value = 1 Then
        Sql = ""
        Sql = "AND (InvNo Is null OR InvNo = '') " & ""
        StrAppend = StrAppend + Sql
    End If
    
    If (txtInvDate1) <> "__/__/____" And IsDate(txtInvDate1) = True _
        And (txtInvDate2) <> "__/__/____" And IsDate(txtInvDate2) = True Then
        Sql = ""
        Sql = " AND InvDate >= '" & Format(Trim(txtInvDate1), "mm/dd/yyyy") & _
              "' And  InvDate <= '" & Format(Trim(txtInvDate2), "mm/dd/yyyy") & "'"
        StrAppend = StrAppend + Sql
    End If
    
    If Len(Trim(txtInvDate1)) > 0 And IsDate(txtInvDate1) = True Then
        StrAppend = StrAppend + " And InvDate >= '" & Format(txtInvDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtInvDate2)) > 0 And IsDate(txtInvDate2) = True Then
        StrAppend = StrAppend + " And InvDate <= '" & Format(txtInvDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkInvDate.value = 1 Then
        Sql = ""
        Sql = "AND (PVDate Is null OR PVDate = '') " & ""
        StrAppend = StrAppend + Sql
    End If
    
    If Len(txtInvAmt1) Then
        StrAppend = StrAppend + " and INVAmount >= '" & Trim(txtInvAmt1) & "'"
    End If
        
    If Len(TxtInvAmt2) Then
        StrAppend = StrAppend + " and INVAmount <= '" & Trim(TxtInvAmt2) & "'"
    End If
    
    If chkInvAmt.value = 1 Then
        Sql = ""
        Sql = "AND (INVAmount Is null OR INVAmount = '') " & ""
        StrAppend = StrAppend + Sql
    End If
    
    If Len(txtPVNo1) Then
        StrAppend = StrAppend + " and PVCLNNo >= '" & Trim(txtPVNo1) & "'"
    End If
        
    If Len(txtPVNo2) Then
        StrAppend = StrAppend + " and PVCLNNo <= '" & Trim(txtPVNo2) & "'"
    End If
    
    If chkPVNo.value = 1 Then
        Sql = ""
        Sql = "AND (PVNo Is null OR PVNo = '') " & ""
        StrAppend = StrAppend + Sql
    End If
    
    If (txtPVDate1) <> "__/__/____" And IsDate(txtPVDate1) = True _
        And (txtPVDate2) <> "__/__/____" And IsDate(txtPVDate2) = True Then
        Sql = ""
        Sql = " AND PVCLNDate >= '" & Format(Trim(txtPVDate1), "mm/dd/yyyy") & _
              "' And PVCLNDate <= '" & Format(Trim(txtPVDate2), "mm/dd/yyyy") & "'"
        StrAppend = StrAppend + Sql
    End If
    
    If Len(Trim(txtPVDate1)) > 0 And IsDate(txtPVDate1) = True Then
        StrAppend = StrAppend + " And PVCLNDate >= '" & Format(txtPVDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPVDate2)) > 0 And IsDate(txtPVDate2) = True Then
        StrAppend = StrAppend + " And PVCLNDate <= '" & Format(txtPVDate2, "mm/dd/yyyy") & "'"
    End If
    
    If chkPVDate.value = 1 Then
        Sql = ""
        Sql = "AND (PVCLNDate Is null OR PVCLNDate = '') " & ""
        StrAppend = StrAppend + Sql
    End If
        
    If Len(cboACPeriodMthFr) Then
        StrAppend = StrAppend + " and CLNAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
        
    If Len(cboACPeriodMthTo) Then
        StrAppend = StrAppend + " and CLNAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If chkPeriodMth.value = 1 Then
        Sql = ""
        Sql = " AND CLNAccPeriodMth Is null " & ""
        StrAppend = StrAppend + Sql
    End If
    
    If Len(cboACPeriodYRTo) Then
        StrAppend = StrAppend + " and CLNAccPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
    End If
        
    If Len(cboACPeriodYRTo) Then
        StrAppend = StrAppend + " and CLNAccPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
    
    If ChkPeriodYr.value = 1 Then
        Sql = ""
        Sql = " AND CLNAccPeriodYr Is null " & ""
        StrAppend = StrAppend + Sql
    End If
    
    searchCriteria = StrAppend
    Call Payment2CLNListing(StrAppend)

End Function


Private Sub cboCLN_Change()
    If InStr(CboCLN.Text, "null") Then
       CboCLN.Enabled = False
    End If
    
End Sub

Private Sub cboCLN_KeyPress(KeyAscii As Integer)
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
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim INVAmt As Variant
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim Check As Boolean
Dim RECFull As Boolean
Dim RECPartial As String
Dim CASNumber As String
Dim WSVersion As String

   total = 0
   Current = 0
   If ChkScan.value = 0 Then
        lstPaymentCLN.ListItems.Clear
   End If

    
    Sql = " Select * From VIEW_Payment2CLN where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
    
    rst2.Open Sql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
    
    If rst2.RecordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
    Do
        Do Until rst2.EOF
        total = rst2.RecordCount
        lblTotal = total
        txtPVStatus = ""
        If Len(rst2!PVPayCLNStatus) Then
            txtPVStatus = rst2!PVPayCLNStatus
        End If
            
            With rst2
                Set Record = lstPaymentCLN.ListItems.Add(, , "")
                
                If Trim(txtPVStatus) = "F" Then
                    Record.SubItems(1) = "Full"
                ElseIf Trim(txtPVStatus) = "P" Then
                    Record.SubItems(1) = "Partial"
                Else
                    Record.SubItems(1) = ""
                End If

               If Len(rst2!CLMPayableByMedix2c) Then
                    Record.SubItems(5) = Format(rst2!CLMPayableByMedix2c, "#,##0.00")
                End If
                If Len(rst2!CLNPaidAmt) Then
                    Record.SubItems(6) = Format(rst2!CLNPaidAmt, "#,##0.00")
                End If
                
                If rst2!CoverID = "00" Then
                    If Len(rst2!MBMName) Then
                        Record.SubItems(9) = Trim(rst2!MBMName)
                    End If
                Else
                    If Len(rst2!MBMCName) Then
                        Record.SubItems(9) = Trim(rst2!MBMCName)
                    End If
                End If
                
                If Len(rst2!CASNumber) Then
                    Record.SubItems(10) = rst2!CASNumber
                End If
                CASNumber = rst2!CASNumber
               
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(11) = Trim(rst2!BordxRefNo)
                End If
                
                If Len(rst2!DateVisit) Then
                    Record.SubItems(13) = Format(rst2!DateVisit, "dd/mm/yyyy")
                End If
                
                If Len(rst2!PVCLNDate) Then
                    Record.SubItems(15) = Format(rst2!PVCLNDate, "dd/mm/yyyy")
                End If
                
                
                If Len(rst2!PVCLNIndex) Then
                    Record.SubItems(17) = rst2!PVCLNIndex
                End If
                
                If Len(rst2!PVCLNNo) Then
                    Record.SubItems(18) = Trim(rst2!PVCLNNo)
                End If
                
                If Len(rst2!PVCLNChequeNo) Then
                    Record.SubItems(19) = Trim(rst2!PVCLNChequeNo)
                End If
                
                If Len(!PVCLNChequeDate) Then
                    Record.SubItems(20) = Trim(rst2!PVCLNChequeDate)
                End If
                
                If Len(rst2!CLNAccPeriodMth & "/" & rst2!CLNAccPeriodYr) Then
                    Record.SubItems(21) = (rst2!CLNAccPeriodMth & "/" & rst2!CLNAccPeriodYr)
                End If
                
                strsqlPAY = "Select * from Payment2CLN where CASNumber = '" & Trim(CASNumber) & "' and ClaimVersion = '" & Trim(WSVersion) & "'"
                PAY.Open strsqlPAY, MediConnCISPayment, adOpenKeyset, adLockOptimistic
                INVAmt = 0
                If Trim(txtPVStatus) = "P" Then
                    Do Until PAY.EOF
                        INVAmt = INVAmt + PAY!CLNPaidAmt
                    PAY.MoveNext
                    Loop
                    If Len(INVAmt) Then
                        Record.SubItems(6) = INVAmt
                        Record.SubItems(7) = rst2!CLMPayableByMedix2H - INVAmt
                    End If
                ElseIf Trim(txtPVStatus) = "" Then
                    If Len(rst2!CLMPayableByMedix2c) Then
                        Record.SubItems(7) = rst2!CLMPayableByMedix2c
                    End If
                End If
                Current = Current + 1
                lblCurrent = Current
            End With
            PAY.Close
            Set PAY = Nothing

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
End Sub

Private Sub cmdSave_Click()
Dim PAYCode As String
Dim CASNo As String
Dim CASVer As String
Dim status As String
Dim InvoiceNo As String
Dim ii As Integer
Dim Cnt As Integer
Dim strsqlCLM As String
Dim delimiter1, delimiter2, delimiter3, delimiter4 As Integer
Cnt = 0
Dim Update
    
        If lstPaymentCLN.ListItems.count = 0 Then
            MsgBox "Please select the record", vbOKOnly + vbCritical
           
        ElseIf cboHLTCode = "" Then
            MsgBox "Please select the Health Type", vbOKOnly + vbCritical
        ElseIf IsDate(txtPVDate) = False Then
            MsgBox "Invalid PV Date!", vbCritical
            txtPVDate.SetFocus
        ElseIf txtPVDate = "__/__/____" Then
            MsgBox "Please Enter PV Date! ", vbCritical
            txtPVDate.SetFocus
        ElseIf txtChequeAmt = "" Then
            MsgBox "Please Enter Cheque Amt!", vbCritical
            txtChequeAmt.SetFocus
        ElseIf txtChequeAmt <> LblAmt And TxtRemark = "" Then
            MsgBox "Amount to paid is incorrect" & vbNewLine & "Please key in remark for reason.", vbCritical + vbOKOnly
            TxtRemark.SetFocus
                    
        Else
            Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If Update = vbYes Then
            
                If Mid(cboHLTCode, 1, 1) = "S" Then
                    PVNumber = GenerateSPVNo
                Else
                    PVNumber = GenerateNPVNo
                End If
                    TxtPVNo = ""
                    TxtPVNo = PVNumber
               
               For ii = 0 To UBound(CaseNoArray)
                 If Trim(CaseNoArray(ii)) <> "" Then
                    delimiter1 = InStr(1, CaseNoArray(ii), "-")
                    CASNo = Trim(Mid(CaseNoArray(ii), 1, IIf(delimiter1 = 0, 10, delimiter1 - 1)))
                    delimiter2 = InStr(1, CaseNoArray(ii), "~")
                    CASVer = Trim(Mid(CaseNoArray(ii), delimiter1 + 1, delimiter2 - delimiter1 - 1))
                    delimiter3 = InStr(delimiter2 + 1, CaseNoArray(ii), "~")
                    InvoiceNo = Trim(Mid(CaseNoArray(ii), delimiter2 + 1, delimiter3 - delimiter2 - 1))
                    delimiter4 = InStr(delimiter3 + 1, CaseNoArray(ii), "~")
                    status = Trim(Mid(CaseNoArray(ii), delimiter3 + 1, delimiter4 - delimiter3 - 1)) 'Len(delimiter4 - delimiter3 - 1)))
                    PaymentAmt = Trim(Mid(CaseNoArray(ii), delimiter4 + 1, Len(CaseNoArray(ii)) - delimiter4))

                    Call UpdateRecord(CASNo, CInt(CASVer), InvoiceNo, status, PaymentAmt)
                 End If
               Next ii
               Erase CaseNoArray
               Arraycnt = 0
               ReDim CaseNoArray(0)
               MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
            End If
        End If

End Sub

Private Sub UpdateRecord(CASNo As String, CASVer As Integer, InvoiceNo As String, status As String, PaymentAmt As Variant)
Dim Update
Dim strsql As String
Dim PAY As New ADODB.Recordset
Dim strsqlPAY As String
Dim CLM As New ADODB.Recordset
Dim strsqlCLM As String
    
    
    strsqlPAY = " Select PVCLNIndex, * from Payment2CLN " & _
                " Where CASNumber = '" & Trim(CASNo) & "' AND Claimversion = '" & Trim(CASVer) & "'"
    PAY.Open strsqlPAY, MediConnCISPayment, adOpenKeyset, adLockOptimistic
    
    If PAY.RecordCount > 0 And PVPayCLNStatus = "F" Then
        MsgBox "Record Duplicated!", vbCritical
    Else
        If PAY.RecordCount = 0 Then
            PAY.AddNew
        
        End If
            With PAY
                !CASNumber = Trim(ChkStr(CASNo))
                !claimversion = Trim(CASVer)
                !PVCLNNo = Trim(ChkStr(TxtPVNo))
                !PVCLNDate = Format(txtPVDate, "dd/mm/yyyy")
                If txtChequeDate <> "__/__/____" Then
                    !PVCLNChequeDate = Format(txtChequeDate, "dd/mm/yyyy")
                Else
                    !PVCLNChequeDate = ""
                End If
                If txtChequeNo <> "" Then
                    !PVCLNChequeNo = Trim(ChkStr(txtChequeNo))
                Else
                    !PVCLNChequeNo = ""
                End If
                If Len(TxtRemark) Then
                    !CLNRemark = Trim(ChkStr(TxtRemark))
                Else
                    !CLNRemark = ""
                End If
                If status = "Full" Then
                    If IsNumeric(!CLNPaidAmt) = False Then
                        !CLNPaidAmt = 0
                    End If
                    
                    If !CLNPaidAmt <> PaymentAmt Then
                        !CLNPaidAmt = PaymentAmt + !CLNPaidAmt
                    Else
                        !CLNPaidAmt = PaymentAmt
                    End If
                Else
                    If IsNumeric(!CLNPaidAmt) = False Then
                        !CLNPaidAmt = 0
                    End If
                        !CLNPaidAmt = PaymentAmt + !CLNPaidAmt
                End If
                If CboACPeriodMth <> "" Then
                    !CLNAccPeriodMth = Trim(CboACPeriodMth)
                Else
                    !CLNAccPeriodMth = ""
                End If
                If CboACPeriodYr <> "" Then
                    !CLNAccPeriodYr = Trim(CboACPeriodYr)
                Else
                    !CLNAccPeriodYr = ""
                End If
                !Pay2CLNLastUpdateDate = Date
                !Pay2CLNLastUpdateUser = Logon
                !PVPayCLNStatus = Mid(status, 1, 1)
            End With
            PAY.Update
        
        PAY.Close
        Set PAY = Nothing
    End If
    LblAmt = 0
End Sub



Private Sub lstPaymentCLN_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String

If one = True Then

CmdSave.Enabled = False

If lstPaymentCLN.ListItems.count Then
    TxtPVNo.Text = ""
    txtPVDate.mask = ""
    txtPVDate.Text = ""
    txtPVDate.mask = "##/##/####"
    txtChequeAmt.Text = ""
    txtChequeNo.Text = ""
    TxtRemark.Text = ""
    CboACPeriodMth.Text = ""
    CboACPeriodYr.Text = ""
    txtChequeDate.mask = ""
    txtChequeDate.Text = ""
    txtChequeDate.mask = "##/##/####"

    strsql = " SELECT PVCLNIndex, PVCLNNo, PVCLNDate, CLNPaidAmt, PVCLNChequeNo, " & _
    " CLNRemark, CLNAccPeriodMth, CLNAccPeriodYr,PVCLNChequeDate From Payment2CLN Where PVCLNIndex = '" & lstPaymentCLN.SelectedItem.SubItems(17) & "'"
    rst3.Open strsql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
    
    If rst3.RecordCount = 0 Then
        MsgBox "Payment hasn't been made!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        TxtPVNo = Trim(rst3!PVCLNNo)
        If Len(rst3!PVCLNDate) Then
            txtPVDate = Format(rst3!PVCLNDate, "DD/MM/YYYY")
        End If
        If Len(rst3!CLNPaidAmt) Then
            txtChequeAmt = rst3!CLNPaidAmt
        End If
        If Len(rst3!PVCLNChequeDate) Then
            txtChequeDate = Format(rst3!PVCLNChequeDate, "DD/MM/YYYY")
        End If
        If Len(rst3!PVCLNChequeNo) Then
            txtChequeNo = Trim(rst3!PVCLNChequeNo)
        End If
        If Len(rst3!CLNRemark) Then
            TxtRemark = Trim(rst3!CLNRemark)
        End If
        If Len(rst3!CLNAccPeriodMth) Then
            CboACPeriodMth = Trim(rst3!CLNAccPeriodMth)
        End If
        If Len(rst3!CLNAccPeriodYr) Then
            CboACPeriodYr = Trim(rst3!CLNAccPeriodYr)
        End If
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
    End If
End If
End If

End Sub

Private Sub lstPaymentCLN_DblClick()
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim caseno As String
Dim caseversion As String

    CmdSave.Enabled = True
    
        
    If lstPaymentCLN.SelectedItem.Checked = True Then
        MsgBox "Please Uncheck the record!", vbCritical
    Else
        caseno = Trim(Mid(lstPaymentCLN.SelectedItem.SubItems(10), 1, IIf(InStr(1, lstPaymentCLN.SelectedItem.SubItems(10), "-") = 0, 10, InStr(1, lstPaymentCLN.SelectedItem.SubItems(10), "-") - 1)))
        caseversion = Trim(Mid(lstPaymentCLN.SelectedItem.SubItems(10), InStr(1, lstPaymentCLN.SelectedItem.SubItems(10), "-") + 1, Len(lstPaymentCLN.SelectedItem.SubItems(10)) - InStr(1, lstPaymentCLN.SelectedItem.SubItems(10), "-")))
            
        strsqlCLM = "Select * From Payment2CLN where CASNumber = '" & Trim(caseno) & "' And ClaimVersion = '" & Trim(caseversion) & "'"
        CLM.Open strsqlCLM, MediConnCISPayment, adOpenKeyset, adLockOptimistic
                
        If CLM.RecordCount Then
            If Trim(CLM!PVPayCLNStatus) = "F" Then
                MsgBox "Full Payment has been made", vbCritical
                lstPaymentCLN.SelectedItem.Checked = False
            ElseIf Trim(CLM!PVPayCLNStatus) = "P" Then
                    MsgBox "Partial Payment has been made", vbCritical
                    FrmMBMAccount.txtOSAmt = lstPaymentCLN.SelectedItem.SubItems(7)
                    FrmMBMAccount.txtCLMNo = lstPaymentCLN.SelectedItem.SubItems(10)
                    FrmMBMAccount.Source = 3
                    FrmMBMAccount.Show
            End If
        Else
            FrmMBMAccount.txtOSAmt = lstPaymentCLN.SelectedItem.SubItems(7)
            FrmMBMAccount.txtCLMNo = lstPaymentCLN.SelectedItem.SubItems(10)
            FrmMBMAccount.Source = 3
            FrmMBMAccount.Show
        End If
    End If
End Sub

Private Sub lstPaymentCLN_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    CmdSave.Enabled = True
    one = False
    Call ProcessCheckedItem(Item)
End Sub

Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
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
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim caseno As String
Dim caseversion As String
Dim InvoiceNo As String
Static Arraycnt As Integer
Dim strsqlPay2CLN As String
Dim Payment2CLN As New ADODB.Recordset
    
    caseno = Trim(Mid(Item.SubItems(10), 1, IIf(InStr(1, Item.SubItems(10), "-") = 0, 10, InStr(1, Item.SubItems(10), "-") - 1)))
    caseversion = Trim(Mid(Item.SubItems(10), InStr(1, Item.SubItems(10), "-") + 1, Len(Item.SubItems(10)) - InStr(1, Item.SubItems(10), "-")))
    InvoiceNo = Trim(Item.SubItems(3))
    status = Trim(Mid(Item.SubItems(1), InStr(1, Item.SubItems(1), "-") + 1, Len(Item.SubItems(10)) - InStr(1, Item.SubItems(10), "-")))

    strsqlCLM = " Select * From ClaimInvoice " & _
                " where CASNumber = '" & Trim(caseno) & "' " & _
                " And INVWSVersion = '" & Trim(caseversion) & "'"
    CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    strsqlPAY = "Select * From Payment2CLN where CASNumber = '" & Trim(caseno) & "' And ClaimVersion = '" & Trim(caseversion) & "'"
    PAY.Open strsqlPAY, MediConnCISPayment, adOpenKeyset, adLockPessimistic


    If Item.Checked = True Then
        If PAY.RecordCount = 0 Then 'And Trim(PAY!PVPayCLNStatus) = "F" Then
            If Partial = False Then
                If Item.SubItems(6) = "" Then
                    Item.SubItems(8) = CDbl(Item.SubItems(5))
                Else
                    Item.SubItems(8) = CDbl(Item.SubItems(7))
                End If
                    Item.SubItems(1) = "Full"
            End If
                LblAmt = LblAmt + CDbl(Format(Item.SubItems(8), "#,##0.00"))
                ReDim Preserve CaseNoArray(Arraycnt)
                CaseNoArray(Arraycnt) = Item.SubItems(10) & "~" & Item.SubItems(3) & "~" & Item.SubItems(1) & "~" & Item.SubItems(8)
                Arraycnt = Arraycnt + 1
        Else
            If Trim(PAY!PVPayCLNStatus) = "F" Then
                MsgBox "Payment has been made!", vbCritical
                Item.Checked = False
            Else
                If Partial = False Then
                    If Item.SubItems(6) = "" Then
                        Item.SubItems(8) = CDbl(Item.SubItems(5))
                    Else
                        Item.SubItems(8) = CDbl(Item.SubItems(7))
                    End If
                End If
                Item.SubItems(1) = "Full"
                LblAmt = LblAmt + CDbl(Format(Item.SubItems(8), "#,##0.00"))
                ReDim Preserve CaseNoArray(Arraycnt)
                CaseNoArray(Arraycnt) = Item.SubItems(10) & "~" & Item.SubItems(3) & "~" & Item.SubItems(1) & "~" & Item.SubItems(8)
                Arraycnt = Arraycnt + 1
            End If
        End If
    Else
        newbound = 0

         If UBound(CaseNoArray) <> 0 Then
            For i = 0 To UBound(CaseNoArray)
                If (Item.SubItems(10)) <> CaseNoArray(i) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = CaseNoArray(i)
                    newbound = newbound + 1
                End If
            Next
            If Item.SubItems(1) = "Full" Then
                If Item.SubItems(6) = "" Then
                    LblAmt = LblAmt - CDbl(Format(Item.SubItems(5), "#,##0.00"))
                Else
                    LblAmt = LblAmt - CDbl(Format(Item.SubItems(8), "#,##0.00"))
                End If
            Else
                If Len(Item.SubItems(8)) Then
                    LblAmt = LblAmt - CDbl(Format(Item.SubItems(8), "#,##0.00"))
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
        
        If PAY.RecordCount Then
            If PAY!PVPayCLNStatus <> "" Then
                If Trim(PAY!PVPayCLNStatus) = "P" Then
                    Item.SubItems(1) = "Partial"
                End If
            Else
                Item.SubItems(1) = ""
            End If
            
        Else
            Item.SubItems(1) = ""
        End If
                Item.SubItems(8) = ""
    Else
        If Item.SubItems(1) = "Full" Then
            If Item.SubItems(6) = "" Then
                LblAmt = LblAmt - CDbl(Format(Item.SubItems(5), "#,##0.00"))
            Else
                LblAmt = LblAmt - CDbl(Format(Item.SubItems(8), "#,##0.00"))
            End If
        ElseIf Item.SubItems(1) = "Partial" Then
            If Len(Item.SubItems(8)) Then
                If CDbl(Format(Item.SubItems(8), "#,##0.00")) > LblAmt Then
                    LblAmt = 0
                Else
                    LblAmt = LblAmt - CDbl(Format(Item.SubItems(8), "#,##0.00"))
                End If
            End If
        End If
        
        Erase CaseNoArray
        ReDim CaseNoArray(0)
        Arraycnt = 0
  
        
        If PAY.RecordCount Then
            If Len((PAY!PVPayCLNStatus)) Then
                Item.SubItems(1) = "Partial"
                Item.SubItems(8) = ""
            End If
        Else
            Item.SubItems(1) = ""
            Item.SubItems(8) = ""
        End If
                   
    End If
    End If
    CLM.Close
    Set CLM = Nothing
    PAY.Close
    Set PAY = Nothing
End Function

'*************************** Start Disable Non Integer Value
Private Sub TxtChequeAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
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

Private Sub txtInvAmt1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtInvAmt2_KeyPress(KeyAscii As Integer)
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

Private Sub txtPVDate1_LostFocus()
    If IsDate(txtPVDate1) Then
    Else
        If txtPVDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtPVDate2_LostFocus()
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
    Else
        If txtPVDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDate.SetFocus
        End If
    End If
End Sub

Private Sub lstPaymentCLN_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        'SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
    Case 6
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtDateTime, m_blnDirection(14)
    Case 15
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtDateTime, m_blnDirection(15)
    Case 16
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtDateTime, m_blnDirection(16)
    Case 17
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 19
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(19)
    Case 20
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtDateTime, m_blnDirection(21)
    Case 22
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(22)
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

Private Sub TxtWSNo_GotFocus()
    cmdSearch.Default = True
End Sub

Private Sub TxtWSNo_LostFocus()
    TxtWSNo = ChkStr(TxtWSNo)
    cmdSearch.Default = False
End Sub

Private Sub TxtWSNoTo_LostFocus()
TxtWSNoTo = ChkStr(TxtWSNoTo)
End Sub

Private Function CaseNoCheck()
    If Trim(ChkStr(TxtWSNo)) = Trim(Mid(lstPaymentCLN.ListItems(listloop1).SubItems(10), 1, IIf(InStr(1, lstPaymentCLN.ListItems(listloop1).SubItems(10), "-") = 0, 10, InStr(1, lstPaymentCLN.ListItems(listloop1).SubItems(10), "-") - 1))) Then
        Check = True
    End If
End Function
