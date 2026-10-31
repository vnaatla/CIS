VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPaymentCLNByInvoice 
   Caption         =   "Payment to Clinic By Invoice"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FramePV 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   0
      TabIndex        =   42
      Top             =   6480
      Width           =   11775
      Begin VB.ComboBox CboBank 
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
         ItemData        =   "frmPaymentCLNByBatch.frx":0000
         Left            =   5280
         List            =   "frmPaymentCLNByBatch.frx":0002
         TabIndex        =   25
         Top             =   160
         Width           =   2535
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
         Left            =   8520
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   200
         Width           =   3105
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
         Left            =   2880
         MaxLength       =   20
         TabIndex        =   22
         Top             =   200
         Width           =   1215
      End
      Begin VB.TextBox TxtPVNo 
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
         Left            =   720
         MaxLength       =   15
         TabIndex        =   19
         Top             =   200
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtPVDate 
         Height          =   315
         Left            =   720
         TabIndex        =   20
         Top             =   450
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSMask.MaskEdBox txtChequeDate 
         Height          =   315
         Left            =   2880
         TabIndex        =   23
         Top             =   435
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.TextBox txtPayIndex 
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
         Left            =   720
         MaxLength       =   15
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtChequeAmt 
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
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtAcNo 
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
         Left            =   5280
         MaxLength       =   30
         TabIndex        =   26
         Top             =   440
         Width           =   2535
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
         ItemData        =   "frmPaymentCLNByBatch.frx":0004
         Left            =   5280
         List            =   "frmPaymentCLNByBatch.frx":0006
         TabIndex        =   27
         Top             =   690
         Width           =   855
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
         ItemData        =   "frmPaymentCLNByBatch.frx":0008
         Left            =   6840
         List            =   "frmPaymentCLNByBatch.frx":000A
         TabIndex        =   28
         Top             =   690
         Width           =   975
      End
      Begin VB.TextBox txtCaseNo 
         Height          =   285
         Left            =   7320
         TabIndex        =   43
         Top             =   680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label22 
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
         Left            =   2040
         TabIndex        =   53
         Top             =   480
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
         Left            =   7920
         TabIndex        =   52
         Top             =   240
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
         Left            =   4320
         TabIndex        =   51
         Top             =   735
         Width           =   615
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
         Left            =   6240
         TabIndex        =   50
         Top             =   765
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
         Left            =   2040
         TabIndex        =   49
         Top             =   720
         Width           =   975
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
         Left            =   2040
         TabIndex        =   48
         Top             =   225
         Width           =   975
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
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   480
         Width           =   615
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
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   4320
         TabIndex        =   45
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label27 
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
         Left            =   4320
         TabIndex        =   44
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   0
      TabIndex        =   87
      Top             =   7490
      Width           =   11775
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
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
         Left            =   5520
         TabIndex        =   32
         Top             =   180
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
         Left            =   5040
         TabIndex        =   31
         Top             =   180
         Width           =   495
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
         Left            =   8760
         TabIndex        =   38
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print to &File"
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
         Left            =   9480
         TabIndex        =   39
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
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
         Left            =   6720
         TabIndex        =   34
         Top             =   180
         Width           =   735
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
         Left            =   6120
         TabIndex        =   33
         Top             =   180
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
         Left            =   11040
         TabIndex        =   41
         Top             =   180
         Width           =   615
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
         Left            =   10440
         TabIndex        =   40
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdStop 
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
         Left            =   8160
         TabIndex        =   37
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
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
         Left            =   7440
         TabIndex        =   35
         Top             =   180
         Width           =   735
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
         Left            =   4200
         TabIndex        =   30
         Top             =   180
         Width           =   855
      End
      Begin VB.Label LblAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Left            =   3120
         TabIndex        =   93
         Top             =   210
         Width           =   975
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Left            =   960
         TabIndex        =   92
         Top             =   210
         Width           =   645
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
         TabIndex        =   91
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         TabIndex        =   90
         Top             =   210
         Width           =   525
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
         Left            =   720
         TabIndex        =   89
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label29 
         Caption         =   "Amt"
         Height          =   255
         Left            =   2640
         TabIndex        =   88
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   0
      TabIndex        =   55
      Top             =   360
      Width           =   11775
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   3240
         TabIndex        =   65
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox ChkInvDate 
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
         Left            =   11280
         TabIndex        =   64
         Top             =   270
         Width           =   375
      End
      Begin VB.CheckBox ChkInvNo 
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
         Left            =   5280
         TabIndex        =   63
         Top             =   570
         Width           =   375
      End
      Begin VB.CheckBox chkPeriodMth 
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
         Left            =   11280
         TabIndex        =   62
         Top             =   1350
         Width           =   375
      End
      Begin VB.CheckBox chkPVDate 
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
         Left            =   11280
         TabIndex        =   61
         Top             =   1095
         Width           =   375
      End
      Begin VB.CheckBox chkPVNo 
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
         Left            =   11280
         TabIndex        =   60
         Top             =   810
         Width           =   375
      End
      Begin VB.CheckBox chkInvAmt 
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
         Left            =   11280
         TabIndex        =   59
         Top             =   540
         Width           =   375
      End
      Begin VB.CheckBox ChkPeriodYr 
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
         Left            =   11280
         TabIndex        =   58
         Top             =   1665
         Width           =   375
      End
      Begin VB.CheckBox ChkClinic 
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
         Left            =   5280
         TabIndex        =   57
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpComp 
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
         Left            =   5280
         TabIndex        =   56
         Top             =   1320
         Width           =   255
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
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   2
         Top             =   255
         Visible         =   0   'False
         Width           =   1580
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   1
         Top             =   255
         Visible         =   0   'False
         Width           =   1560
      End
      Begin MSMask.MaskEdBox txtInvDate2 
         Height          =   315
         Left            =   9840
         TabIndex        =   8
         Top             =   240
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
      Begin MSMask.MaskEdBox txtInvDate1 
         Height          =   315
         Left            =   7920
         TabIndex        =   7
         Top             =   240
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
      Begin VB.TextBox txtInvAmt1 
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
         Left            =   7920
         MaxLength       =   8
         TabIndex        =   9
         Top             =   525
         Width           =   1335
      End
      Begin VB.TextBox txtPVNo1 
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
         Left            =   7920
         MaxLength       =   15
         TabIndex        =   11
         Top             =   780
         Width           =   1335
      End
      Begin VB.TextBox TxtInvAmt2 
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
         Left            =   9840
         MaxLength       =   8
         TabIndex        =   10
         Top             =   525
         Width           =   1335
      End
      Begin VB.TextBox txtPVNo2 
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
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   12
         Top             =   780
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtPVDate1 
         Height          =   315
         Left            =   7920
         TabIndex        =   13
         Top             =   1035
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
         ItemData        =   "frmPaymentCLNByBatch.frx":000C
         Left            =   7920
         List            =   "frmPaymentCLNByBatch.frx":000E
         TabIndex        =   15
         Top             =   1320
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
         ItemData        =   "frmPaymentCLNByBatch.frx":0010
         Left            =   7920
         List            =   "frmPaymentCLNByBatch.frx":0012
         TabIndex        =   17
         Top             =   1605
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtPVDate2 
         Height          =   315
         Left            =   9840
         TabIndex        =   14
         Top             =   1035
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
         ItemData        =   "frmPaymentCLNByBatch.frx":0014
         Left            =   9840
         List            =   "frmPaymentCLNByBatch.frx":0016
         TabIndex        =   16
         Top             =   1320
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
         ItemData        =   "frmPaymentCLNByBatch.frx":0018
         Left            =   9840
         List            =   "frmPaymentCLNByBatch.frx":001A
         TabIndex        =   36
         Top             =   1605
         Width           =   1335
      End
      Begin VB.TextBox TxtInvNo 
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
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   0
         Top             =   540
         Width           =   3615
      End
      Begin VB.TextBox txtOPBatchNo 
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
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   3
         Top             =   780
         Width           =   3615
      End
      Begin VB.ComboBox CboCLN 
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1020
         Width           =   3615
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
         Left            =   1560
         TabIndex        =   5
         Top             =   1305
         Width           =   3615
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1590
         Width           =   3615
      End
      Begin VB.Label Label4 
         Caption         =   "Claim Number"
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
         TabIndex        =   86
         Top             =   285
         Visible         =   0   'False
         Width           =   1215
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
         TabIndex        =   85
         Top             =   285
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   9885
         TabIndex        =   84
         Top             =   945
         Width           =   255
      End
      Begin VB.Label Label30 
         Caption         =   "A/C Period Year"
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
         Left            =   6360
         TabIndex        =   83
         Top             =   1665
         Width           =   1455
      End
      Begin VB.Label Label31 
         Caption         =   "To"
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
         Left            =   9450
         TabIndex        =   82
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label32 
         Caption         =   "To"
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
         Left            =   9450
         TabIndex        =   81
         Top             =   1380
         Width           =   375
      End
      Begin VB.Label Label34 
         Caption         =   "A/C Period Month"
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
         Left            =   6360
         TabIndex        =   80
         Top             =   1380
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "To"
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
         Left            =   9450
         TabIndex        =   79
         Top             =   260
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Invoice Date"
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
         Left            =   6360
         TabIndex        =   78
         Top             =   315
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "To"
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
         Left            =   9450
         TabIndex        =   77
         Top             =   1100
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "PV Date"
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
         Left            =   6360
         TabIndex        =   76
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Left            =   9450
         TabIndex        =   75
         Top             =   820
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "PV Number"
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
         Left            =   6360
         TabIndex        =   74
         Top             =   855
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "To"
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
         Left            =   9450
         TabIndex        =   73
         Top             =   560
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Invoice Amount"
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
         Left            =   6360
         TabIndex        =   72
         Top             =   585
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Payment Status"
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
         TabIndex        =   71
         Top             =   1635
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Invoice No"
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
         TabIndex        =   70
         Top             =   550
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Group Company"
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
         TabIndex        =   69
         Top             =   1365
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Clinic"
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
         TabIndex        =   68
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5320
         TabIndex        =   67
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label35 
         Caption         =   "OP Batch No"
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
         TabIndex        =   66
         Top             =   820
         Width           =   1215
      End
   End
   Begin VB.TextBox txtPVStatus 
      Height          =   285
      Left            =   4800
      TabIndex        =   54
      Top             =   8400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView lstPaymentCLN 
      Height          =   3975
      Left            =   0
      TabIndex        =   18
      Top             =   2520
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7011
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
      NumItems        =   27
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
         Object.Width           =   0
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
         SubItemIndex    =   6
         Text            =   "Bill No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Pay2CLN"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Amt Paid"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "OS Amt"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Amt Paying"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Patient Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Clm No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Bordx No"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "Date Visit"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "Pay Date"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Entry By"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "PV CLN Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "PV No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Cheque No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   20
         Text            =   "Cheque Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Chq Amt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Text            =   "A/C Period"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Bank"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "AccountNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Remarks"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "OP Batch"
         Object.Width           =   1764
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
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   2520
      TabIndex        =   97
      Top             =   8040
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
      TabIndex        =   96
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label Label37 
      Caption         =   "* Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   95
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "   Payment to Clinic - Invoice"
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
      Left            =   0
      TabIndex        =   94
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmPaymentCLNByInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer
Public intMBMAccount As Integer
Public intExit As Integer
Private m_blnDirection(1 To 22) As Boolean

Dim EditFlag As Boolean
Dim ScanCode As String
Dim Check As Boolean
Dim listloop1 As Integer
Dim RecordSelected As Integer

Dim PVNumber As String

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
        ClnRec.Open "Select CLNCode, CLNName from CLN WHERE CLNPanel = 'Y'", medixmasterconn, adOpenKeyset, adLockOptimistic
        Do Until ClnRec.EOF
            If Len(ClnRec!CLNname) Then
                CboCLN.AddItem ClnRec!CLNCode & " - " & Trim(ClnRec!CLNname)
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

Private Sub ChkScan_Click()
    lstPaymentCLN.ListItems.Clear
    If ChkScan.Value = 1 Then
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


Private Sub cmdAdd_Click()
Dim ZZ As Integer
Dim ItemChk As Boolean
    
    ItemChk = False
    For ZZ = 1 To lstPaymentCLN.ListItems.count
        If lstPaymentCLN.ListItems(ZZ).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next ZZ

    CmdUpdate.Enabled = False
    
    If ItemChk Then
        Call ClearFramePV
        CmdSave.Enabled = True
        FramePV.Enabled = True
        CboBank.Enabled = True
        txtPVDate.SetFocus
    Else
        MsgBox "Please checked at least 1 Claim  ! ", vbCritical + vbOKOnly, DialogBoxTitle
    End If
End Sub

Private Sub CmdCheck_Click()
    If lstPaymentCLN.ListItems.count <> 0 Then
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
Dim R As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstPaymentCLN.ListItems.count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        R = SendMessageAny(lstPaymentCLN.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub

Public Sub SetCheck(hwnd As Long, lItemIndex As Long, bState As Boolean)
Dim LV As LVITEM
Dim R As Long
    With LV
    .mask = LVIF_STATE
    .state = IIf(bState, &H2000, &H1000)
    .stateMask = LVIS_STATEIMAGEMASK
    End With
    R = SendMessageAny(hwnd, LVM_SETITEMSTATE, lItemIndex, LV)
End Sub

Private Sub cmdPrintPV_Click()
'On Error Resume Next
    Dim REceipt As New ADODB.Recordset
    Dim ReceiptNo As String
    Dim objExcel As excel.Application
    Dim objWorkbook As excel.Workbook
    Dim objWorksheet As excel.Worksheet
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
    
        strsql = "select CASNumber,BordxRefNo, BillNo, BillDate, PVCLNNo,PVCLNChequeNo, PVCLNDate, " & _
                "CLNName, CLNAdd1, CLNAdd2,CLNAdd3, CLNPostCode, CLNCity, CLNState, CLNPaidAmt " & _
                "from View_Payment2CLN where PVCLNNo = '" & TxtPVNo & "'"
        PV.Open strsql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
        'Set objExcel = New excel.Application
        If PV.RecordCount Then
            RecordSelected = PV.RecordCount
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
        
            objExcel.WindowState = xlMaximized
    
        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\PV-Clinic.xlt")
        Set objWorksheet = objWorkbook.Sheets(1)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
        'Do While objWorkbook.Worksheets.Count > 1
         '   Set objWorksheet = objWorkbook.Worksheets.Item(objWorkbook.Worksheets.Count)
          '  objWorksheet.Delete
        'Loop
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
            
            objWorksheet.Cells(8, "C") = PV!CLNname
            objWorksheet.Cells(9, "C") = PV!CLNAdd1
            objWorksheet.Cells(10, "C") = PV!CLNAdd2 & "," & PV!CLNAdd3
            objWorksheet.Cells(11, "C") = Trim(PV!CLNPostCode & " " & Trim(PV!CLNCity) & " " & Trim(PV!CLNState))
            objWorksheet.Cells(8, "I") = TxtPVNo '& "/" & Format(PV!PVCLNDate, "MM/YY")
            objWorksheet.Cells(9, "I") = PV!PVCLNChequeNo
            objWorksheet.Cells(10, "I") = PV!PVCLNDate
            objWorksheet.Cells(37, "C") = Logon
            For i = 1 To RecordSelected
                ReceiptNo = ""
                REceipt.Open "SELECT ReceiptNo from REceiptFrPAY where CASNumber ='" & PV!CASNumber & "'", MediConnCISPayment, adOpenKeyset, adLockOptimistic
                If REceipt.RecordCount <> 0 Then
                    ReceiptNo = REceipt!ReceiptNo
                Else
                    REceipt.Close
                    Set REceipt = Nothing
                    REceipt.Open "SELECT ReceiptNo from REceiptFrPAY where CASNumber ='" & PV!CASNumber & "'", MediConnCISPayment, adOpenKeyset, adLockOptimistic
                    If REceipt.RecordCount <> 0 Then
                        ReceiptNo = REceipt!ReceiptNo
                    End If
                End If
                REceipt.Close
                Set REceipt = Nothing
                
                objWorksheet.Cells(15 + CDbl(i), "A") = i
                objWorksheet.Cells(15 + CDbl(i), "B") = PV!BillNo
                objWorksheet.Cells(15 + CDbl(i), "C") = PV!BillDate
                objWorksheet.Cells(15 + CDbl(i), "D") = PV!CASNumber
                objWorksheet.Cells(15 + CDbl(i), "E") = PV!BordxRefNo
                objWorksheet.Cells(15 + CDbl(i), "I") = PV!CLNPaidAmt
                objWorksheet.Cells(15 + CDbl(i), "G") = ReceiptNo
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

Private Sub CmdEdit_Click()
Dim ZZ As Integer
Dim ItemChk As Boolean

    EditFlag = True
    ItemChk = False
    For ZZ = 1 To lstPaymentCLN.ListItems.count
        If lstPaymentCLN.ListItems(ZZ).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next ZZ

    CmdSave.Enabled = False
    CmdUpdate.Enabled = False
    FramePV.Enabled = False
    If ItemChk = False And Len(Trim(TxtPVNo)) Then
        CmdUpdate.Enabled = True
        FramePV.Enabled = True
        EditFlag = True
        If Len(Trim(lstPaymentCLN.SelectedItem.SubItems(1))) Then
            FramePV.Enabled = True
        End If
        txtPVDate.SetFocus
    End If
End Sub

Public Sub CmdSearch_Click()
Dim listrecord As Integer

'    ScanCode = TxtWSNo
'    If Mid(ScanCode, 1, 2) = "LO" Then
'        TxtWSNo = ""
'        Call SearchForm(ScanCode)
'    Else
        
        Screen.MousePointer = vbHourglass
        LblAmt = 0
        EditFlag = False
        If ChkScan.Value = 1 Then
            If TxtWSNo = "" Then
                MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
                TxtWSNo.SetFocus
            Else
                listrecord = lstPaymentCLN.ListItems.count
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
                TxtWSNo = ""
                TxtWSNo.SetFocus
            End If
            TxtWSNo.SetFocus
        Else
            Call SearchRecord
        End If
 '   End If
    Screen.MousePointer = Default
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    Check = False
    TxtWSNo.SetFocus
    CmdExit.Enabled = True
End Sub

Private Sub CmdUpdate_Click()
Dim RecordSaved
Dim TEmpStatus As String
Dim TempPayAmt As String
Dim ZZ As Integer
Dim ItemChk As Boolean
Dim PVSql As String

    ItemChk = False
    For ZZ = 1 To lstPaymentCLN.ListItems.count
        If lstPaymentCLN.ListItems(ZZ).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next ZZ
    If ItemChk = True Then
    ElseIf lstPaymentCLN.ListItems.count = 0 Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    ElseIf Trim(txtPayIndex) = "" Then
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
            TEmpStatus = Mid(lstPaymentCLN.SelectedItem.SubItems(1), 1, 1)
            TempPayAmt = Trim(lstPaymentCLN.SelectedItem.SubItems(8))
            Call UpdateRecord(txtCaseNo, TEmpStatus, TempPayAmt, txtPayIndex)
    
            CmdSave.Enabled = True
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
    intMBMAccount = 1
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblAmt = 0
    Label25.Visible = False
    TxtWSNoTo.Visible = False
    Check = False
    EditFlag = False
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    CmdSave.Enabled = True
    FramePV.Enabled = False
    EditFlag = False
    lstPaymentCLN.ListItems.Clear
    CboPayment.ListIndex = 0
    lblCurrent = 0
    lblTotal = 0
    LblAmt = 0
    Check = False
    ChkScan.Value = 1
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
Dim monthstart
Dim yearstart
Dim HLT As New ADODB.Recordset
Dim BankRec As New ADODB.Recordset
    
'    HLT.Open "HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
'    Do Until HLT.EOF
'        cboHLTCode.AddItem ChkStr(HLT!HLTCode) & " - " & ChkStr(HLT!HLTDescription)
'        HLT.MoveNext
'    Loop
'    HLT.Close
'    Set HLT = Nothing
    
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
        .AddItem "3 - Partial Paid"
        .AddItem "4 - Paid & Partial Paid"
        .AddItem "5 - Unpaid & Partial Paid"
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
        
End Sub
'************************Start ComboListing**************************

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim GrpCompanyName As String
    
    
    strAppend = ""
    GrpCompanyName = ""
    
    If Len(Trim(CboPayor)) Then
         GrpCompanyName = Replace(CboPayor, "'", "''")
         strAppend = strAppend + " AND GroupCompany = '" & Trim(GrpCompanyName) & "'"
    End If
    If Len(Trim(txtOPBatchNo)) Then
        strAppend = strAppend + " AND OPBatchNo = '" & Trim(txtOPBatchNo) & "'"
    End If
    If Len(Trim(CboCLN)) Then
        strAppend = strAppend + " AND ClinicCode = '" & Trim(Left(CboCLN, IIf(InStr(1, CboCLN, "-") - 1 < 0, Len(CboCLN), InStr(1, CboCLN, "-") - 1))) & "'"
    End If
    
    If Mid(CboPayment, 1, 1) = 1 Then
        strAppend = strAppend + " and PVPayCLNStatus = 'F'"
    End If
 
    If Mid(CboPayment, 1, 1) = 2 Then
        strAppend = strAppend + " and (PVPayCLNStatus = '' or PVPayCLNStatus is null) "
    End If
 
    If Mid(CboPayment, 1, 1) = 3 Then
        strAppend = strAppend + " and PVPayCLNStatus = 'P'"
    End If
 
    If Mid(CboPayment, 1, 1) = 4 Then
        strAppend = strAppend + " and (PVPayCLNStatus = 'P' or PVPayCLNStatus = 'F') "
    End If
            
    If ChkScan.Value = 1 Then
        strAppend = strAppend + " AND CASNumber = '" & Trim(TxtWSNo) & "'"
    Else
        If Len(TxtWSNo) Then
            strAppend = strAppend + " AND CASNumber >= '" & Trim(TxtWSNo) & "'"
        End If
    
        If Len(TxtWSNoTo) Then
            strAppend = strAppend + " AND CASNumber <= '" & Trim(TxtWSNoTo) & "'"
        End If
    End If
    
    If Len(TxtInvNo) Then
        strAppend = strAppend + " and BillNo = '" & Trim(TxtInvNo) & "'"
    End If
    
    If ChkInvNo.Value = 1 Then
        sql = ""
        sql = "AND (BillNo Is null OR BillNo = '') " & ""
        strAppend = strAppend + sql
    End If
    
    If (txtInvDate1) <> "__/__/____" And IsDate(txtInvDate1) = True _
        And (txtInvDate2) <> "__/__/____" And IsDate(txtInvDate2) = True Then
        sql = ""
        sql = " AND BillDate >= '" & Format(Trim(txtInvDate1), "mm/dd/yyyy") & _
              "' And  BillDate <= '" & Format(Trim(txtInvDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtInvDate1)) > 0 And IsDate(txtInvDate1) = True Then
        strAppend = strAppend + " And BillDate >= '" & Format(txtInvDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtInvDate2)) > 0 And IsDate(txtInvDate2) = True Then
        strAppend = strAppend + " And BillDate <= '" & Format(txtInvDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkInvDate.Value = 1 Then
        sql = ""
        sql = "AND (PVDate Is null OR PVDate = '') " & ""
        strAppend = strAppend + sql
    End If
    
    If Len(txtInvAmt1) Then
        strAppend = strAppend + " and INVAmount >= '" & Trim(txtInvAmt1) & "'"
    End If
        
    If Len(TxtInvAmt2) Then
        strAppend = strAppend + " and INVAmount <= '" & Trim(TxtInvAmt2) & "'"
    End If
    
    If chkInvAmt.Value = 1 Then
        sql = ""
        sql = "AND (INVAmount Is null OR INVAmount = '') " & ""
        strAppend = strAppend + sql
    End If
    
    If Len(txtPVNo1) Then
        strAppend = strAppend + " and PVCLNNo >= '" & Trim(txtPVNo1) & "'"
    End If
        
    If Len(txtPVNo2) Then
        strAppend = strAppend + " and PVCLNNo <= '" & Trim(txtPVNo2) & "'"
    End If
    
    If chkPVNo.Value = 1 Then
        sql = ""
        sql = "AND (PVNo Is null OR PVNo = '') " & ""
        strAppend = strAppend + sql
    End If
    
    If (txtPVDate1) <> "__/__/____" And IsDate(txtPVDate1) = True _
        And (txtPVDate2) <> "__/__/____" And IsDate(txtPVDate2) = True Then
        sql = ""
        sql = " AND PVCLNDate >= '" & Format(Trim(txtPVDate1), "mm/dd/yyyy") & _
              "' And PVCLNDate <= '" & Format(Trim(txtPVDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtPVDate1)) > 0 And IsDate(txtPVDate1) = True Then
        strAppend = strAppend + " And PVCLNDate >= '" & Format(txtPVDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPVDate2)) > 0 And IsDate(txtPVDate2) = True Then
        strAppend = strAppend + " And PVCLNDate <= '" & Format(txtPVDate2, "mm/dd/yyyy") & "'"
    End If
    
    If chkPVDate.Value = 1 Then
        sql = ""
        sql = "AND (PVCLNDate Is null OR PVCLNDate = '') " & ""
        strAppend = strAppend + sql
    End If
        
    If Len(cboACPeriodMthFr) Then
        strAppend = strAppend + " and CLNAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
        
    If Len(cboACPeriodMthTo) Then
        strAppend = strAppend + " and CLNAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If chkPeriodMth.Value = 1 Then
        sql = ""
        sql = " AND CLNAccPeriodMth Is null " & ""
        strAppend = strAppend + sql
    End If
    
    If Len(cboACPeriodYRTo) Then
        strAppend = strAppend + " and CLNAccPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
    End If
        
    If Len(cboACPeriodYRTo) Then
        strAppend = strAppend + " and CLNAccPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
    
    If ChkPeriodYr.Value = 1 Then
        sql = ""
        sql = " AND CLNAccPeriodYr Is null " & ""
        strAppend = strAppend + sql
    End If
    strAppend = strAppend + " AND CLMPayableByMedix2c <> 0 "
    Call Payment2CLNListing(strAppend)

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

Private Sub Payment2CLNListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
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
Dim PatientDB As New ADODB.Recordset
Dim PatientSql As String

   total = 0
   Current = 0
   If ChkScan.Value = 0 Then
        lstPaymentCLN.ListItems.Clear
   End If

    
    sql = " Select * From VIEW_Payment2CLN where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
    
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
                Record.SubItems(2) = IIf(Len(rst2!ClinicCode), rst2!ClinicCode, "")
                Record.SubItems(3) = IIf(Len(rst2!CLNname), rst2!CLNname, "")
                Record.SubItems(4) = IIf(Len(rst2!InvoiceNo), rst2!InvoiceNo, "")
                If Len(rst2!InvoiceDate) Then
                    Record.SubItems(5) = Format(rst2!InvoiceDate, "dd/mm/yyyy")
                End If
                Record.SubItems(6) = IIf(Len(rst2!BillNo), rst2!BillNo, "")
                
                If Len(rst2!CLMPayableByMedix2C) Then
                    Record.SubItems(7) = Format(rst2!CLMPayableByMedix2C, "#,##0.00")
                End If
                If Len(rst2!CLNPaidAmt) Then
                    Record.SubItems(8) = Format(rst2!CLNPaidAmt, "#,##0.00")
                End If
                
                If rst2!CoverID = "00" Then
                    PatientSql = "Select MBMName as PatientName from MBM where MBMNumber='" & rst2!MBMNumber & "'"
                Else
                    PatientSql = "Select MBMCName as PatientName from MBMCoveredPersons where MBMCNumber='" & rst2!MBMNumber & "' AND MBMCCoverID ='" & rst2!CoverID & "'"
                End If
                PatientDB.Open PatientSql, medixmasterconn, adOpenKeyset, adLockOptimistic

                If Len(PatientDB!PatientName) Then
                    Record.SubItems(11) = Trim(PatientDB!PatientName)
                End If
                PatientDB.Close
                Set PatientDB = Nothing
                
                If Len(rst2!CASNumber) Then
                    Record.SubItems(12) = rst2!CASNumber
                End If
                CASNumber = rst2!CASNumber
                'WSVersion = rst2!clmversion
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(13) = Trim(rst2!BordxRefNo)
                End If
                
                If Len(rst2!DateVisit) Then
                    Record.SubItems(14) = Format(rst2!DateVisit, "dd/mm/yyyy")
                End If
                
                If Len(rst2!PVCLNDate) Then
                    Record.SubItems(15) = Format(rst2!PVCLNDate, "dd/mm/yyyy")
                End If
                If Len(rst2!INVEnteredBy) Then
                    Record.SubItems(16) = Trim(rst2!INVEnteredBy)
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
                If IsNumeric(rst2!PVCLNChequeAmt) Then
                    Record.SubItems(21) = Trim(rst2!PVCLNChequeAmt)
                End If

                If Len(Trim(rst2!CLNAccPeriodMth) & Trim(rst2!CLNAccPeriodYr)) Then
                    Record.SubItems(22) = Format(rst2!CLNAccPeriodMth & "/" & rst2!CLNAccPeriodYr, "mm/yyyy")
                End If
                If Len(Trim(rst2!PVCLNBank)) Then
                    Record.SubItems(23) = Trim(rst2!PVCLNBank)
                End If
                If Len(Trim(rst2!PVCLNAccount)) Then
                    Record.SubItems(24) = Trim(rst2!PVCLNAccount)
                End If
                If Len(Trim(rst2!CLNRemark)) Then
                    Record.SubItems(25) = Trim(rst2!CLNRemark)
                End If
                If Len(Trim(rst2!OPBatchNo)) Then
                    Record.SubItems(26) = Trim(rst2!OPBatchNo)
                End If
                strsqlPAY = "Select * from Payment2CLN where CASNumber = '" & Trim(CASNumber) & "'"
                PAY.Open strsqlPAY, MediConnCISPayment, adOpenKeyset, adLockOptimistic
                INVAmt = 0
                If Trim(txtPVStatus) = "P" Then
                    Do Until PAY.EOF
                        INVAmt = INVAmt + PAY!CLNPaidAmt
                    PAY.MoveNext
                    Loop
                    If Len(INVAmt) Then
                        Record.SubItems(8) = INVAmt
                        Record.SubItems(9) = rst2!CLMPayableByMedix2C - INVAmt
                    End If
                ElseIf Trim(txtPVStatus) = "" Then
                    If Len(rst2!CLMPayableByMedix2C) Then
                        Record.SubItems(9) = rst2!CLMPayableByMedix2C
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

Private Sub CmdSave_Click()
Dim PAYCode As String
Dim caseno As String
Dim status As String
Dim PaymentAmt As String
Dim ii As Integer
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

    Cnt = 0
    ItemChecked = False
    caseno = ""
    status = ""
    PaymentAmt = ""
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

    If lstPaymentCLN.ListItems.count = 0 Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
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
'            If Mid(cboHLTCode, 1, 1) = "S" Then
                PVNumber = GenerateSPVNo(TempBank, TmpPVCode)
'            Else
'                PVNumber = GenerateNPVNo(TempBank,TmpPVCode)
'            End If
                TxtPVNo = ""
                TxtPVNo = PVNumber
           
            For i = 1 To lstPaymentCLN.ListItems.count
                If lstPaymentCLN.ListItems(i).Checked = True Then
                    ItemChecked = True
                Exit For
                End If
            Next i
            
            If ItemChecked = True Then
                For i = 1 To lstPaymentCLN.ListItems.count
                     If lstPaymentCLN.ListItems(i).Checked = True Then
                        caseno = Trim(lstPaymentCLN.ListItems(i).SubItems(12))
                        status = Trim(Mid(lstPaymentCLN.ListItems(i).SubItems(1), 1, 1))
                        PaymentAmt = Trim(lstPaymentCLN.ListItems(i).SubItems(10))
                        PymtIndex = Trim(lstPaymentCLN.ListItems(i).SubItems(17))
                        Call UpdateRecord(caseno, status, PaymentAmt, PymtIndex)
                
                        caseno = ""
                        status = ""
                        PaymentAmt = ""
                        PymtIndex = ""
                    End If
                Next i
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                lstPaymentCLN.ListItems.Clear
                LblAmt = 0
        End If
        PVSql = " and PVCLNNo = '" & Trim(TxtPVNo) & "'"
        Call ClearFramePV
        lstPaymentCLN.ListItems.Clear
        Call Payment2CLNListing(PVSql)
    End If
End If
End Sub

Private Sub UpdateRecord(caseno As String, status As String, PaymentAmt As Variant, PymtIndex As String)
Dim Update
Dim strsql As String
Dim PAY As New ADODB.Recordset
Dim strsqlPAY As String
Dim CLM As New ADODB.Recordset
Dim strsqlCLM As String
Dim ZZ As String
Dim TempBank As String

    ZZ = InStr(1, CboBank, "-")
    TempBank = Trim(Left(CboBank, IIf(ZZ - 1 <= 0, Len(Trim(CboBank)), ZZ - 1)))

    strsqlPAY = " Select PVCLNIndex, * from Payment2CLN " & _
                " Where PVCLNIndex = '" & Trim(PymtIndex) & "'"
    PAY.Open strsqlPAY, MediConnCISPayment, adOpenKeyset, adLockOptimistic
    
    If PAY.RecordCount > 0 Then
        If EditFlag = False Then
            If Trim(PAY!PVPayCLNStatus) = "P" Then
                PAY!PVPayCLNStatus = Mid(status, 1, 1)
                PAY.Update
                PAY.AddNew
            End If
        End If
    Else
        PAY.AddNew
    End If
    With PAY
        !CASNumber = Trim(caseno)
        !PVCLNNo = TxtPVNo
        !PVCLNDate = Format(txtPVDate, "dd/mm/yyyy")
        !PVCLNChequeDate = IIf(txtChequeDate <> "__/__/____", Format(txtChequeDate, "dd/mm/yyyy"), "")
        !PVCLNChequeNo = IIf(Len(Trim(txtChequeNo)), txtChequeNo, "")
        !PVCLNChequeAmt = IIf(IsNumeric(txtChequeAmt), txtChequeAmt, 0)
        !PVCLNBank = IIf(Len(TempBank), Trim(TempBank), "")
        !PVCLNAccount = IIf(Len(Trim(txtAcNo)), Trim(txtAcNo), "")
        !CLNRemark = IIf(Len(Trim(TxtRemark)), TxtRemark, "")
        !CLNPaidAmt = IIf(IsNumeric(PaymentAmt), PaymentAmt, 0)
        !CLNAccPeriodMth = IIf(CboACPeriodMth <> "", Trim(CboACPeriodMth), "")
        !CLNAccPeriodYr = IIf(CboACPeriodYr <> "", Trim(CboACPeriodYr), "")
        !Pay2CLNLastUpdateDate = Date
        !Pay2CLNLastUpdateUser = Logon
        !PVPayCLNStatus = Mid(status, 1, 1)
    End With
    PAY.Update
    
    PAY.Close
    Set PAY = Nothing
End Sub

Private Sub lstPaymentCLN_Click()
    
    If lstPaymentCLN.ListItems.count Then
        If Trim(lstPaymentCLN.SelectedItem.SubItems(17)) <> "" Then
            CmdSave.Enabled = False
            FramePV.Enabled = False
            CboBank.Enabled = False
            txtPayIndex = lstPaymentCLN.SelectedItem.SubItems(17)
            txtCaseNo = lstPaymentCLN.SelectedItem.SubItems(12)
            TxtPVNo = lstPaymentCLN.SelectedItem.SubItems(18)
            If Len(lstPaymentCLN.SelectedItem.SubItems(15)) Then
                txtPVDate = lstPaymentCLN.SelectedItem.SubItems(15)
            End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(19)) Then
                txtChequeNo = lstPaymentCLN.SelectedItem.SubItems(19)
            End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(20)) Then
                txtChequeDate = lstPaymentCLN.SelectedItem.SubItems(20)
            End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(21)) Then
                txtChequeAmt = lstPaymentCLN.SelectedItem.SubItems(21)
            End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(22)) Then
                CboACPeriodMth = Mid(lstPaymentCLN.SelectedItem.SubItems(22), 1, 2)
                CboACPeriodYr = Mid(lstPaymentCLN.SelectedItem.SubItems(22), 4, 4)
            End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(23)) Then
                CboBank = lstPaymentCLN.SelectedItem.SubItems(23)
            End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(24)) Then
                txtAcNo = lstPaymentCLN.SelectedItem.SubItems(24)
            End If
            If Len(lstPaymentCLN.SelectedItem.SubItems(25)) Then
                TxtRemark = lstPaymentCLN.SelectedItem.SubItems(25)
            End If
        Else
            Call ClearFramePV
        End If
    End If
End Sub

Private Sub lstPaymentCLN_DblClick()
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim caseno As String

    CmdSave.Enabled = True
    If lstPaymentCLN.SelectedItem.Checked = True Then
        MsgBox "Please Uncheck the record!", vbCritical
    Else
        caseno = Trim(Mid(lstPaymentCLN.SelectedItem.SubItems(12), 1, IIf(InStr(1, lstPaymentCLN.SelectedItem.SubItems(12), "-") = 0, 10, InStr(1, lstPaymentCLN.SelectedItem.SubItems(12), "-") - 1)))
            
        strsqlCLM = "Select * From Payment2CLN where CASNumber = '" & Trim(caseno) & "'"
        CLM.Open strsqlCLM, MediConnCISPayment, adOpenKeyset, adLockOptimistic
                
        If CLM.RecordCount Then
            If Trim(CLM!PVPayCLNStatus) = "F" Then
                MsgBox "Full Payment has been made", vbCritical
                lstPaymentCLN.SelectedItem.Checked = False
            ElseIf Trim(CLM!PVPayCLNStatus) = "P" Then
                    MsgBox "Partial Payment has been made", vbCritical
            End If
        Else
            FrmMBMAccount.txtOSAmt = lstPaymentCLN.SelectedItem.SubItems(9)
            FrmMBMAccount.txtCLMNo = lstPaymentCLN.SelectedItem.SubItems(12)
            FrmMBMAccount.Source = 3
            FrmMBMAccount.Show
        End If
    End If
End Sub

Private Sub lstPaymentCLN_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    CmdSave.Enabled = True
    EditFlag = False
    Call ProcessCheckedItem(Item)
End Sub

Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
Dim i  As Integer
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim caseno, CASNum, PAYIndex, PYMTStatus As String
Dim Payment2CLN As New ADODB.Recordset
    
    caseno = Trim(Item.SubItems(12))

    strsqlPAY = "Select * From Payment2CLN where CASNumber = '" & Trim(caseno) & "'"
    PAY.Open strsqlPAY, MediConnCISPayment, adOpenKeyset, adLockPessimistic

    If Item.Checked = True Then
        If PAY.RecordCount <> 0 Then
            If Trim(PAY!PVPayCLNStatus) = "F" Then
                MsgBox "Payment has been made!", vbCritical
                Item.Checked = False
            Else
                Item.SubItems(10) = CDbl(Item.SubItems(9))
                Item.SubItems(1) = "Full"
                
                LblAmt = LblAmt + CDbl(Format(Item.SubItems(10), "#,##0.00"))
            End If
        Else
            If Partial = False Then
                Item.SubItems(10) = IIf(Trim(Item.SubItems(8)) = "", CDbl(Item.SubItems(7)), CDbl(Item.SubItems(9)))
                Item.SubItems(1) = "Full"
            End If
            LblAmt = LblAmt + CDbl(Format(Item.SubItems(10), "#,##0.00"))
        End If
    Else
        LblAmt = LblAmt - CDbl(Format(Item.SubItems(10), "#,##0.00"))
        Item.SubItems(10) = ""
        If PAY.RecordCount Then
            If Len((PAY!PVPayCLNStatus)) Then
                Item.SubItems(1) = "Partial"
            End If
        Else
            Item.SubItems(1) = ""
        End If
    End If
    PAY.Close
    Set PAY = Nothing
End Function


Private Sub txtAcNo_LostFocus()
    txtAcNo = ChkStr(txtAcNo)
End Sub

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

Private Sub txtChequeNo_LostFocus()
    txtChequeNo = ChkStr(txtChequeNo)
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

Private Sub TxtInvNo_Change()

End Sub

Private Sub TxtInvNo_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtInvNo_LostFocus()
    TxtInvNo = ChkStr(TxtInvNo)
    CmdSearch.Default = False
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
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
    Case 12
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(14)
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
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(22)
    End Select
End Sub

Private Sub cboPayor_Change()
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

Private Sub TxtRemark_LostFocus()
    TxtRemark = ChkStr(TxtRemark)
End Sub

Private Sub TxtWSNo_GotFocus()
'    CmdSearch.Default = True
End Sub

Private Sub TxtWSNo_LostFocus()
    TxtWSNo = ChkStr(TxtWSNo)
'    CmdSearch.Default = False
End Sub

Private Sub TxtWSNoTo_LostFocus()
    TxtWSNoTo = ChkStr(TxtWSNoTo)
End Sub

Private Function CaseNoCheck()
    If Trim(ChkStr(TxtWSNo)) = Trim(Mid(lstPaymentCLN.ListItems(listloop1).SubItems(12), 1, IIf(InStr(1, lstPaymentCLN.ListItems(listloop1).SubItems(12), "-") = 0, 10, InStr(1, lstPaymentCLN.ListItems(listloop1).SubItems(12), "-") - 1))) Then
        Check = True
    End If
End Function

Private Sub ClearFramePV()
    TxtPVNo = ""
    txtPVDate.Text = "__/__/____"
    txtPayIndex = ""
    txtChequeNo = ""
    txtChequeDate = "__/__/____"
    txtChequeAmt = ""
    CboACPeriodMth.Text = ""
    CboACPeriodYr.Text = ""
    CboBank.Text = ""
    txtAcNo = ""
    txtCaseNo = ""
    TxtRemark = ""
End Sub




