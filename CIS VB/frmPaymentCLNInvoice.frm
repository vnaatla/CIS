VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPaymentCLNInvoice 
   Caption         =   "Payment to Clinic By Invoice"
   ClientHeight    =   8076
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8076
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkAppendRcd 
      Caption         =   "Append Records"
      Height          =   255
      Left            =   4920
      TabIndex        =   116
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Frame FramePV 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   120
      TabIndex        =   47
      Top             =   5760
      Width           =   11775
      Begin VB.ComboBox CboPayto 
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
         Left            =   840
         TabIndex        =   59
         Top             =   150
         Width           =   5290
      End
      Begin VB.TextBox txtCLNEmail 
         Enabled         =   0   'False
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
         Left            =   6840
         MaxLength       =   30
         TabIndex        =   58
         Top             =   150
         Width           =   4695
      End
      Begin VB.TextBox txtChequeAmt 
         Enabled         =   0   'False
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
         Left            =   5040
         MaxLength       =   15
         TabIndex        =   57
         Top             =   435
         Width           =   1095
      End
      Begin VB.ComboBox CboACPeriodMth 
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
         ItemData        =   "frmPaymentCLNInvoice.frx":0000
         Left            =   5280
         List            =   "frmPaymentCLNInvoice.frx":0002
         TabIndex        =   56
         Top             =   1395
         Width           =   855
      End
      Begin VB.ComboBox CboACPeriodYr 
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
         ItemData        =   "frmPaymentCLNInvoice.frx":0004
         Left            =   6840
         List            =   "frmPaymentCLNInvoice.frx":0006
         TabIndex        =   55
         Top             =   1395
         Width           =   975
      End
      Begin VB.TextBox txtChequeNo 
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
         Left            =   840
         MaxLength       =   20
         TabIndex        =   54
         Top             =   440
         Width           =   1215
      End
      Begin VB.TextBox txtAmtWording 
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
         Left            =   840
         MaxLength       =   500
         TabIndex        =   53
         Top             =   690
         Width           =   5295
      End
      Begin VB.TextBox TxtPVNo 
         Enabled         =   0   'False
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
         Left            =   840
         MaxLength       =   15
         TabIndex        =   52
         Top             =   950
         Width           =   1215
      End
      Begin VB.ComboBox CboBank 
         Enabled         =   0   'False
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
         ItemData        =   "frmPaymentCLNInvoice.frx":0008
         Left            =   4440
         List            =   "frmPaymentCLNInvoice.frx":000A
         TabIndex        =   51
         Text            =   "MB - MAYBANK"
         Top             =   950
         Width           =   1695
      End
      Begin VB.TextBox txtAccNo 
         Enabled         =   0   'False
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
         Left            =   6840
         MaxLength       =   30
         TabIndex        =   50
         Top             =   410
         Width           =   4695
      End
      Begin VB.TextBox TxtRemark 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
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
         TabIndex        =   49
         Top             =   660
         Width           =   4695
      End
      Begin VB.TextBox txtCaseNo 
         Height          =   285
         Left            =   10080
         TabIndex        =   48
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSMask.MaskEdBox txtChequeDate 
         Height          =   315
         Left            =   3000
         TabIndex        =   60
         Top             =   435
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   572
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox txtPVDate 
         Height          =   315
         Left            =   2760
         TabIndex        =   61
         Top             =   945
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   550
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label Label6 
         Caption         =   "Wording"
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
         Left            =   120
         TabIndex        =   74
         Top             =   700
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Pay To"
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
         Left            =   120
         TabIndex        =   73
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "Email"
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
         Left            =   6240
         TabIndex        =   72
         Top             =   165
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "Chq. Date"
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
         Left            =   2160
         TabIndex        =   71
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Remark"
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
         TabIndex        =   70
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "A/C Mth"
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
         Left            =   4320
         TabIndex        =   69
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label24 
         Caption         =   "A/C Yr"
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
         TabIndex        =   68
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Chq. Amt"
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
         Left            =   4320
         TabIndex        =   67
         Top             =   465
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Chq. No"
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
         TabIndex        =   66
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "PV Date"
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
         Left            =   2160
         TabIndex        =   65
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "PV No"
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
         TabIndex        =   64
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Bank"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   4080
         TabIndex        =   63
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label27 
         Caption         =   "A/C No"
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
         Left            =   6240
         TabIndex        =   62
         Top             =   450
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   11775
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
         Left            =   7560
         TabIndex        =   29
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox ChkInvNo 
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
         Left            =   5280
         TabIndex        =   28
         Top             =   210
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
         Left            =   7080
         TabIndex        =   27
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkPVDate 
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
         Left            =   11280
         TabIndex        =   26
         Top             =   255
         Width           =   375
      End
      Begin VB.CheckBox chkPVNo 
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
         Left            =   5280
         TabIndex        =   25
         Top             =   1020
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
         Left            =   8040
         TabIndex        =   24
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox ChkClinic 
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
         Left            =   5280
         TabIndex        =   23
         Top             =   720
         Width           =   255
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
         Left            =   6840
         TabIndex        =   22
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxtInvNo 
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
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   0
         Top             =   180
         Width           =   1560
      End
      Begin VB.TextBox txtInvNo2 
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
         Left            =   3600
         MaxLength       =   30
         TabIndex        =   1
         Top             =   180
         Width           =   1575
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
         Left            =   8040
         MaxLength       =   20
         TabIndex        =   17
         Top             =   1680
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
         Left            =   8400
         MaxLength       =   20
         TabIndex        =   16
         Top             =   1680
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
         Left            =   8400
         TabIndex        =   15
         Top             =   1650
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox CboPayment 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   440
         Width           =   3615
      End
      Begin VB.ComboBox CboCLN 
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtPVNo1 
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1010
         Width           =   1575
      End
      Begin VB.TextBox txtPVNo2 
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
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1010
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtPVDate1 
         Height          =   315
         Left            =   7920
         TabIndex        =   6
         Top             =   180
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
      Begin MSMask.MaskEdBox txtPVDate2 
         Height          =   315
         Left            =   9840
         TabIndex        =   7
         Top             =   180
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
         ItemData        =   "frmPaymentCLNInvoice.frx":000C
         Left            =   7680
         List            =   "frmPaymentCLNInvoice.frx":000E
         TabIndex        =   21
         Top             =   1680
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
         ItemData        =   "frmPaymentCLNInvoice.frx":0010
         Left            =   7320
         List            =   "frmPaymentCLNInvoice.frx":0012
         TabIndex        =   20
         Top             =   1680
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
         ItemData        =   "frmPaymentCLNInvoice.frx":0014
         Left            =   7920
         List            =   "frmPaymentCLNInvoice.frx":0016
         TabIndex        =   19
         Top             =   1680
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
         ItemData        =   "frmPaymentCLNInvoice.frx":0018
         Left            =   7680
         List            =   "frmPaymentCLNInvoice.frx":001A
         TabIndex        =   18
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtBordxDateFr 
         Height          =   315
         Left            =   6120
         TabIndex        =   94
         Top             =   1680
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
         Left            =   6960
         TabIndex        =   95
         Top             =   1680
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
         Left            =   5040
         TabIndex        =   8
         Top             =   1665
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
         Left            =   6960
         TabIndex        =   9
         Top             =   1665
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
         TabIndex        =   10
         Top             =   1920
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
         Left            =   6960
         TabIndex        =   11
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboINVMonthFr 
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
         Left            =   7920
         TabIndex        =   104
         Top             =   450
         Width           =   1335
      End
      Begin VB.ComboBox cboINVMonthTo 
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
         Left            =   9840
         TabIndex        =   105
         Top             =   450
         Width           =   1335
      End
      Begin VB.ComboBox cboINVYearTo 
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
         Left            =   9840
         TabIndex        =   107
         Top             =   730
         Width           =   1335
      End
      Begin VB.ComboBox cboINVYearFr 
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
         Left            =   7920
         TabIndex        =   106
         Top             =   730
         Width           =   1335
      End
      Begin VB.ComboBox cboEntityType 
         Height          =   315
         ItemData        =   "frmPaymentCLNInvoice.frx":001C
         Left            =   7920
         List            =   "frmPaymentCLNInvoice.frx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   1020
         Width           =   3255
      End
      Begin VB.Label lblEntityType 
         Caption         =   "Entity Type"
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
         Left            =   6360
         TabIndex        =   114
         ToolTipText     =   "Date Pass On"
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Inv Month"
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
         Left            =   6360
         TabIndex        =   111
         Top             =   500
         Width           =   1455
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
         Index           =   7
         Left            =   9480
         TabIndex        =   110
         Top             =   500
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "INV Year"
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
         Left            =   6360
         TabIndex        =   109
         Top             =   800
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
         Index           =   5
         Left            =   9480
         TabIndex        =   108
         Top             =   800
         Width           =   1455
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
         Left            =   6600
         TabIndex        =   102
         Top             =   1965
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
         TabIndex        =   101
         Top             =   1965
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
         TabIndex        =   100
         Top             =   1695
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
         Left            =   3480
         TabIndex        =   99
         Top             =   1695
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
         TabIndex        =   97
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label19 
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
         Left            =   7200
         TabIndex        =   96
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
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
         Left            =   7920
         TabIndex        =   46
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         Left            =   7680
         TabIndex        =   45
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
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
         Left            =   3240
         TabIndex        =   44
         Top             =   195
         Width           =   375
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   8280
         TabIndex        =   43
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
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
         TabIndex        =   42
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label31 
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
         Left            =   8640
         TabIndex        =   41
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label32 
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
         Left            =   8040
         TabIndex        =   40
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
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
         Left            =   5760
         TabIndex        =   39
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label19 
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
         Left            =   9450
         TabIndex        =   38
         Top             =   195
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "PV Date"
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
         Left            =   6360
         TabIndex        =   37
         Top             =   195
         Width           =   855
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
         Left            =   3240
         TabIndex        =   36
         Top             =   1035
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "PV Number"
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
         TabIndex        =   35
         Top             =   1035
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Payment Status"
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
         TabIndex        =   34
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Invoice No"
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
         TabIndex        =   33
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Group Company"
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
         Left            =   6960
         TabIndex        =   32
         Top             =   1725
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Clinic"
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
         TabIndex        =   31
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5340
         TabIndex        =   30
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.TextBox txtPVStatus 
      Height          =   285
      Left            =   5160
      TabIndex        =   12
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView lstPaymentCLN 
      Height          =   3675
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   6477
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
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   40
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
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Inv Month"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Inv Year"
         Object.Width           =   1764
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
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "OS Amt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Amt Paying"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Entry By"
         Object.Width           =   0
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
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Cheque No"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "Cheque Date"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Chq Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Text            =   "A/C Period"
         Object.Width           =   0
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
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "OPBatch"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Email"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "AccNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Payee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "WordingOne"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "CLNAdd1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "CLNAdd2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "CLNAdd3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "Bill Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "Co-Pay"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   31
         Text            =   "App Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "CLN Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "CLN Group"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "FFS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "Payee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "Registration Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Text            =   "FFS %"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Text            =   "Bank Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Text            =   "A/C No"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   1020
      Left            =   120
      TabIndex        =   75
      Top             =   7040
      Width           =   11775
      Begin VB.TextBox LblAmt 
         Height          =   285
         Left            =   840
         TabIndex        =   113
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   840
         TabIndex        =   112
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmbPrinPVRange 
         Caption         =   "Print PV"
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
         Left            =   8040
         TabIndex        =   98
         Top             =   500
         Width           =   735
      End
      Begin VB.TextBox txttmpDollar 
         Height          =   285
         Left            =   3360
         TabIndex        =   90
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txttmpCents 
         Height          =   285
         Left            =   4200
         TabIndex        =   89
         Top             =   240
         Visible         =   0   'False
         Width           =   735
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
         Left            =   10800
         TabIndex        =   88
         Top             =   500
         Width           =   855
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
         Left            =   9600
         TabIndex        =   87
         Top             =   500
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrintChqRange 
         Caption         =   "Print Chq"
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
         Left            =   8760
         TabIndex        =   86
         Top             =   500
         Width           =   855
      End
      Begin VB.CommandButton cmdPrintPV 
         Caption         =   "P&V"
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
         Left            =   5400
         TabIndex        =   85
         Top             =   500
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdPreviewPV 
         Caption         =   "Previ&w PV"
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
         Left            =   4080
         TabIndex        =   84
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
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
         Height          =   330
         Left            =   8040
         TabIndex        =   83
         Top             =   180
         Width           =   735
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
         Height          =   330
         Left            =   8760
         TabIndex        =   82
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
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
         Left            =   10200
         TabIndex        =   81
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
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
         Left            =   9600
         TabIndex        =   80
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdCheck 
         Caption         =   "&Check All"
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
         Left            =   10800
         TabIndex        =   79
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
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
         Left            =   8040
         TabIndex        =   78
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
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
         Left            =   8760
         TabIndex        =   77
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdPrintChq 
         Caption         =   "Print Chq"
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
         Left            =   8640
         TabIndex        =   76
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
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
         TabIndex        =   92
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Label Label11 
      Caption         =   "* Please double-click listing to view invoice details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   103
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "   Payment to Clinic - Invoice"
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
      TabIndex        =   93
      Top             =   0
      Width           =   11925
   End
End
Attribute VB_Name = "frmPaymentCLNInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intExit As Integer
Dim m_blnDirection(1 To 37) As Boolean
Dim EditFlag As Boolean
Dim Check As Boolean
Dim tmpPayee As String
Dim tmpClinicCode As String
Dim tmpInvoiceNo As String
Dim DuplicateStatus As Boolean

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

Private Sub CboBank_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboBank_KeyPress"
On Error GoTo debugs

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
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub cboINVMonthFr_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "cboINVMonthFr_KeyPress"
On Error GoTo debugs

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboINVMonthFr.Text, cboINVMonthFr.SelStart) + Chr(KeyAscii) + Mid(cboINVMonthFr.Text, cboINVMonthFr.SelStart + cboINVMonthFr.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboINVMonthFr.ListCount - 1
 
        If NewText = LCase(Left(cboINVMonthFr.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboINVMonthFr.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboINVMonthFr.Text = ValidValue
                cboINVMonthFr.SelStart = 0
                cboINVMonthFr.SelLength = Len(ValidValue)
            End If
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub cboINVMonthTo_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "cboINVMonthTo_KeyPress"
On Error GoTo debugs

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboINVMonthTo.Text, cboINVMonthTo.SelStart) + Chr(KeyAscii) + Mid(cboINVMonthTo.Text, cboINVMonthTo.SelStart + cboINVMonthTo.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboINVMonthTo.ListCount - 1
 
        If NewText = LCase(Left(cboINVMonthTo.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboINVMonthTo.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboINVMonthTo.Text = ValidValue
                cboINVMonthTo.SelStart = 0
                cboINVMonthTo.SelLength = Len(ValidValue)
            End If
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
        
End Sub

Private Sub cboINVYearFr_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "cboINVYearFr_KeyPress"
On Error GoTo debugs

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboINVYearFr.Text, cboINVYearFr.SelStart) + Chr(KeyAscii) + Mid(cboINVYearFr.Text, cboINVYearFr.SelStart + cboINVYearFr.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboINVYearFr.ListCount - 1
 
        If NewText = LCase(Left(cboINVYearFr.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboINVYearFr.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboINVYearFr.Text = ValidValue
                cboINVYearFr.SelStart = 0
                cboINVYearFr.SelLength = Len(ValidValue)
            End If
        End If
        
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cboINVYearTo_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "cboINVYearTo_KeyPress"
On Error GoTo debugs

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboINVYearTo.Text, cboINVYearTo.SelStart) + Chr(KeyAscii) + Mid(cboINVYearTo.Text, cboINVYearTo.SelStart + cboINVYearTo.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboINVYearTo.ListCount - 1
 
        If NewText = LCase(Left(cboINVYearTo.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboINVYearTo.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboINVYearTo.Text = ValidValue
                cboINVYearTo.SelStart = 0
                cboINVYearTo.SelLength = Len(ValidValue)
            End If
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ChkClinic_Click()
Const METHOD_NAME = "ChkClinic_Click"
On Error GoTo debugs

Dim ClnRec As New ADODB.Recordset

    CboCLN.Clear
    If ChkClinic.Value = 1 Then
        ClnRec.Open "Select CLNCode, CLNName from CLN WHERE CLNPanel = 'Y' OR CLNPanel = 'C'", medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do Until ClnRec.EOF
            If Len(ClnRec!CLNName) Then
                CboCLN.AddItem ClnRec!CLNCode & " - " & Trim(ClnRec!CLNName)
            End If
            ClnRec.MoveNext
        Loop
        ClnRec.Close
        Set ClnRec = Nothing
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmbPrinPVRange_Click()
    frmPrintPVRange.chqMode = ""
    Load frmPrintPVRange
    frmPrintPVRange.chqMode = "MedixPayCLN"
    frmPrintPVRange.Show vbModal
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
        'txtPVDate.SetFocus
        CmdSave.Enabled = True
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
Const METHOD_NAME = "cmdPrintChq_Click"
On Error GoTo debugs

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
    '    On Error GoTo HANDLE_ERROR
        ' Don't allow user to affect workbook
    '        strsql = "select CASNumber, ClaimVersion, BordxRefNo, MBMNumber, MBMPatientCovID, " & _
    '                " HLTCode, ClaimCloseDate, PVNo, PVChequeNo, PVDate, MBMPaidAmt, CLMPayableByMedix2M, " & _
    '                " PVBankCharges from View_Payment2MBMNew where PVNo = '" & TxtPVNo & "' AND PVStatus='A'"
    
            
        strsql = "select PVCLNChequeNo, PVCLNChequeDate, PVCLNChequeAmt, " & _
        " PVWordingOne,PVWordingTwo, CLNPName from VIEW_Payment2CLN_Chq where PVCLNChequeNo = '" & Trim(txtChequeNo) & "'"
        chq.Open strsql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
        
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
        objWorkbook.SaveAs (crdPath & "Template\PaymentVoucher\PaymentCLN\" & "tmpPaycheque" & ".xlt")
        objExcel.Quit

ExitFunction:
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            Screen.MousePointer = vbDefault
            Exit Sub
'HANDLE_ERROR:
'            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
'            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
'            Set objWorksheet = Nothing
'            Set objWorkbook = Nothing
'            objExcel.Quit
'            GoTo ExitFunction
    End If
'    Screen.MousePointer = vbDefault
    
    
debugs:
Screen.MousePointer = vbDefault
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
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
Const METHOD_NAME = "CmdEdit_Click"
On Error GoTo debugs

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
    CboPayto.Enabled = True
    txtChequeNo.Enabled = True
    txtChequeDate.Enabled = True
    txtChequeAmt.Enabled = True
    txtPVDate.Enabled = True
    
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
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If


End Sub

Private Sub CmdSearch_Click()
Const METHOD_NAME = "cmdSearch_Click"
On Error GoTo debugs


    Screen.MousePointer = vbHourglass
    CmdEdit.Visible = True
    cmdAdd.Visible = True
    CmdUpdate.Visible = False
    CmdSave.Visible = False
    LblAmt = 0
    EditFlag = False
    Call SearchRecord
    Screen.MousePointer = Default
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    Check = False
    TxtInvNo.SetFocus
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
Dim tmpAmt As String

    Screen.MousePointer = vbHourglass

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
        
            'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
            'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
            
                    StrInvSql = "Select CASNumber, PVPayCLNStatus, PVCLNIndex, CLNPaidAmt" & _
                                 " from Payment2CLN WHERE PVCLNNo='" & Trim(TxtPVNo) & _
                                 "' AND PVPayCLNStatus='F'"
                                
                    UPdRec.Open StrInvSql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
                    'If UPdRec.RecordCount Then
                        Do While Not UPdRec.EOF
                            CaseNo = ""
                            PVStatus = ""
                            PymtIndex = ""
                            CaseNo = Trim(UPdRec!CasNumber)
                            PVStatus = "F"
                            PaymentAmt = UPdRec!CLNPaidAmt
                            PymtIndex = UPdRec!PVCLNIndex
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
                            
                            Call UpdateRecord(CaseNo, PVStatus, PaymentAmt, PymtIndex)
                            UPdRec.MoveNext
                        Loop
                    'End If
                    UPdRec.Close
                    Set UPdRec = Nothing
                    CmdSave.Enabled = True
                    CmdEdit.Visible = True
                    cmdAdd.Visible = True
                    CmdUpdate.Visible = False
                    CmdSave.Visible = False
                    EditFlag = False
                    LblAmt = 0
                    PVSql = " PVCLNNo = '" & Trim(TxtPVNo) & "'"
                    Call ClearFramePV
                    lstPaymentCLN.ListItems.Clear
                    Call Payment2CLNListing(PVSql)
                    CboPayto.Enabled = False
                    txtChequeNo.Enabled = False
                    txtChequeDate.Enabled = False
                    txtChequeAmt.Enabled = False
                    txtPVDate.Enabled = False
        End If
    End If
    Screen.MousePointer = Default

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "frmPaymentCLNInvoice")
    Call ComboListing
    CboPayment.ListIndex = 0
    intExit = 0
    lblCurrent = 0
    LblAmt = 0
    Check = False
    EditFlag = False
    cboEntityType.ListIndex = 0
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
    Unload Me
End Sub

'************************Start ComboListing**************************
Private Sub ComboListing()
Const METHOD_NAME = "ComboListing"
On Error GoTo debugs

Dim monthstart
Dim yearstart
Dim BankRec As New ADODB.Recordset
Dim ClnRec As New ADODB.Recordset
        BankRec.Open "Banks", medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
        Do While Not BankRec.EOF
            CboBank.AddItem ChkStr(BankRec!BankCode) & " - " & ChkStr(BankRec!BankDesc)
            BankRec.MoveNext
        Loop
        BankRec.Close
        Set BankRec = Nothing
    
    With CboPayment
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
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
        
    
    For monthstart = 1 To 12
        cboINVMonthFr.AddItem Format(monthstart, "00")
        cboINVMonthTo.AddItem Format(monthstart, "00")
    Next monthstart
    
    For yearstart = 1999 To 3000
        cboINVYearFr.AddItem yearstart
        cboINVYearTo.AddItem yearstart
    Next yearstart
        
    cboINVYearFr.Text = Year(Date)
    cboINVYearTo.Text = Year(Date)

    'CboPayto.Clear
    'ClnRec.Open "Select CLNPName from CLN WHERE CLNPanel = 'Y' OR CLNPanel = 'C'", medixmasterconn, adOpenKeyset, adLockOptimistic
    'Do Until ClnRec.EOF
    '    If Len(ClnRec!CLNPName) Then
    '        CboPayto.AddItem ClnRec!CLNPName
    '    End If
    '    ClnRec.MoveNext
    'Loop
    'ClnRec.Close
    'Set ClnRec = Nothing
    

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Function SearchRecordTopup()
Const METHOD_NAME = "SearchRecordTopup"
On Error GoTo debugs
Dim strsql As String
Dim StrAppend As String
Dim sql As String
Dim GrpCompanyName As String
    
    StrAppend = ""
    GrpCompanyName = ""
    
    StrAppend = StrAppend & " SUBSTRING(HISDB.dbo.MBM.MBMNumber, 1, 2)  <> 'EM' and SUBSTRING(HISDB.dbo.MBM.MBMNumber, 1, 2)  <> 'ES' "
   
   If Len(lstPaymentCLN.SelectedItem.SubItems(5)) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and CASInvYear >= '" & Trim(lstPaymentCLN.SelectedItem.SubItems(5)) & "'"
        Else
            StrAppend = StrAppend & " CASInvYear >= '" & Trim(lstPaymentCLN.SelectedItem.SubItems(5)) & "'"
        End If
    End If
    
    If Len(lstPaymentCLN.SelectedItem.SubItems(5)) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and CASInvYear <= '" & Trim(lstPaymentCLN.SelectedItem.SubItems(5)) & "'"
        Else
            StrAppend = StrAppend & " CASInvYear <= '" & Trim(lstPaymentCLN.SelectedItem.SubItems(5)) & "'"
        End If
    End If

    
    If Len(lstPaymentCLN.SelectedItem.SubItems(4)) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and CASInvMonth >= '" & Trim(lstPaymentCLN.SelectedItem.SubItems(4)) & "'"
        Else
            StrAppend = StrAppend & " CASInvMonth >= '" & Trim(lstPaymentCLN.SelectedItem.SubItems(4)) & "'"
        End If
    End If
    
    If Len(lstPaymentCLN.SelectedItem.SubItems(4)) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and CASInvMonth <= '" & Trim(lstPaymentCLN.SelectedItem.SubItems(4)) & "'"
        Else
            StrAppend = StrAppend & " CASInvMonth <= '" & Trim(lstPaymentCLN.SelectedItem.SubItems(4)) & "'"
        End If
    End If
    
    If StrAppend <> "" Then
        StrAppend = StrAppend & " and (CLMPayableByMedix2C  <> '0') "
    Else
        StrAppend = StrAppend & " (CLMPayableByMedix2C  <> '0') "
    End If
    
    If StrAppend <> "" Then
        StrAppend = StrAppend & " and (INVPayTo  = 'C') "
    Else
        StrAppend = StrAppend & " (INVPayTo  = 'C') "
    End If
    
    
    If Mid(CboPayment, 1, 1) = 1 Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and Payment2CLN.PVPayCLNStatus = 'F'"
        Else
            StrAppend = StrAppend & " Payment2CLN.PVPayCLNStatus = 'F'"
        End If
    End If
 
    If Mid(CboPayment, 1, 1) = 2 Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and (Payment2CLN.PVPayCLNStatus = '' or Payment2CLN.PVPayCLNStatus is null) "
            'StrAppend = StrAppend & " and PVPayCLNStatus <> 'F' "
        Else
            StrAppend = StrAppend & "  (Payment2CLN.PVPayCLNStatus = '' or Payment2CLN.PVPayCLNStatus is null) "
        End If
    End If
    
    If Len(txtPVNo1) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and PVCLNNo >= '" & Trim(txtPVNo1) & "'"
        Else
            StrAppend = StrAppend & " PVCLNNo >= '" & Trim(txtPVNo1) & "'"
        End If
    End If
        
    If Len(txtPVNo2) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and PVCLNNo <= '" & Trim(txtPVNo2) & "'"
        Else
            StrAppend = StrAppend & " PVCLNNo <= '" & Trim(txtPVNo2) & "'"
        End If
    End If
    
    If chkPVNo.Value = 1 Then
        sql = ""
        If StrAppend <> "" Then
            sql = "AND (PVCLNNo Is null OR PVCLNNo = '') " & ""
        Else
            sql = "(PVCLNNo Is null OR PVCLNNo = '') " & ""
        End If
        StrAppend = StrAppend & sql
    End If
    
    If IsDate(txtPVDate1) = True Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " And PVCLNDate >= '" & Format(txtPVDate1, "mm/dd/yyyy") & "'"
        Else
            StrAppend = StrAppend & " PVCLNDate >= '" & Format(txtPVDate1, "mm/dd/yyyy") & "'"
        End If
    End If
    If IsDate(txtPVDate2) = True Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " And PVCLNDate <= '" & Format(txtPVDate2, "mm/dd/yyyy") & "'"
        Else
            StrAppend = StrAppend & " PVCLNDate <= '" & Format(txtPVDate2, "mm/dd/yyyy") & "'"
        End If
    End If
    If chkPVDate.Value = 1 Then
        sql = ""
        If StrAppend <> "" Then
            sql = "AND (PVCLNDate Is null OR PVCLNDate = '') " & ""
        Else
            sql = "(PVCLNDate Is null OR PVCLNDate = '') " & ""
        End If
        StrAppend = StrAppend & sql
    End If
    If StrAppend <> "" Then
        StrAppend = StrAppend & "And ClinicCode = '" & Trim(lstPaymentCLN.SelectedItem.SubItems(2)) & "'"
    Else
        StrAppend = StrAppend & "ClinicCode = '" & Trim(lstPaymentCLN.SelectedItem.SubItems(2)) & "'"
    End If
    Call Payment2CLNPopUp(StrAppend)
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim strsql As String
Dim StrAppend As String
Dim sql As String
Dim GrpCompanyName As String
    
    StrAppend = ""
    GrpCompanyName = ""
    
    StrAppend = StrAppend & " SUBSTRING(HISDB.dbo.MBM.MBMNumber, 1, 2)  <> 'EM' and SUBSTRING(HISDB.dbo.MBM.MBMNumber, 1, 2)  <> 'ES' "
   
    If Len(Trim(CboCLN)) Then
        StrAppend = StrAppend & " And ClinicCode = '" & Trim(Left(CboCLN, IIf(InStr(1, CboCLN, "-") - 1 < 0, Len(CboCLN), InStr(1, CboCLN, "-") - 1))) & "'"
    End If
    
    If Len(cboINVYearFr) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and CASInvYear >= '" & Trim(cboINVYearFr) & "'"
        Else
            StrAppend = StrAppend & " CASInvYear >= '" & Trim(cboINVYearFr) & "'"
        End If
    End If
    
    If Len(cboINVYearTo) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and CASInvYear <= '" & Trim(cboINVYearTo) & "'"
        Else
            StrAppend = StrAppend & " CASInvYear <= '" & Trim(cboINVYearTo) & "'"
        End If
    End If

    
    If Len(cboINVMonthFr) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and CASInvMonth >= '" & Trim(cboINVMonthFr) & "'"
        Else
            StrAppend = StrAppend & " CASInvMonth >= '" & Trim(cboINVMonthFr) & "'"
        End If
    End If
    
    If Len(cboINVMonthTo) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and CASInvMonth <= '" & Trim(cboINVMonthTo) & "'"
        Else
            StrAppend = StrAppend & " CASInvMonth <= '" & Trim(cboINVMonthTo) & "'"
        End If
    End If

    If Mid(CboPayment, 1, 1) = 1 Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and Payment2CLN.PVPayCLNStatus = 'F'"
        Else
            StrAppend = StrAppend & " Payment2CLN.PVPayCLNStatus = 'F'"
        End If
    End If
 
    If Mid(CboPayment, 1, 1) = 2 Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and (Payment2CLN.PVPayCLNStatus = '' or Payment2CLN.PVPayCLNStatus is null) "
        Else
            StrAppend = StrAppend & "  (Payment2CLN.PVPayCLNStatus = '' or Payment2CLN.PVPayCLNStatus is null) "
        End If
    End If
    
    If StrAppend <> "" Then
        StrAppend = StrAppend & " and (INVPayTo  = 'C') "
    Else
        StrAppend = StrAppend & " (INVPayTo  = 'C') "
    End If
    
    
    If StrAppend <> "" Then
        StrAppend = StrAppend & " and (CLMPayableByMedix2C  <> '0') "
    Else
        StrAppend = StrAppend & " (CLMPayableByMedix2C  <> '0') "
    End If
    
    If Len(txtPVNo1) Then
    
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and PVCLNNo >= '" & Trim(txtPVNo1) & "'"
        Else
            StrAppend = StrAppend & "  PVCLNNo >= '" & Trim(txtPVNo1) & "'"
        End If
    End If
        
    If Len(txtPVNo2) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and PVCLNNo <= '" & Trim(txtPVNo2) & "'"
        Else
            StrAppend = StrAppend & " PVCLNNo <= '" & Trim(txtPVNo2) & "'"
        End If
    End If
    
    If chkPVNo.Value = 1 Then
        sql = ""
        If StrAppend <> "" Then
            sql = "AND (PVCLNNo Is null OR PVCLNNo = '') " & ""
        Else
            sql = " (PVCLNNo Is null OR PVCLNNo = '') " & ""
        End If
        StrAppend = StrAppend & sql
    End If
    
    If IsDate(txtPVDate1) = True Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " And PVCLNDate >= '" & Format(txtPVDate1, "mm/dd/yyyy") & "'"
        Else
            StrAppend = StrAppend & " PVCLNDate >= '" & Format(txtPVDate1, "mm/dd/yyyy") & "'"
        End If
    End If
    
    If IsDate(txtPVDate2) = True Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " And PVCLNDate <= '" & Format(txtPVDate2, "mm/dd/yyyy") & "'"
        Else
            StrAppend = StrAppend & " PVCLNDate <= '" & Format(txtPVDate2, "mm/dd/yyyy") & "'"
        End If
    End If
    
    If chkPVDate.Value = 1 Then
        sql = ""
        If StrAppend <> "" Then
            sql = "AND (PVCLNDate Is null OR PVCLNDate = '') " & ""
        Else
            sql = " (PVCLNDate Is null OR PVCLNDate = '') " & ""
        End If
        StrAppend = StrAppend & sql
    End If
 
    If Len(cboACPeriodMthFr) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and CLNAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
        Else
            StrAppend = StrAppend & " CLNAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
        End If
    End If
        
    If Len(cboACPeriodMthTo) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and CLNAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
        Else
            StrAppend = StrAppend & "  CLNAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
        End If
    End If
    
    If chkPeriodMth.Value = 1 Then
        sql = ""
        If StrAppend <> "" Then
            sql = " AND CLNAccPeriodMth Is null "
        Else
            sql = "  CLNAccPeriodMth Is null "
        End If
        StrAppend = StrAppend & sql
    End If
    
    If Len(cboACPeriodYRTo) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and CLNAccPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
        Else
            StrAppend = StrAppend & " CLNAccPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
        End If
    End If
        
    If Len(cboACPeriodYRTo) Then
        If StrAppend <> "" Then
            StrAppend = StrAppend & " and CLNAccPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
        Else
            StrAppend = StrAppend & " CLNAccPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
        End If
    End If
    
    If cboEntityType.ListIndex = 0 Then
        StrAppend = StrAppend + " AND Substring(CISCAS.CasNumber, 1, 2) <>'SJ' "
    ElseIf cboEntityType.ListIndex = 1 Then
        StrAppend = StrAppend + " AND Substring(CISCAS.CasNumber, 1, 2) ='SJ' "
    End If
    
    If ChkPeriodYr.Value = 1 Then
        sql = ""
        If StrAppend <> "" Then
            sql = " AND CLNAccPeriodYr Is null "
        Else
            sql = "  CLNAccPeriodYr Is null "
        End If
        StrAppend = StrAppend & sql
    End If
 
    Call Payment2CLNListing(StrAppend)

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
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
Const METHOD_NAME = "Payment2CLNListing"
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
Dim tmpPay2CLN As Double
Dim StrCLNSql As String
Dim ClnRec As New ADODB.Recordset
Dim tmpPVCLNBank
Dim tmpPVCLNAccount As String
Dim tmpCLNRemark As String
Dim tmpCLNBankAcc As String
Dim tmpCLNPName As String
Dim tmpCLNPAdd1 As String
Dim tmpCLNPAdd2 As String
Dim tmpCLNPAdd3 As String
Dim tmpCLNType As String
Dim tmpCLNFFSStatus As String
Dim tmpCLNRegNo As String
Dim tmpCLNMCOProcfeePerc As String
Dim tmpCLNBankName As String
Dim tmpCLNName As String
Dim tmpCLNAdd1  As String
Dim tmpCLNAdd2  As String
Dim tmpCLNAdd3 As String
Dim tmpPCode  As String
Dim tmpCity  As String
Dim tmpState   As String
Dim tmpClinicGrpCode As String

    If chkAppendRcd.Value = 0 Then
        lstPaymentCLN.ListItems.Clear
    End If

    Current = 0
    'lstPaymentCLN.ListItems.Clear
    Call ClearFramePV
    Call ReadRecord(StrAppend, PVRec)
    Do
        Do While Not PVRec.EOF
        txtPVStatus = ""
        If Len(PVRec!PVPayCLNStatus) Then
            txtPVStatus = PVRec!PVPayCLNStatus
        End If
            
            With PVRec
                StrCLNSql = ""
                StrCLNSql = "SELECT Top 10 * FROM CLN WHERE CLNCode='" & Trim(PVRec!ClinicCode) & "'"
                ClnRec.Open StrCLNSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                Do While Not ClnRec.EOF
                    With ClnRec

                        tmpCLNName = Trim(!CLNName)
                        tmpCLNAdd1 = IIf(Len(!CLNAdd1), !CLNAdd1, "")
                        tmpCLNAdd2 = IIf(Len(!CLNAdd2), !CLNAdd2, "")
                        tmpCLNAdd3 = IIf(Len(!CLNAdd3), !CLNAdd3, "")
                        tmpPCode = IIf(Len(!CLNPostCode), !CLNPostCode, "")
                        tmpCity = IIf(Len(!CLNCity), !CLNCity, "")
                        tmpState = IIf(Len(!CLNState), !CLNState, "")
                        tmpCLNBankAcc = IIf(Len(!CLNBankAcc), !CLNBankAcc, "")
                        tmpCLNPName = IIf(Len(!CLNPName), !CLNPName, "")
                        tmpCLNPAdd1 = IIf(Len(!CLNPAdd1), !CLNPAdd1, "")
                        tmpCLNPAdd2 = IIf(Len(!CLNPAdd2), !CLNPAdd2, "")
                        tmpCLNPAdd3 = IIf(Len(!CLNPAdd3), !CLNPAdd3, "")
                        tmpCLNType = IIf(Len(!CLNType), !CLNType, "")
                        tmpCLNFFSStatus = IIf(Len(!CLNFFSStatus), !CLNFFSStatus, "")
                        tmpCLNPName = IIf(Len(!CLNPName), !CLNPName, "")
                        tmpCLNRegNo = IIf(Len(!CLNRegNo), !CLNRegNo, "")
                        tmpCLNMCOProcfeePerc = IIf(Len(!CLNMCOProcfeePerc), !CLNMCOProcfeePerc, "")
                        tmpCLNBankName = IIf(Len(!CLNBankName), !CLNBankName, "")
                        tmpCLNBankAcc = IIf(Len(!CLNBankAcc), !CLNBankAcc, "")
                        tmpClinicGrpCode = IIf(Len(!ClinicGrpCode), !ClinicGrpCode, "")
                    End With
                    ClnRec.MoveNext
                Loop
                ClnRec.Close
                Set ClnRec = Nothing

                Set Record = lstPaymentCLN.ListItems.Add(, , "")
                Record.SubItems(1) = ""
                If Trim(txtPVStatus) = "F" Then
                    Record.SubItems(1) = "Full"
                ElseIf Trim(txtPVStatus) = "P" Then
                    Record.SubItems(1) = "Partial"
                End If
                Record.SubItems(2) = IIf(Len(PVRec!ClinicCode), PVRec!ClinicCode, "")
                Record.SubItems(3) = IIf(Len(tmpCLNName), tmpCLNName, "")
                Record.SubItems(4) = IIf(Len(PVRec!CASINVMonth), PVRec!CASINVMonth, "")
                If Len(PVRec!CASINVYear) Then
                    Record.SubItems(5) = PVRec!CASINVYear
                End If
                tmpPay2CLN = IIf(IsNumeric(PVRec!CLMPayableByMedix2C), Format(PVRec!CLMPayableByMedix2C, "#,##0.00"), 0)
                Record.SubItems(6) = Format(tmpPay2CLN, "#,##0.00")
                Record.SubItems(7) = IIf(IsNumeric(PVRec!CLNPaidAmt), Format(PVRec!CLNPaidAmt, "#,##0.00"), 0)
                Record.SubItems(8) = CDbl(Record.SubItems(6) - Record.SubItems(7))
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
                
                If Len(Trim(!PVCLNBank)) Then
                    Record.SubItems(18) = Trim(!PVCLNBank)
                End If
                If Len(Trim(!PVCLNAccount)) Then
                    Record.SubItems(19) = Trim(!PVCLNAccount)
                End If
                
                If Len(Trim(!CLNRemark)) Then
                    Record.SubItems(20) = Trim(!CLNRemark)
                End If
                
                If Len(Trim(tmpCLNBankAcc)) Then
                    Record.SubItems(23) = Trim(tmpCLNBankAcc)
                End If
                
                If Len(Trim(tmpCLNPName)) Then
                    Record.SubItems(24) = Trim(tmpCLNPName)
                End If
                
                If Len(Trim(PVRec!PVWordingOne)) Then
                    Record.SubItems(25) = Trim(PVRec!PVWordingOne)
                End If
                
                If Len(Trim(PVRec!PVWordingTwo)) Then
                    Record.SubItems(25) = Record.SubItems(25) & PVRec!PVWordingTwo
                End If
                                  
                If Len(Trim(tmpCLNPAdd1)) Then
                    Record.SubItems(26) = Trim(tmpCLNPAdd1)
                End If
                
                If Len(Trim(tmpCLNPAdd2)) Then
                    Record.SubItems(27) = Trim(tmpCLNPAdd2)
                End If
                
                If Len(Trim(tmpCLNPAdd3)) Then
                    Record.SubItems(28) = Trim(tmpCLNPAdd3)
                End If
                
                Record.SubItems(29) = IIf(Len(PVRec!CLMTotalActual), Format(PVRec!CLMTotalActual, "#,##0.00"), 0#)
                Record.SubItems(30) = IIf(Len(PVRec!COPAY), Format(PVRec!COPAY, "#,##0.00"), 0#)
                Record.SubItems(31) = IIf(Len(PVRec!CLMFinalApprove), Format(PVRec!CLMFinalApprove, "#,##0.00"), 0#)
                                
                If Len(Trim(tmpCLNType)) Then
                    Record.SubItems(32) = Trim(tmpCLNType)
                End If
                
                If Len(Trim(tmpClinicGrpCode)) Then
                    Record.SubItems(33) = Trim(tmpClinicGrpCode)
                End If
                
                If Len(Trim(tmpCLNFFSStatus)) Then
                    Record.SubItems(34) = Trim(tmpCLNFFSStatus)
                End If
                
                If Len(Trim(tmpCLNPName)) Then
                    Record.SubItems(35) = Trim(tmpCLNPName)
                End If
                
                If Len(Trim(tmpCLNRegNo)) Then
                    Record.SubItems(36) = Trim(tmpCLNRegNo)
                End If
                
                If Len(Trim(tmpCLNMCOProcfeePerc)) And cboEntityType.ListIndex = 0 Then
                    Record.SubItems(37) = Trim(tmpCLNMCOProcfeePerc)
                ElseIf cboEntityType.ListIndex = 1 Then
                    Record.SubItems(37) = Trim("NA")
                End If
                
                If Len(Trim(tmpCLNBankName)) Then
                    Record.SubItems(38) = Trim(tmpCLNBankName)
                End If
                
                If Len(Trim(tmpCLNBankAcc)) Then
                    Record.SubItems(39) = Trim(tmpCLNBankAcc)
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
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub CmdSave_Click()
Const METHOD_NAME = "CmdSave_Click"
On Error GoTo debugs

On Err GoTo ErrorSave
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
Dim pymtcnt As Integer
Dim chqcode As String
Dim tmpchqno As String
Dim tmpPVNoFr As String
Dim tmpPVNoto As String
Dim startTrans As Boolean
    Screen.MousePointer = vbHourglass


   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
   ' MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    tmpPVNoFr = ""
    tmpPVNoto = ""
    pymtcnt = 0
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
    'For i = 1 To lstPaymentCLN.ListItems.Count
    '    If lstPaymentCLN.ListItems(i).Checked = True Then
    '        CLNCnt = CLNCnt + 1
    '        CurrCLN = Trim(lstPaymentCLN.ListItems(i).SubItems(2))
    '        If CLNCnt = 1 Then
    '            PrevCLN = Trim(lstPaymentCLN.ListItems(i).SubItems(2))
    '        End If
    '        If PrevCLN <> CurrCLN Then
    '            DiffCLN = True
    '            Exit For
    '        End If
    '    End If
    'Next i

    If ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    ElseIf txtChequeNo = "" Then
        MsgBox "Please Enter Cheque Number!", vbCritical
        txtChequeNo.SetFocus
    Else
        Update = MsgBox("Do you want to update the record ? ", vbYesNo + vbExclamation)
        If Update = vbYes Then
        
            'startTrans = True
            'MediConnCISPayment.beginTrans
            
            ' New Code Start here
            
            Dim SubTotal As Double
            Dim ClinicCode As String
            Dim IsFirst As Boolean
            Dim IsUpdated As Boolean
            Dim lstPosition As Integer
            Dim MnthSubqry As String
            Dim subQry As String
            subQry = ""
            MnthSubqry = ""
            IsFirst = True
            IsUpdated = True
            lstPosition = 0
            SubTotal = 0
            
            For i = 1 To lstPaymentCLN.ListItems.Count
                If lstPaymentCLN.ListItems(i).Checked = True Then
                    
                    If IsFirst Then
                        ClinicCode = lstPaymentCLN.ListItems(i).SubItems(2)
                        IsFirst = False
                        IsUpdated = False
                    End If
                    
                    If ClinicCode = lstPaymentCLN.ListItems(i).SubItems(2) Then
                        SubTotal = SubTotal + CDbl(lstPaymentCLN.ListItems(i).SubItems(6))
                        'write mnth sub queery
                        If cboEntityType.ListIndex = 0 Then
                            MnthSubqry = MnthSubqry + "or   (dbo.CISCas.CASINVMonth >= '" & Trim(lstPaymentCLN.ListItems(i).SubItems(4)) & "' And dbo.CISCas.CASINVMonth <= '" & Trim(lstPaymentCLN.ListItems(i).SubItems(4)) & "')  AND Substring(CISCAS.CasNumber, 1, 2) <>'SJ' "
                        ElseIf cboEntityType.ListIndex = 1 Then
                            MnthSubqry = MnthSubqry + "or   (dbo.CISCas.CASINVMonth >= '" & Trim(lstPaymentCLN.ListItems(i).SubItems(4)) & "' And dbo.CISCas.CASINVMonth <= '" & Trim(lstPaymentCLN.ListItems(i).SubItems(4)) & "')  AND Substring(CISCAS.CasNumber, 1, 2) ='SJ' "
                        End If
                        lstPosition = i
                    Else
                                                
                        'If SubTotal <> 0 Then
                        If SubTotal <> 0 Then
                                pymtcnt = pymtcnt + 1
                                If cboEntityType.ListIndex = 0 Then
                                    PVNumber = GenerateSPVNo(TempBank, TmpPVCode)
                                ElseIf cboEntityType.ListIndex = 1 Then
                                    PVNumber = GenerateSPVNoforSJ("MB", "CS")
                                End If
                                
                                TxtPVNo = ""
                                TxtPVNo = PVNumber
                                CboPayto = lstPaymentCLN.ListItems(lstPosition).SubItems(24)
                                txtChequeDate = Format(Date, "DD/MM/YYYY")
                                'txtChequeAmt = lstPaymentCLN.ListItems(i).SubItems(6)
                                txtChequeAmt = SubTotal
                                txtPVDate = Format(Date, "DD/MM/YYYY")
                                If CDbl(pymtcnt) = "1" Then
                                    txtChequeNo = txtChequeNo
                                    tmpPVNoFr = PVNumber
                                Else
                                    chqcode = Mid(txtChequeNo, 1, 3)
                                    tmpchqno = Mid(txtChequeNo, 4, 8)
                                    tmpchqno = Format(tmpchqno + 1, "000000")
                                    txtChequeNo = chqcode & tmpchqno
                                    tmpPVNoto = PVNumber
                                End If
                                StrInvSql = ""
                                        
                            startTrans = True
                            MediConnCISPayment.beginTrans
                            
                            subQry = ""
                            subQry = Mid(MnthSubqry, 3)
                            StrInvSql = ""
                            StrInvSql = "SELECT dbo.CISCas.CASNumber, CISWorksheet.dbo.Bordx.BordxDate, CISPayment.dbo.Payment2CLN.PVCLNNo, CISWorksheet.dbo.Claim.CLMPayableByMedix2C " & _
                            "FROM HISDB.dbo.MBM INNER JOIN CISWorksheet.dbo.Claim INNER JOIN " & _
                            " dbo.CISCas ON CISWorksheet.dbo.Claim.CASNumber = dbo.CISCas.CASNumber ON HISDB.dbo.MBM.MBMNumber = dbo.CISCas.MBMnumber LEFT OUTER JOIN " & _
                            " CISPayment.dbo.Payment2CLN ON CISWorksheet.dbo.Claim.CASNumber = CISPayment.dbo.Payment2CLN.CASNumber LEFT OUTER JOIN " & _
                            " CISWorksheet.dbo.Bordx ON CISWorksheet.dbo.Claim.BordxRefNo = CISWorksheet.dbo.Bordx.BordxRefNo " & _
                            " WHERE ClinicCode='" & Trim(lstPaymentCLN.ListItems(lstPosition).SubItems(2)) & "' " & _
                            " and  (CISDB.dbo.CISCas.InvoiceNo IS NOT NULL) AND  CISDB.dbo.CISCas.InvoiceNo<>N'' and dbo.CISCas.CASINVYear >= '" & Trim(lstPaymentCLN.ListItems(lstPosition).SubItems(5)) & "'" & _
                            " and dbo.CISCas.CASINVYear <= '" & Trim(lstPaymentCLN.ListItems(lstPosition).SubItems(5)) & "'" & _
                            " and CISPayment.dbo.Payment2CLN.PVCLNNo IS NULL AND Substring(HISDB.dbo.MBM.MBMNumber,1,2) <> 'EM'  AND Substring(HISDB.dbo.MBM.MBMNumber,1,2) <> 'ES' and  ( " & subQry & ")"
                            
                            UPdRec.Open StrInvSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                    
                        Do While Not UPdRec.EOF
                           CaseNo = ""
                           PVStatus = ""
                           PymtIndex = ""
                           CaseNo = Trim(UPdRec!CasNumber)
                           PVStatus = Mid(lstPaymentCLN.ListItems(lstPosition).SubItems(1), 1, 1)
                           tmpAmtPay2C = IIf(IsNumeric(UPdRec!CLMPayableByMedix2C), UPdRec!CLMPayableByMedix2C, 0)
                           PymtIndex = ""
                           Call UpdateRecord(CaseNo, PVStatus, tmpAmtPay2C, PymtIndex)
                           UPdRec.MoveNext
                        Loop
                   
                        UPdRec.Close
                        Set UPdRec = Nothing
                        CaseNo = ""
                        PVStatus = ""
                        PymtIndex = ""
                        MediConnCISPayment.commitTrans
                        startTrans = False
                    End If
                        
                        'Before this call the code to update
                        ClinicCode = lstPaymentCLN.ListItems(i).SubItems(2)
                        SubTotal = 0
                        MnthSubqry = ""
                        'Set to string null and assign back query
                        If cboEntityType.ListIndex = 0 Then
                            MnthSubqry = "or   (dbo.CISCas.CASINVMonth >= '" & Trim(lstPaymentCLN.ListItems(i).SubItems(4)) & "' And dbo.CISCas.CASINVMonth <= '" & Trim(lstPaymentCLN.ListItems(i).SubItems(4)) & "') AND Substring(CISCAS.CasNumber, 1, 2) <>'SJ' "
                        ElseIf cboEntityType.ListIndex = 1 Then
                            MnthSubqry = "or   (dbo.CISCas.CASINVMonth >= '" & Trim(lstPaymentCLN.ListItems(i).SubItems(4)) & "' And dbo.CISCas.CASINVMonth <= '" & Trim(lstPaymentCLN.ListItems(i).SubItems(4)) & "') AND Substring(CISCAS.CasNumber, 1, 2) ='SJ' "
                        End If
                        lstPosition = i
                        SubTotal = SubTotal + CDbl(lstPaymentCLN.ListItems(i).SubItems(6))
                        IsUpdated = False
                    End If
                   
                End If
                
                 If i = lstPaymentCLN.ListItems.Count Then
                        If Not IsUpdated Then
                            'Call the function that has the value
                             If SubTotal > 0 Then
                                pymtcnt = pymtcnt + 1
                                'PVNumber = GenerateSPVNo(TempBank, TmpPVCode)
                                If cboEntityType.ListIndex = 0 Then
                                    PVNumber = GenerateSPVNo(TempBank, TmpPVCode)
                                ElseIf cboEntityType.ListIndex = 1 Then
                                    PVNumber = GenerateSPVNoforSJ("MB", "CS")
                                End If
                                
                                TxtPVNo = ""
                                TxtPVNo = PVNumber
                                CboPayto = lstPaymentCLN.ListItems(lstPosition).SubItems(24)
                                txtChequeDate = Format(Date, "DD/MM/YYYY")
                                'txtChequeAmt = lstPaymentCLN.ListItems(i).SubItems(6)
                                txtChequeAmt = SubTotal
                                txtPVDate = Format(Date, "DD/MM/YYYY")
                                If CDbl(pymtcnt) = "1" Then
                                    txtChequeNo = txtChequeNo
                                    tmpPVNoFr = PVNumber
                                Else
                                    chqcode = Mid(txtChequeNo, 1, 3)
                                    tmpchqno = Mid(txtChequeNo, 4, 8)
                                    tmpchqno = Format(tmpchqno + 1, "000000")
                                    txtChequeNo = chqcode & tmpchqno
                                    tmpPVNoto = PVNumber
                                End If
                                StrInvSql = ""
                                        
                            startTrans = True
                            
                            MediConnCISPayment.beginTrans
                            
                          
                            
                            subQry = ""
                            subQry = Mid(MnthSubqry, 3)
                            StrInvSql = ""
                            StrInvSql = "SELECT dbo.CISCas.CASNumber, CISWorksheet.dbo.Bordx.BordxDate, CISPayment.dbo.Payment2CLN.PVCLNNo, CISWorksheet.dbo.Claim.CLMPayableByMedix2C " & _
                            "FROM HISDB.dbo.MBM INNER JOIN CISWorksheet.dbo.Claim INNER JOIN " & _
                            " dbo.CISCas ON CISWorksheet.dbo.Claim.CASNumber = dbo.CISCas.CASNumber ON HISDB.dbo.MBM.MBMNumber = dbo.CISCas.MBMnumber LEFT OUTER JOIN " & _
                            " CISPayment.dbo.Payment2CLN ON CISWorksheet.dbo.Claim.CASNumber = CISPayment.dbo.Payment2CLN.CASNumber LEFT OUTER JOIN " & _
                            " CISWorksheet.dbo.Bordx ON CISWorksheet.dbo.Claim.BordxRefNo = CISWorksheet.dbo.Bordx.BordxRefNo " & _
                            " WHERE ClinicCode='" & Trim(lstPaymentCLN.ListItems(lstPosition).SubItems(2)) & "' " & _
                            " and  (CISDB.dbo.CISCas.InvoiceNo IS NOT NULL) AND  CISDB.dbo.CISCas.InvoiceNo<>N'' and dbo.CISCas.CASINVYear >= '" & Trim(lstPaymentCLN.ListItems(lstPosition).SubItems(5)) & "'" & _
                            " and dbo.CISCas.CASINVYear <= '" & Trim(lstPaymentCLN.ListItems(lstPosition).SubItems(5)) & "'" & _
                            " and CISPayment.dbo.Payment2CLN.PVCLNNo IS NULL AND Substring(HISDB.dbo.MBM.MBMNumber,1,2) <> 'EM'  AND Substring(HISDB.dbo.MBM.MBMNumber,1,2) <> 'ES' and ( " & subQry & ")"
                        
                            UPdRec.Open StrInvSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                            Dim UpdateAmount As Double
                            Dim UpdateCnt As Integer
                            UpdateAmount = 0
                            UpdateCnt = 0
                            
                            Do While Not UPdRec.EOF
                               CaseNo = ""
                               PVStatus = ""
                               PymtIndex = ""
                               CaseNo = Trim(UPdRec!CasNumber)
                               PVStatus = Mid(lstPaymentCLN.ListItems(lstPosition).SubItems(1), 1, 1)
                               tmpAmtPay2C = IIf(IsNumeric(UPdRec!CLMPayableByMedix2C), UPdRec!CLMPayableByMedix2C, 0)
                               UpdateAmount = UpdateAmount + CDbl(tmpAmtPay2C)
                               UpdateCnt = UpdateCnt + 1
                               PymtIndex = ""
                               Call UpdateRecord(CaseNo, PVStatus, tmpAmtPay2C, PymtIndex)
                               UPdRec.MoveNext
                            Loop
                        
                             UPdRec.Close
                             Set UPdRec = Nothing
                             CaseNo = ""
                             PVStatus = ""
                             PymtIndex = ""
                             MediConnCISPayment.commitTrans
                             startTrans = False
                        End If
                        SubTotal = 0
                        MnthSubqry = ""
                    End If
                 End If
            Next i
            
            'End of the new code
           'Old code Start here
          '  For i = 1 To lstPaymentCLN.ListItems.Count
              '  If lstPaymentCLN.ListItems(i).Checked = True Then
             '       pymtcnt = pymtcnt + 1
            '        PVNumber = GenerateSPVNo(TempBank, TmpPVCode)
           '         TxtPVNo = ""
          '          TxtPVNo = PVNumber
         '           CboPayto = lstPaymentCLN.ListItems(i).SubItems(24)
        '            txtChequeDate = Format(Date, "DD/MM/YYYY")
       '             txtChequeAmt = lstPaymentCLN.ListItems(i).SubItems(6)
      '              txtPVDate = Format(Date, "DD/MM/YYYY")
     '               If CDbl(pymtcnt) = "1" Then
    '                    txtChequeNo = txtChequeNo
   '                     tmpPVNoFr = PVNumber
                '    Else
               '         chqcode = Mid(txtChequeNo, 1, 3)
              '          tmpchqno = Mid(txtChequeNo, 4, 8)
             '           tmpchqno = Format(tmpchqno + 1, "000000")
            '            txtChequeNo = chqcode & tmpchqno
           '             tmpPVNoto = PVNumber
          '          End If
         '           StrInvSql = ""
                      'StrInvSql = "SELECT     dbo.CISCas.CASNumber, CISWorksheet.dbo.Bordx.BordxDate " & _
                      '            " FROM CISWorksheet.dbo.Bordx RIGHT OUTER JOIN " & _
                      '          " CISWorksheet.dbo.Claim INNER JOIN " & _
                      '          " dbo.CISCas ON CISWorksheet.dbo.Claim.CASNumber = dbo.CISCas.CASNumber ON " & _
                      '          " CISWorksheet.dbo.Bordx.BordxRefNo = CISWorksheet.dbo.Claim.BordxRefNo" & _
                      '          " WHERE ClinicCode='" & Trim(lstPaymentCLN.ListItems(i).SubItems(2)) & "' " & _
                      '          " And Month(CISWorksheet.dbo.Bordx.BordxDate) >= '" & Trim(cboBordxMonthfr) & "'" & _
                      '          " And Month(CISWorksheet.dbo.Bordx.BordxDate) <= '" & Trim(cboBordxMonthTo) & "'" & _
                      '          " and Year(CISWorksheet.dbo.Bordx.BordxDate) >= '" & Trim(cboBordxYearfr) & "'" & _
                      '          " and Year(CISWorksheet.dbo.Bordx.BordxDate) <= '" & Trim(cboBordxYearTo) & "'"
                    
        '            startTrans = True
       '             MediConnCISPayment.beginTrans
      ' StrInvSql = "SELECT dbo.CISCas.CASNumber, CISWorksheet.dbo.Bordx.BordxDate, CISPayment.dbo.Payment2CLN.PVCLNNo, CISWorksheet.dbo.Claim.CLMPayableByMedix2C " & _
     '               "FROM HISDB.dbo.MBM INNER JOIN CISWorksheet.dbo.Claim INNER JOIN " & _
                    '" dbo.CISCas ON CISWorksheet.dbo.Claim.CASNumber = dbo.CISCas.CASNumber ON HISDB.dbo.MBM.MBMNumber = dbo.CISCas.MBMnumber LEFT OUTER JOIN " & _
                    '" CISPayment.dbo.Payment2CLN ON CISWorksheet.dbo.Claim.CASNumber = CISPayment.dbo.Payment2CLN.CASNumber LEFT OUTER JOIN " & _
                    '" CISWorksheet.dbo.Bordx ON CISWorksheet.dbo.Claim.BordxRefNo = CISWorksheet.dbo.Bordx.BordxRefNo " & _
                    '" WHERE ClinicCode='" & Trim(lstPaymentCLN.ListItems(i).SubItems(2)) & "' " & _
                    '" And dbo.CISCas.CASINVMonth >= '" & Trim(lstPaymentCLN.ListItems(i).SubItems(4)) & "'" & _
                    '" And dbo.CISCas.CASINVMonth <= '" & Trim(lstPaymentCLN.ListItems(i).SubItems(4)) & "'" & _
                    '" and dbo.CISCas.CASINVYear >= '" & Trim(lstPaymentCLN.ListItems(i).SubItems(5)) & "'" & _
                    '" and dbo.CISCas.CASINVYear <= '" & Trim(lstPaymentCLN.ListItems(i).SubItems(5)) & "'" & _
                    '" and CISPayment.dbo.Payment2CLN.PVCLNNo IS NULL AND Substring(HISDB.dbo.MBM.MBMNumber,1,2) <> 'EM'  AND Substring(HISDB.dbo.MBM.MBMNumber,1,2) <> 'ES'"
                    
                    'UPdRec.Open StrInvSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                    
                    
                    'If UPdRec.RecordCount Then
                        'Do While Not UPdRec.EOF
                         '   strsqlpymt = "Select CASNumber " & _
                         '       " from Payment2CLN WHERE CASNumber = '" & UPdRec!CasNumber & "'"
                         '       pymt.Open strsqlpymt, MediConnCISPayment, adOpenKeyset, adLockOptimistic
                         '   If pymt.RecordCount Then
                         '       PYMTSTATUS = True
                         '   Else
                         '       PYMTSTATUS = False
                         '   End If
                         '   pymt.Close
                         '   Set pymt = Nothing
                            
                          '  If PYMTSTATUS = False Then
                                'strlsqlCLM = "Select CLMPayableByMedix2C  from Claim WHERE CASNumber ='" & UPdRec!CasNumber & "'"
                                'CLM.Open strlsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
                                
                             '   CaseNo = ""
                            '    PVStatus = ""
                           '     PymtIndex = ""
                          '      CaseNo = Trim(UPdRec!CasNumber)
                         '       PVStatus = Mid(lstPaymentCLN.ListItems(i).SubItems(1), 1, 1)
                        '        tmpAmtPay2C = IIf(IsNumeric(UPdRec!CLMPayableByMedix2C), UPdRec!CLMPayableByMedix2C, 0)
                       '         PymtIndex = ""
                                'CLM.Close
                                'Set CLM = Nothing
                      '          Call UpdateRecord(CaseNo, PVStatus, tmpAmtPay2C, PymtIndex)
                                
                           ' End If
                     '       UPdRec.MoveNext
                    '    Loop
                   ' End If
                   ' UPdRec.Close
                  '  Set UPdRec = Nothing
                 '   CaseNo = ""
                '    PVStatus = ""
               '     PymtIndex = ""
              '      MediConnCISPayment.CommitTrans
             '       startTrans = False
                    
            '    End If
           ' Next i
           'Old code End here
           
            CmdSave.Visible = False
            cmdAdd.Visible = True
            lstPaymentCLN.ListItems.Clear
            LblAmt = 0
            PVSql = " PVCLNNo >= '" & Trim(tmpPVNoFr) & "' and PVCLNNo <= '" & Trim(tmpPVNoto) & "'"
            Call ClearFramePV
            lstPaymentCLN.ListItems.Clear
            Call Payment2CLNListing(PVSql)
        
            
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
        
        End If
    End If
    
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
    'MediConnCISPayment.Close
    'Set MediConnCISPayment = Nothing
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
    Screen.MousePointer = Default
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
        If startTrans = True Then
            MediConnCISPayment.RollbackTrans
        End If
        
    If Err.Description <> "" Then
        Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
    End If
    
debugs:
        Screen.MousePointer = vbDefault
        If startTrans = True Then
            MediConnCISPayment.RollbackTrans
        End If
        If Err.Description <> "" Then
            Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
        End If
End Sub

Private Sub UpdateRecord(CaseNo As String, status As String, PaymentAmt As Double, PymtIndex As String)

Const METHOD_NAME = "UpdateRecord"
On Error GoTo debugs

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
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
 
End Sub


Private Sub lstPaymentCLN_Click()
Const METHOD_NAME = "lstPaymentCLN_Click"
On Error GoTo debugs

    
    If lstPaymentCLN.ListItems.Count Then
        CmdEdit.Visible = True
        cmdAdd.Visible = True
        CmdUpdate.Visible = False
        CmdSave.Visible = False

        ClearFramePV
        'If Trim(lstPaymentCLN.SelectedItem.SubItems(17)) <> "" Then
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
                txtChequeDate = Format(lstPaymentCLN.SelectedItem.SubItems(15), "dd/MM/YYYY")
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
    End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstPaymentCLN_DblClick()
    Call SearchRecordTopup
End Sub

Private Sub lstPaymentCLN_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    CmdSave.Enabled = True
    EditFlag = False
    DuplicateStatus = False
    Call ProcessCheckedItem(Item)
End Sub

Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
Const METHOD_NAME = "ProcessCheckedItem"
On Error GoTo debugs

    If Item.Checked = True Then
        tmpClinicCode = Item.SubItems(2)
        tmpInvoiceNo = Item.SubItems(4)
        Call CheckDuplicate
        
        If Item.SubItems(24) <> "" And Trim(Item.SubItems(24)) <> "" And Trim(Item.SubItems(26)) <> "" And DuplicateStatus = False Then
            If Mid(Item.SubItems(1), 1, 1) <> "" Then
                MsgBox "Payment has been made!", vbCritical
                Item.Checked = False
            Else
                Item.SubItems(9) = CDbl(Item.SubItems(8))
                Item.SubItems(1) = "Full"
                
                LblAmt = LblAmt + CDbl(Format(Item.SubItems(9), "#,##0.00"))
            End If
        Else
            If DuplicateStatus = True Then
                MsgBox "Please Check Duplicate Records! ", vbCritical
            ElseIf Trim(Item.SubItems(24)) = "" Then
                MsgBox "Please Check Payee! ", vbCritical
            ElseIf Trim(Item.SubItems(26)) = "" Then
                MsgBox "Please Check Payee's Address 1 ! ", vbCritical
            End If
            Item.SubItems(1) = ""
            Item.Checked = False
        End If
    Else
        LblAmt = LblAmt - CDbl(Format(Item.SubItems(9), "#,##0.00"))
        Item.SubItems(9) = ""
        Item.SubItems(1) = ""
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Sub lstPaymentCLN_KeyDown(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "lstPaymentCLN_KeyDown"
On Error GoTo debugs


    If lstPaymentCLN.ListItems.Count Then
        CmdEdit.Visible = True
        cmdAdd.Visible = True
        CmdUpdate.Visible = False
        CmdSave.Visible = False

        ClearFramePV
        'If Trim(lstPaymentCLN.SelectedItem.SubItems(17)) <> "" Then
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
                txtChequeDate = Format(lstPaymentCLN.SelectedItem.SubItems(15), "dd/mm/yyyy")
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
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
   

End Sub

Private Sub lstPaymentCLN_KeyUp(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "lstPaymentCLN_KeyUp"
On Error GoTo debugs


    If lstPaymentCLN.ListItems.Count Then
        CmdEdit.Visible = True
        cmdAdd.Visible = True
        CmdUpdate.Visible = False
        CmdSave.Visible = False

        ClearFramePV
        'If Trim(lstPaymentCLN.SelectedItem.SubItems(17)) <> "" Then
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
                txtChequeDate = Format(lstPaymentCLN.SelectedItem.SubItems(15), "dd/mm/yyyy")
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
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
  

End Sub

Private Sub txtBordxDateFr_LostFocus()
    If IsDate(txtBordxDateFr) Then
    Else
        If txtBordxDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDateFr.SetFocus
        End If
    End If
End Sub

Private Sub txtBordxDateTo_LostFocus()
    If IsDate(txtBordxDateTo) Then
    Else
        If txtBordxDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDateTo.SetFocus
        End If
    End If
End Sub

Private Sub txtChequeAmt_Change()
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
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
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
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
    Case 14
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtDateTime, m_blnDirection(16)
    Case 17
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
    Case 18
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 19
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(19)
    Case 20
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(21)
    Case 22
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(22)
    Case 23
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(23)
    Case 24
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(24)
    Case 25
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(25)
    Case 26
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(26)
    Case 27
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(27)
    Case 28
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(28)
    Case 29
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(29)
    Case 30
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(30)
    Case 31
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(31)
    Case 32
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(32)
    Case 33
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(33)
    Case 34
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(34)
        
    Case 35
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(35)
    Case 36
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(36)
    Case 37
        SortListView lstPaymentCLN, ColumnHeader.Index, ldtString, m_blnDirection(37)
        
    End Select
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
    CboBank.Text = "MB - MAYBANK"
    txtAccNo = ""
    txtCaseNo = ""
    TxtRemark = ""
    txtCLNEmail = ""
    CboPayto = ""
    txtAmtWording = ""
End Sub

Private Sub ClinicHQInfo(tmpClinicCode, tmpCLNName, tmpCLNAdd1, tmpCLNAdd2, tmpCLNAdd3, tmpPCode, tmpCity, tmpState As String)
Const METHOD_NAME = "ClinicHQInfo"
On Error GoTo debugs


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
        objWorksheet.Cells(15 + CDbl(Cnt), "D") = IIf(Len(PVRec!Bordxrefno), PVRec!Bordxrefno, "")
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
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            GoTo ExitFunction
            
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
           
            
End Sub

Private Sub ReadRecord(StrAppend As String, PVRec As ADODB.Recordset)
Const METHOD_NAME = "ReadRecord"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "SP_CISPayment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("PaymentType", adVarChar, adParamInput, 50, "PaymentCLNV2")
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, StrAppend)
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

Private Sub ReadRecordPopup(StrAppend As String, PVRec As ADODB.Recordset)
Const METHOD_NAME = "ReadRecordPopup"
On Error GoTo debugs
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
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub CheckDuplicate()
Dim ForLoop As Integer
Dim ListCount As Integer
Dim tmpcnt As Integer
    tmpcnt = 0
    ForLoop = 1
    ListCount = lstPaymentCLN.ListItems.Count
    For ForLoop = 1 To ListCount
        'With lstPaymentCLN.ListItems(ForLoop)
            If lstPaymentCLN.ListItems(ForLoop).Checked Then
                If Trim(lstPaymentCLN.ListItems(ForLoop).SubItems(2)) = tmpClinicCode And Trim(lstPaymentCLN.ListItems(ForLoop).SubItems(4)) = Trim(tmpInvoiceNo) Then
                    tmpcnt = tmpcnt + 1
                    If tmpcnt > 1 Then
                        DuplicateStatus = True
                        Exit Sub
                    End If
                End If
            End If
        Next
        
End Sub


Private Sub Payment2CLNPopUp(Optional StrAppend As String)
Const METHOD_NAME = "Payment2CLNPopUp"
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
Dim tmpPay2CLN As Double
Dim StrCLNSql As String
Dim ClnRec As New ADODB.Recordset
Dim tmpPVCLNBank
Dim tmpPVCLNAccount As String
Dim tmpCLNRemark As String
Dim tmpCLNBankAcc As String
Dim tmpCLNPName As String
Dim tmpCLNPAdd1 As String
Dim tmpCLNPAdd2 As String
Dim tmpCLNPAdd3 As String
Dim tmpCLNType As String
Dim tmpCLNFFSStatus As String
Dim tmpCLNRegNo As String
Dim tmpCLNMCOProcfeePerc As String
Dim tmpCLNBankName As String
Dim tmpCLNName As String
Dim tmpCLNAdd1  As String
Dim tmpCLNAdd2  As String
Dim tmpCLNAdd3 As String
Dim tmpPCode  As String
Dim tmpCity  As String
Dim tmpState   As String

    Screen.MousePointer = vbHourglass

    Current = 0
    Load frmInvoiceDetPopup
    frmInvoiceDetPopup.lstPaymentCLN.ListItems.Clear
    'Call ClearFramePV
    'sql = " Select * From VIEW_Payment2CLN_Invoice where 0 = 0 "
    Call ReadRecordPopup(StrAppend, PVRec)
    Do
        Do While Not PVRec.EOF
        txtPVStatus = ""
        If Len(PVRec!PVPayCLNStatus) Then
            txtPVStatus = PVRec!PVPayCLNStatus
        End If
            
            With PVRec
            
                StrCLNSql = "SELECT Top 10 * FROM CLN WHERE CLNCode='" & Trim(PVRec!ClinicCode) & "'"
                ClnRec.Open StrCLNSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                Do While Not ClnRec.EOF
                    With ClnRec
                         tmpCLNName = Trim(!CLNName)
                        tmpCLNAdd1 = IIf(Len(!CLNAdd1), !CLNAdd1, "")
                        tmpCLNAdd2 = IIf(Len(!CLNAdd2), !CLNAdd2, "")
                        tmpCLNAdd3 = IIf(Len(!CLNAdd3), !CLNAdd3, "")
                        tmpPCode = Trim(!CLNPostCode)
                        tmpCity = Trim(!CLNCity)
                        tmpState = Trim(!CLNState)
                        'tmpPVCLNBank = !PVCLNBank
                        'tmpPVCLNAccount = !PVCLNAccount
                        tmpCLNRemark = IIf(Len(!CLNRemark), !CLNRemark, "")
                        tmpCLNBankAcc = IIf(Len(!CLNBankAcc), !CLNBankAcc, "")
                        tmpCLNPName = IIf(Len(!CLNPName), !CLNPName, "")
                        tmpCLNPAdd1 = IIf(Len(!CLNPAdd1), !CLNPAdd1, "")
                        tmpCLNPAdd2 = IIf(Len(!CLNPAdd2), !CLNPAdd2, "")
                        tmpCLNPAdd3 = IIf(Len(!CLNPAdd3), !CLNPAdd3, "")
                        tmpCLNType = IIf(Len(!CLNType), !CLNType, "")
                        tmpCLNFFSStatus = IIf(Len(!CLNFFSStatus), !CLNFFSStatus, "")
                        tmpCLNPName = IIf(Len(!CLNPName), !CLNPName, "")
                        tmpCLNRegNo = IIf(Len(!CLNRegNo), !CLNRegNo, "")
                        tmpCLNMCOProcfeePerc = IIf(Len(!CLNMCOProcfeePerc), !CLNMCOProcfeePerc, "")
                        tmpCLNBankName = IIf(Len(!CLNBankName), !CLNBankName, "")
                        tmpCLNBankAcc = IIf(Len(!CLNBankAcc), !CLNBankAcc, "")

                    End With
                    ClnRec.MoveNext
                Loop
                ClnRec.Close
                Set ClnRec = Nothing
            
                Set Record = frmInvoiceDetPopup.lstPaymentCLN.ListItems.Add(, , "")
                Record.SubItems(1) = ""
                If Trim(txtPVStatus) = "F" Then
                    Record.SubItems(1) = "Full"
                ElseIf Trim(txtPVStatus) = "P" Then
                    Record.SubItems(1) = "Partial"
                End If
                Record.SubItems(2) = IIf(Len(PVRec!ClinicCode), PVRec!ClinicCode, "")
                Record.SubItems(3) = IIf(Len(tmpCLNName), tmpCLNName, "")
                Record.SubItems(4) = IIf(Len(PVRec!InvoiceNo), PVRec!InvoiceNo, "")
                If Len(PVRec!InvoiceDate) Then
                    Record.SubItems(5) = Format(PVRec!InvoiceDate, "dd/mm/yyyy")
                End If
                tmpPay2CLN = IIf(IsNumeric(PVRec!CLMPayableByMedix2C), Format(PVRec!CLMPayableByMedix2C, "#,##0.00"), 0)
                Record.SubItems(6) = Format(tmpPay2CLN, "#,##0.00")
                Record.SubItems(7) = IIf(IsNumeric(PVRec!CLNPaidAmt), Format(PVRec!CLNPaidAmt, "#,##0.00"), 0)
                Record.SubItems(8) = CDbl(Record.SubItems(6) - Record.SubItems(7))
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

                'If Len(Trim(PVRec!CLNAccPeriodMth) & Trim(PVRec!CLNAccPeriodYr)) Then
                '    Record.SubItems(17) = Format(PVRec!CLNAccPeriodMth & "/" & PVRec!CLNAccPeriodYr, "mm/yyyy")
                'End If
                If Len(Trim(!PVCLNBank)) Then
                    Record.SubItems(18) = Trim(!PVCLNBank)
                End If
                If Len(Trim(!PVCLNAccount)) Then
                    Record.SubItems(19) = Trim(!PVCLNAccount)
                End If
                If Len(Trim(tmpCLNRemark)) Then
                    Record.SubItems(20) = Trim(tmpCLNRemark)
                End If
                If Len(Trim(tmpCLNBankAcc)) Then
                    Record.SubItems(23) = Trim(tmpCLNBankAcc)
                End If
                
                If Len(Trim(tmpCLNPName)) Then
                    Record.SubItems(24) = Trim(tmpCLNPName)
                End If
                
                If Len(Trim(PVRec!PVWordingOne)) Then
                    Record.SubItems(25) = Trim(PVRec!PVWordingOne)
                End If
                
                If Len(Trim(PVRec!PVWordingTwo)) Then
                    Record.SubItems(25) = Record.SubItems(25) & PVRec!PVWordingTwo
                End If
                                  
                If Len(Trim(tmpCLNPAdd1)) Then
                    Record.SubItems(26) = Trim(tmpCLNPAdd1)
                End If
                
                If Len(Trim(tmpCLNPAdd2)) Then
                    Record.SubItems(27) = Trim(tmpCLNPAdd2)
                End If
                
                If Len(Trim(tmpCLNPAdd3)) Then
                    Record.SubItems(28) = Trim(tmpCLNPAdd3)
                End If
                Record.SubItems(29) = Format(PVRec!CLMFinalApprove, "#,##0.00")
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
    Screen.MousePointer = Default
    
    frmInvoiceDetPopup.Show vbModal
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
   

End Sub



