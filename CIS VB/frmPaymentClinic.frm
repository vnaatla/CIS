VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPaymentClinic 
   Caption         =   "Payment to Clinic"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FramePV 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   0
      TabIndex        =   50
      Top             =   6600
      Width           =   11775
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
         TabIndex        =   29
         Top             =   200
         Width           =   1215
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
         TabIndex        =   31
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
         Left            =   8520
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   200
         Width           =   3105
      End
      Begin MSMask.MaskEdBox txtPVDate 
         Height          =   315
         Left            =   720
         TabIndex        =   30
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
         TabIndex        =   32
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
         TabIndex        =   33
         Top             =   720
         Width           =   1215
      End
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
         ItemData        =   "frmPaymentClinic.frx":0000
         Left            =   5280
         List            =   "frmPaymentClinic.frx":0002
         TabIndex        =   34
         Text            =   "CB - CITIBANK"
         Top             =   160
         Width           =   2535
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
         TabIndex        =   35
         Top             =   440
         Width           =   2535
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
         TabIndex        =   93
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
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
         ItemData        =   "frmPaymentClinic.frx":0004
         Left            =   6840
         List            =   "frmPaymentClinic.frx":0006
         TabIndex        =   37
         Top             =   690
         Width           =   975
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
         ItemData        =   "frmPaymentClinic.frx":0008
         Left            =   5280
         List            =   "frmPaymentClinic.frx":000A
         TabIndex        =   36
         Top             =   690
         Width           =   855
      End
      Begin VB.TextBox txtCaseNo 
         Height          =   285
         Left            =   7320
         TabIndex        =   94
         Top             =   680
         Visible         =   0   'False
         Width           =   495
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
         TabIndex        =   87
         Top             =   480
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
         TabIndex        =   86
         Top             =   240
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
         TabIndex        =   58
         Top             =   240
         Width           =   615
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
         TabIndex        =   57
         Top             =   480
         Width           =   615
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
         TabIndex        =   56
         Top             =   225
         Width           =   975
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
         TabIndex        =   55
         Top             =   720
         Width           =   975
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
         TabIndex        =   54
         Top             =   765
         Width           =   735
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
         TabIndex        =   53
         Top             =   735
         Width           =   615
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
         TabIndex        =   51
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.TextBox txtPVStatus 
      Height          =   285
      Left            =   1920
      TabIndex        =   85
      Top             =   8280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2400
      Left            =   0
      TabIndex        =   63
      Top             =   330
      Width           =   11775
      Begin VB.OptionButton optInvNo 
         Height          =   255
         Left            =   5520
         TabIndex        =   115
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Append"
         Height          =   255
         Left            =   5760
         TabIndex        =   113
         Top             =   1995
         Width           =   975
      End
      Begin VB.CommandButton cmdFileLocationSUN 
         Caption         =   " >"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3840
         TabIndex        =   111
         Top             =   1920
         Width           =   435
      End
      Begin VB.CommandButton cmdUploadDirect 
         Caption         =   "Upload Excel"
         Height          =   375
         Left            =   4320
         TabIndex        =   110
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CheckBox chkAppendRcd 
         Caption         =   "Append Records"
         Height          =   255
         Left            =   8040
         TabIndex        =   109
         Top             =   1920
         Width           =   1695
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
         Left            =   5280
         TabIndex        =   99
         Top             =   1680
         Width           =   375
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
         TabIndex        =   89
         Top             =   840
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
         TabIndex        =   88
         Top             =   600
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
         TabIndex        =   27
         Top             =   2385
         Visible         =   0   'False
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
         TabIndex        =   15
         Top             =   540
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
         TabIndex        =   18
         Top             =   810
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
         TabIndex        =   21
         Top             =   1095
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
         TabIndex        =   24
         Top             =   2400
         Visible         =   0   'False
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
         TabIndex        =   9
         Top             =   1410
         Width           =   375
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
         TabIndex        =   12
         Top             =   270
         Width           =   375
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
         TabIndex        =   0
         Top             =   255
         Width           =   1560
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   270
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtInvDate1 
         Height          =   315
         Left            =   7920
         TabIndex        =   10
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
         TabIndex        =   13
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
         TabIndex        =   16
         Top             =   780
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtPVDate1 
         Height          =   315
         Left            =   7920
         TabIndex        =   19
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
      Begin MSMask.MaskEdBox txtInvDate2 
         Height          =   315
         Left            =   9840
         TabIndex        =   11
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
         TabIndex        =   14
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
         TabIndex        =   17
         Top             =   780
         Width           =   1335
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
         ItemData        =   "frmPaymentClinic.frx":000C
         Left            =   7920
         List            =   "frmPaymentClinic.frx":000E
         TabIndex        =   22
         Top             =   2400
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
         ItemData        =   "frmPaymentClinic.frx":0010
         Left            =   7800
         List            =   "frmPaymentClinic.frx":0012
         TabIndex        =   25
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtPVDate2 
         Height          =   315
         Left            =   9840
         TabIndex        =   20
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
         ItemData        =   "frmPaymentClinic.frx":0014
         Left            =   9720
         List            =   "frmPaymentClinic.frx":0016
         TabIndex        =   23
         Top             =   2400
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
         ItemData        =   "frmPaymentClinic.frx":0018
         Left            =   9720
         List            =   "frmPaymentClinic.frx":001A
         TabIndex        =   26
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
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
         TabIndex        =   3
         Top             =   540
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
         TabIndex        =   4
         Top             =   825
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
         TabIndex        =   5
         Top             =   1110
         Width           =   3615
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
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1395
         Width           =   3615
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
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1650
         Width           =   1575
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
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1650
         Width           =   1575
      End
      Begin VB.TextBox txtSUNnvNoFr 
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
         TabIndex        =   100
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtSUNnvNoTo 
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
         TabIndex        =   101
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox cboEntityType 
         Height          =   315
         ItemData        =   "frmPaymentClinic.frx":001C
         Left            =   7920
         List            =   "frmPaymentClinic.frx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   1580
         Width           =   3255
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6840
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtImportFilenameSUN 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   112
         Top             =   1900
         Width           =   2265
      End
      Begin VB.Label Label30 
         Caption         =   "File Name"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   114
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lblEntityType 
         Caption         =   "Entity Type"
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
         Index           =   2
         Left            =   6360
         TabIndex        =   108
         Top             =   1590
         Width           =   1095
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
         Index           =   1
         Left            =   9450
         TabIndex        =   103
         Top             =   1365
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "SUN Invoice No"
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
         Index           =   1
         Left            =   6360
         TabIndex        =   102
         Top             =   1365
         Width           =   1575
      End
      Begin VB.Label Label38 
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
         Left            =   3240
         TabIndex        =   98
         Top             =   1680
         Width           =   375
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
         Left            =   5325
         TabIndex        =   90
         Top             =   465
         Width           =   255
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
         TabIndex        =   82
         Top             =   600
         Width           =   1215
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
         TabIndex        =   81
         Top             =   885
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
         TabIndex        =   80
         Top             =   1410
         Width           =   1335
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
         TabIndex        =   79
         Top             =   1155
         Width           =   1455
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
         TabIndex        =   78
         Top             =   560
         Width           =   1575
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
         TabIndex        =   77
         Top             =   560
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
         Index           =   0
         Left            =   6360
         TabIndex        =   76
         Top             =   820
         Width           =   1095
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
         Index           =   0
         Left            =   9450
         TabIndex        =   75
         Top             =   820
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
         TabIndex        =   74
         Top             =   1100
         Width           =   855
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
         TabIndex        =   73
         Top             =   1100
         Width           =   375
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
         TabIndex        =   72
         Top             =   260
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
         TabIndex        =   71
         Top             =   260
         Width           =   495
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
         TabIndex        =   70
         Top             =   2400
         Visible         =   0   'False
         Width           =   1575
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
         TabIndex        =   69
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
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
         TabIndex        =   68
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
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
         Index           =   0
         Left            =   6360
         TabIndex        =   67
         Top             =   2385
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   5325
         TabIndex        =   66
         Top             =   1305
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
         TabIndex        =   65
         Top             =   285
         Width           =   375
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
         TabIndex        =   64
         Top             =   285
         Width           =   1215
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
         TabIndex        =   95
         Top             =   1620
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   555
      Left            =   0
      TabIndex        =   59
      Top             =   7680
      Width           =   11775
      Begin VB.TextBox LblAmt 
         Height          =   285
         Left            =   3120
         TabIndex        =   106
         Top             =   210
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   960
         TabIndex        =   105
         Top             =   210
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   104
         Top             =   210
         Visible         =   0   'False
         Width           =   525
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
         Left            =   4440
         TabIndex        =   97
         Top             =   180
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
         Left            =   5280
         TabIndex        =   39
         Top             =   180
         Width           =   855
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
         TabIndex        =   44
         Top             =   180
         Width           =   735
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
         TabIndex        =   45
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
         TabIndex        =   48
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
         TabIndex        =   49
         Top             =   180
         Width           =   615
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
         Left            =   6720
         TabIndex        =   42
         Top             =   180
         Width           =   735
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
         TabIndex        =   43
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
         TabIndex        =   47
         Top             =   180
         Width           =   975
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
         TabIndex        =   46
         Top             =   180
         Width           =   735
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
         Left            =   6120
         TabIndex        =   40
         Top             =   180
         Width           =   615
      End
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
         Left            =   6120
         TabIndex        =   41
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label29 
         Caption         =   "Amt"
         Height          =   255
         Left            =   2640
         TabIndex        =   62
         Top             =   240
         Width           =   375
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
         TabIndex        =   61
         Top             =   240
         Width           =   255
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
         TabIndex        =   60
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstPaymentCLN 
      Height          =   3855
      Left            =   0
      TabIndex        =   28
      Top             =   2760
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
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   36
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Pay Status"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CLNCode"
         Object.Width           =   1587
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
         Object.Width           =   1764
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
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "Company Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "Company Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "SUN Invoice No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "SUN Post Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "FFS Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "Co-Pay Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "Bordx Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "Final Approve amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "Total Actual amt"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CheckBox CheckAll 
      Caption         =   "Check All"
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
      Left            =   50
      TabIndex        =   96
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "   Payment to Clinic"
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
      TabIndex        =   92
      Top             =   0
      Width           =   11805
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
      TabIndex        =   91
      Top             =   8040
      Width           =   255
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
      TabIndex        =   84
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
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   2520
      TabIndex        =   83
      Top             =   8040
      Width           =   2295
   End
End
Attribute VB_Name = "frmPaymentClinic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer
Public intMBMAccount As Integer
Public intExit As Integer
Private m_blnDirection(1 To 36) As Boolean

Dim EditFlag As Boolean
Dim ScanCode As String
Dim Check As Boolean
Dim listloop1 As Integer
Dim RecordSelected As Integer

Dim PVNumber As String

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

Private Sub CheckAll_Click()
Dim AA As Integer

    If lstPaymentCLN.ListItems.Count > 0 Then
        If CheckAll.Value = 1 Then
            For AA = 1 To lstPaymentCLN.ListItems.Count
                If Trim(lstPaymentCLN.ListItems(AA).SubItems(1)) = "" Then
                    lstPaymentCLN.ListItems(AA).Checked = True
                    lstPaymentCLN.ListItems(AA).SubItems(1) = "Full"
                    lstPaymentCLN.ListItems(AA).SubItems(10) = IIf(Trim(lstPaymentCLN.ListItems(AA).SubItems(8)) = "", CDbl(lstPaymentCLN.ListItems(AA).SubItems(7)), CDbl(lstPaymentCLN.ListItems(AA).SubItems(9)))
                    LblAmt = LblAmt + CDbl(Format(lstPaymentCLN.ListItems(AA).SubItems(10), "#,##0.00"))
                End If
            Next AA
        Else
            For AA = 1 To lstPaymentCLN.ListItems.Count
                If lstPaymentCLN.ListItems(AA).Checked = True Then
                    lstPaymentCLN.ListItems(AA).Checked = False
                    LblAmt = LblAmt - CDbl(Format(lstPaymentCLN.ListItems(AA).SubItems(10), "#,##0.00"))
                    lstPaymentCLN.ListItems(AA).SubItems(10) = ""
                    lstPaymentCLN.ListItems(AA).SubItems(1) = ""
                End If
            Next AA
        End If
    End If
End Sub

Private Sub ChkClinic_Click()
Dim CLNRec As New ADODB.Recordset

    CboCLN.Clear
    If ChkClinic.Value = 1 Then
    
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        CLNRec.Open "Select CLNCode, CLNName from CLN WHERE CLNPanel = 'Y' OR CLNPanel = 'C'", medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do Until CLNRec.EOF
            If Len(CLNRec!CLNName) Then
                CboCLN.AddItem CLNRec!CLNCode & " - " & Trim(CLNRec!CLNName)
            End If
            CLNRec.MoveNext
        Loop
        CLNRec.Close
        Set CLNRec = Nothing
        
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
        
    End If

End Sub

Private Sub ChkGrpComp_Click()
Dim Company As New ADODB.Recordset
    
  '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        Company.Open "Select GRPCoName from GRPCompany", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not Company.EOF
            CboPayor.AddItem Trim(Company!GrpCoName)
            Company.MoveNext
        Loop
        Company.Close
        Set Company = Nothing
   ' MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing
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
        CmdSave.Visible = True
        cmdAdd.Visible = False
        FramePV.Enabled = True
        CboBank.Enabled = True
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

Private Sub cmdFileLocationSUN_Click()
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.InitDir = "C:\"
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Dat Files (*.Dat)|*.Dat"
    CommonDialog1.ShowOpen
    txtImportFilenameSUN = CommonDialog1.FileTitle
    Exit Sub
  
ErrHandler:
  Exit Sub

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
        Call PreviewPV
    End If
End Sub

Private Sub CmdPrintPV_Click()
Const METHOD_NAME = "CmdPrintPV_Click"
On Error GoTo debugs


    Dim REceipt As New ADODB.Recordset
    Dim ReceiptNo As String
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim i As Integer
    Dim strSql As String
    Dim PV As New ADODB.Recordset
    Dim tmpClinicCode As String
    Dim tmpCLNName As String
    Dim tmpCLNAdd1 As String
    Dim tmpCLNAdd2 As String
    Dim tmpCLNAdd3 As String
    Dim tmpPCode As String
    Dim tmpCity As String
    Dim tmpState As String
    Dim Cnt As Integer
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
        If objExcel.Visible = False Then
            objExcel.Visible = True
        End If
        Cnt = 0
        strSql = "select CASNumber, BordxRefNo, BillNo, BillDate, PVCLNNo,PVCLNChequeNo, PVCLNDate, CLMTotalActual, CLMMCOProcFee, " & _
                "ClinicGrpCode, ClinicCode, CLNName, CLNAdd1, CLNAdd2,CLNAdd3, CLNPostCode, CLNCity, CLNState, CLNPaidAmt " & _
                "from View_Payment2CLN where PVCLNNo = '" & TxtPVNo & "'"
        PV.Open strSql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
        'Set objExcel = New excel.Application
        If PV.EOF = False Then
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\PV-Clinic.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        Do While Not PV.EOF
            Cnt = Cnt + 1
            tmpClinicCode = Trim(PV!ClinicCode)
            tmpCLNName = IIf(Len(Trim(PV!CLNName)), Trim(PV!CLNName), "")
            tmpCLNAdd1 = IIf(Len(Trim(PV!CLNAdd1)), Trim(PV!CLNAdd1), "")
            tmpCLNAdd2 = IIf(Len(Trim(PV!CLNAdd2)), Trim(PV!CLNAdd2), "")
            tmpCLNAdd3 = IIf(Len(Trim(PV!CLNAdd3)), Trim(PV!CLNAdd3), "")
            tmpPCode = IIf(Len(Trim(PV!CLNPostCode)), Trim(PV!CLNPostCode), "")
            tmpCity = IIf(Len(Trim(PV!CLNCity)), Trim(PV!CLNCity), "")
            tmpState = IIf(Len(Trim(PV!CLNState)), Trim(PV!CLNState), "")
          
            objExcel.WindowState = xlMaximized
            objWorksheet.Cells(8, "C") = tmpCLNName & " (" & tmpClinicCode & ")"
            objWorksheet.Cells(9, "C") = tmpCLNAdd1
            objWorksheet.Cells(10, "C") = tmpCLNAdd2 & "," & tmpCLNAdd3
            objWorksheet.Cells(11, "C") = Trim(tmpPCode & " " & Trim(tmpCity) & " " & Trim(tmpState))
            objWorksheet.Cells(8, "H") = TxtPVNo
            objWorksheet.Cells(9, "H") = PV!PVCLNChequeNo
            objWorksheet.Cells(10, "H") = PV!PVCLNDate
            objWorksheet.Cells(37, "C") = Logon
            objWorksheet.Cells(15, "D") = "Case No."
            objWorksheet.Cells(15, "E") = "Receipt No."
            ReceiptNo = ""
            REceipt.Open "SELECT ReceiptNo from REceiptFrPAY where CASNumber ='" & PV!CasNumber & "'", MediConnCISPayment, adOpenKeyset, adLockOptimistic
            If REceipt.RecordCount <> 0 Then
                ReceiptNo = REceipt!ReceiptNo
            Else
                REceipt.Close
                Set REceipt = Nothing
                REceipt.Open "SELECT ReceiptNo from REceiptFrPAY where CASNumber ='" & PV!CasNumber & "'", MediConnCISPayment, adOpenKeyset, adLockOptimistic
                If REceipt.RecordCount <> 0 Then
                    ReceiptNo = REceipt!ReceiptNo
                End If
            End If
            REceipt.Close
            Set REceipt = Nothing
            
            objWorksheet.Cells(15 + CDbl(Cnt), "A") = Cnt
            objWorksheet.Cells(15 + CDbl(Cnt), "B") = PV!BillNo
            objWorksheet.Cells(15 + CDbl(Cnt), "C") = PV!BillDate
            objWorksheet.Cells(15 + CDbl(Cnt), "D") = PV!CasNumber
            objWorksheet.Cells(15 + CDbl(Cnt), "E") = ReceiptNo
            objWorksheet.Cells(15 + CDbl(Cnt), "F") = PV!CLMTotalActual
            objWorksheet.Cells(15 + CDbl(Cnt), "G") = IIf(IsNumeric(PV!CLMMCOProcFee), PV!CLMMCOProcFee, 0)
            objWorksheet.Cells(15 + CDbl(Cnt), "H") = PV!CLNPaidAmt
            PV.MoveNext
            
            Loop
            PV.Close
            Set PV = Nothing
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
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            GoTo ExitFunction
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdUploadDirect_Click()
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
    
Dim xl As New Excel.Application
Dim xlsheet As Excel.Worksheet
Dim xlwbook As Excel.Workbook
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim tmpPathName As String

    Screen.MousePointer = vbHourglass
        If Check1.Value = 0 Then
            lstPaymentCLN.ListItems.Clear
        End If
        'ChkScan.Value = 0
    If optInvNo.Value <> False Then
        If txtImportFilenameSUN <> "" Then
                Val1 = 2
                error = False
                'cmpUpload.Enabled = False
                FileExists = 0
            
                lSize = -1
                lSize = FileLen(CommonDialog1.Filename)
                
                If lSize <= 0 Then
                    FileExists = 0
                ElseIf lSize > 0 Then
                    FileExists = 1
                End If
                
                txtImportFilenameSUN = CommonDialog1.Filename
                If FileExists = 1 Then
                    Set xlwbook = xl.Workbooks.Open(CommonDialog1.Filename)
                    Set xlsheet = xlwbook.Sheets.Item(1)
                    
                    Do Until Val1 = 60000
                        If xlsheet.Cells(Val1, 1) <> "" Then
                            error = False
                            TransRef = xlsheet.Cells(Val1, 1)
                            
                            Call SearchRecordExcel(TransRef)
                        
                        Else
                            Screen.MousePointer = vbDefault
                            'cmpUpload.Enabled = True
                            'If ListView2.ListItems.Count = 0 Then
                            '    MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
                            '    CmdPrint.Enabled = True
                            '    CmdClear.Enabled = True
                            '    cmdExit.Enabled = True
                            'End If
                            Exit Sub
                        End If
                        Val1 = Val1 + 1
                        
                    Loop
                Else
                    Screen.MousePointer = vbDefault
                    'cmpUpload.Enabled = True
                    MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
                
                End If
        Else
            Screen.MousePointer = vbDefault
           ' cmpUpload.Enabled = True
            MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
    
        End If
    
        
     Else
            MsgBox "Please select Invoice Number!", vbOKOnly + vbCritical, DialogBoxTitle
            Screen.MousePointer = vbDefault
        End If

    'Screen.MousePointer = vbDefault


End Sub

Private Function SearchRecordExcel(SearchStrAppend As String)
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim strSql As String
Dim strAppend As String
Dim sql As String
Dim GrpCompanyName As String
    
    
    strAppend = ""
    GrpCompanyName = ""
    
    
    If Len(SearchStrAppend) Then
        strAppend = strAppend & " Where 0=0 and CISCas.InvoiceNo = '" & Trim(SearchStrAppend) & "'"
    End If
    
    
    Call Payment2CLNListing(strAppend)

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Function

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

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelAnalysis1(lstPaymentCLN)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEdit_Click()
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
    CmdUpdate.Visible = False
    FramePV.Enabled = False
    If ItemChk = False And Len(Trim(TxtPVNo)) Then
        CmdUpdate.Visible = True
        CmdEdit.Visible = False
        FramePV.Enabled = True
        EditFlag = True
        If Len(Trim(lstPaymentCLN.SelectedItem.SubItems(1))) Then
            FramePV.Enabled = True
        End If
        txtPVDate.SetFocus
    End If
End Sub

Public Sub cmdSearch_Click()
Const METHOD_NAME = "CmdSearch_Click"
On Error GoTo debugs

Dim listrecord As Integer
        Screen.MousePointer = vbHourglass
            LblAmt = 0
            EditFlag = False
            If ChkScan.Value = 1 Then
                If TxtWSNo = "" Then
                    MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
                    TxtWSNo.SetFocus
                Else
                   'If ChkScan.Value = 0 Then
                   '     If chkAppendRcd.Value = 0 Then
                   '         lstPaymentCLN.ListItems.Clear
                   '     End If
                   ' End If

                    listrecord = lstPaymentCLN.ListItems.Count
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
                If ChkScan.Value = 0 Then
                    If chkAppendRcd.Value = 0 Then
                        lstPaymentCLN.ListItems.Clear
                    End If
                End If
                Call SearchRecord
            End If
    Screen.MousePointer = Default
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    Check = False
    TxtWSNo.SetFocus
    CmdExit.Enabled = True
End Sub

Private Sub CmdUpdate_Click()
Const METHOD_NAME = "CmdUpdate_Click"
On Error GoTo debugs


Dim RecordSaved
Dim TEmpStatus As String
Dim TempPayAmt As String
Dim TempCaseNo As String
Dim TempIndex As String
Dim ZZ As Integer
Dim ItemChk As Boolean
Dim PVSql As String
Dim PayRec As New ADODB.Recordset
Dim StrPaySql As String
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
            StrPaySql = "SELECT PVCLNIndex, CASNumber, PVPAYCLNStatus, CLNPaidAmt " & _
                        " From Payment2CLN WHERE PVCLNNo='" & Trim(TxtPVNo) & "'"
            PayRec.Open StrPaySql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
            Do While Not PayRec.EOF
                With PayRec
                    TempCaseNo = !CasNumber
                    TEmpStatus = !PVPayCLNStatus
                    TempPayAmt = !CLNPaidAmt
                    TempIndex = !PVCLNIndex
                    If Trim(txtPayIndex) = Trim(PayRec!PVCLNIndex) Then
                        TempCaseNo = txtCaseNo
                        TEmpStatus = Mid(lstPaymentCLN.SelectedItem.SubItems(1), 1, 1)
                        TempPayAmt = Trim(lstPaymentCLN.SelectedItem.SubItems(8))
                        TempIndex = txtPayIndex
                    End If
                    Call UpdateRecord(TempCaseNo, TEmpStatus, TempPayAmt, TempIndex)
                End With
                PayRec.MoveNext
            Loop
            CmdSave.Enabled = True
            EditFlag = False
            LblAmt = 0
            PVSql = " Where 0=0 and Payment2CLN.PVCLNNo = '" & Trim(TxtPVNo) & "'"
            Call ClearFramePV
            lstPaymentCLN.ListItems.Clear
            Call Payment2CLNListing(PVSql)
        End If
        
        
    End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "frmPaymentClinic")
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
    CmdEdit.Enabled = False
    cboEntityType.ListIndex = 0
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    CmdSave.Visible = False
    cmdAdd.Visible = True
    CmdUpdate.Visible = False
    CmdEdit.Visible = True
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
    Unload Me
End Sub

'************************Start ComboListing**************************
Private Sub ComboListing()
Dim monthstart
Dim yearstart
Dim HLT As New ADODB.Recordset
Dim BankRec As New ADODB.Recordset

    
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
        
End Sub
'************************Start ComboListing**************************

Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim strSql As String
Dim strAppend As String
Dim sql As String
Dim GrpCompanyName As String
    
    
    strAppend = ""
    GrpCompanyName = ""
    
    strAppend = strAppend & " Where ClaimInvoice.INVPayTo = 'C' AND CISCas.CASStatus='O' "
    
    If Len(Trim(CboPayor)) Then
         GrpCompanyName = Replace(CboPayor, "'", "''")
         strAppend = strAppend & " AND MBMOthers.MBMGroupCompany = '" & Trim(GrpCompanyName) & "'"
    End If
    If Len(Trim(CboCLN)) Then
        strAppend = strAppend & " AND CISCas.ClinicCode = '" & Trim(Left(CboCLN, IIf(InStr(1, CboCLN, "-") - 1 < 0, Len(CboCLN), InStr(1, CboCLN, "-") - 1))) & "'"
    End If
    
    If Mid(CboPayment, 1, 1) = 1 Then
        strAppend = strAppend & " and Payment2CLN.PVPayCLNStatus = 'F'"
    End If
 
    If Mid(CboPayment, 1, 1) = 2 Then
        strAppend = strAppend & " and COALESCE(Payment2CLN.PVPayCLNStatus , '')='' "
    End If
 
    If Mid(CboPayment, 1, 1) = 3 Then
        strAppend = strAppend & " and Payment2CLN.PVPayCLNStatus = 'P' "
    End If
 
    If Mid(CboPayment, 1, 1) = 4 Then
        strAppend = strAppend & " and (Payment2CLN.PVPayCLNStatus = 'P' or Payment2CLN.PVPayCLNStatus = 'F') "
    End If
            
    If ChkScan.Value = 1 Then
        strAppend = strAppend & " AND CISCas.CASNumber = '" & Trim(TxtWSNo) & "'"
    Else
        If Len(TxtWSNo) Then
            strAppend = strAppend & " AND CISCas.CASNumber >= '" & Trim(TxtWSNo) & "'"
        End If
    
        If Len(TxtWSNoTo) Then
            strAppend = strAppend & " AND CISCas.CASNumber <= '" & Trim(TxtWSNoTo) & "'"
        End If
    End If
    
    If Len(TxtInvNo) Then
        strAppend = strAppend & " and CISCas.InvoiceNo = '" & Trim(TxtInvNo) & "'"
    End If
    
    If ChkInvNo.Value = 1 Then
        sql = ""
        sql = "AND COALESCE(CISCas.InvoiceNo, '')='' " & ""
        strAppend = strAppend & sql
    End If
    
    If (txtInvDate1) <> "__/__/____" And IsDate(txtInvDate1) = True _
        And (txtInvDate2) <> "__/__/____" And IsDate(txtInvDate2) = True Then
        sql = ""
        sql = " AND CISCas.InvoiceDate >= '" & Format(Trim(txtInvDate1), "mm/dd/yyyy") & _
              "' And  CISCas.InvoiceDate <= '" & Format(Trim(txtInvDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend & sql
    End If
    
    If Len(Trim(txtInvDate1)) > 0 And IsDate(txtInvDate1) = True Then
        strAppend = strAppend & " And CISCas.InvoiceDate >= '" & Format(txtInvDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtInvDate2)) > 0 And IsDate(txtInvDate2) = True Then
        strAppend = strAppend & " And CISCas.InvoiceDate <= '" & Format(txtInvDate2, "mm/dd/yyyy") & "'"
    End If
    If ChkBatchNo.Value = 1 Then
        sql = ""
        sql = "AND COALESCE(CISCas.OPBatchNo, '')='' " & ""
        strAppend = strAppend & sql
    End If
    
    If Len(Trim(txtOPBatchNo1)) Then
        strAppend = strAppend & " AND CISCas.OPBatchNo >= '" & Trim(txtOPBatchNo1) & "'"
    End If
    If Len(Trim(txtOPBatchNo2)) Then
        strAppend = strAppend & " AND CISCas.OPBatchNo <= '" & Trim(txtOPBatchNo2) & "'"
    End If
    
    If ChkInvDate.Value = 1 Then
        sql = ""
        sql = "AND COALESCE(CISCas.InvoiceDate, '')='' " & ""
        strAppend = strAppend & sql
    End If
    
    If Len(txtInvAmt1) Then
        strAppend = strAppend & " and CISCas.InvoiceStmtAmt >= '" & Trim(txtInvAmt1) & "'"
    End If
        
    If Len(TxtInvAmt2) Then
        strAppend = strAppend & " and CISCas.InvoiceStmtAmt <= '" & Trim(TxtInvAmt2) & "'"
    End If
    
    If chkInvAmt.Value = 1 Then
        sql = ""
        sql = "AND COALESCE(CISCas.InvoiceStmtAmt , '')='' " & ""
        strAppend = strAppend & sql
    End If
    
    If Len(txtPVNo1) Then
        strAppend = strAppend & " and Payment2CLN.PVCLNNo >= '" & Trim(txtPVNo1) & "'"
    End If
        
    If Len(txtPVNo2) Then
        strAppend = strAppend & " and Payment2CLN.PVCLNNo <= '" & Trim(txtPVNo2) & "'"
    End If
    
    If chkPVNo.Value = 1 Then
        sql = ""
        sql = "AND COALESCE(Payment2CLN.PVCLNNo ,'')='' " & ""
        strAppend = strAppend & sql
    End If
    
    If (txtPVDate1) <> "__/__/____" And IsDate(txtPVDate1) = True _
        And (txtPVDate2) <> "__/__/____" And IsDate(txtPVDate2) = True Then
        sql = ""
        sql = " AND Payment2CLN.PVCLNDate >= '" & Format(Trim(txtPVDate1), "mm/dd/yyyy") & _
              "' And Payment2CLN.PVCLNDate <= '" & Format(Trim(txtPVDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend & sql
    End If
    
    If Len(Trim(txtPVDate1)) > 0 And IsDate(txtPVDate1) = True Then
        strAppend = strAppend & " And Payment2CLN.PVCLNDate >= '" & Format(txtPVDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPVDate2)) > 0 And IsDate(txtPVDate2) = True Then
        strAppend = strAppend & " And Payment2CLN.PVCLNDate <= '" & Format(txtPVDate2, "mm/dd/yyyy") & "'"
    End If
    
    If chkPVDate.Value = 1 Then
        sql = ""
        sql = "AND COALESCE(Payment2CLN.PVCLNDate, '')='' " & ""
        strAppend = strAppend & sql
    End If
        
    If Len(cboACPeriodMthFr) Then
        strAppend = strAppend & " and Payment2CLN.CLNAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
        
    If Len(cboACPeriodMthTo) Then
        strAppend = strAppend & " and Payment2CLN.CLNAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If chkPeriodMth.Value = 1 Then
        sql = ""
        sql = " AND COALESCE(Payment2CLN.CLNAccPeriodMth,'')='' " & ""
        strAppend = strAppend & sql
    End If
    
    If Len(cboACPeriodYRTo) Then
        strAppend = strAppend & " and Payment2CLN.CLNAccPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
    End If
        
    If Len(cboACPeriodYRTo) Then
        strAppend = strAppend & " and Payment2CLN.CLNAccPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
    
    If ChkPeriodYr.Value = 1 Then
        sql = ""
        sql = " AND COALESCE(Payment2CLN.CLNAccPeriodYr,'')='' " & ""
        strAppend = strAppend & sql
    End If
    
    If Len(txtSUNnvNoFr) Then
        strAppend = strAppend & " and SUNPosting.CLNInvoiceBatch >= '" & Trim(txtSUNnvNoFr) & "'"
    End If
    If Len(txtSUNnvNoTo) Then
        strAppend = strAppend & " and SUNPosting.CLNInvoiceBatch <= '" & Trim(txtSUNnvNoTo) & "'"
    End If
    
     If cboEntityType.ListIndex = 0 Then
        strAppend = strAppend + " AND Substring(CISCas.CasNumber, 1, 2) <>'SJ' "
    ElseIf cboEntityType.ListIndex = 1 Then
        strAppend = strAppend + " AND Substring(CISCas.CasNumber, 1, 2) ='SJ' "
    End If
    
    Call Payment2CLNListing(strAppend)

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
Const METHOD_NAME = "CboCLN_KeyPress"
On Error GoTo debugs

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
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CboPayment_Change()
    If InStr(CboPayment.Text, "null") Then
       CboPayment.Enabled = False
    End If
End Sub

Private Sub CboPayment_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboPayment_KeyPress"
On Error GoTo debugs

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

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub Payment2CLNListing(Optional strAppend As String)
Const METHOD_NAME = "Payment2CLNListing"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim strsqlPay As String
Dim PAY As New ADODB.Recordset
Dim INVAmt As Variant
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
Dim Check As Boolean
'Dim RECFull As Boolean
'Dim RECPartial As String
Dim CasNumber As String
Dim WSVersion As String
Dim PatientDB As New ADODB.Recordset
Dim PatientSql As String
Dim tmpNonGPFee As Double
Dim tmpPay2CLN As Double
Dim tmpPaidAmt As Double
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
Total = 0
Current = 0
strAppendCase = "26"

    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimBordxModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 3000
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
'    sql = " Select * From VIEW_Payment2CLN where 0 = 0 "
'
'    If Len(Trim(strAppend)) Then
'        sql = sql + strAppend
'    End If
'
'    rst2.Open sql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
    Do
        Do Until rst2.EOF
        Total = rst2.RecordCount
        lblTotal = Total
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
                Record.SubItems(3) = IIf(Len(rst2!CLNName), rst2!CLNName, "")
                Record.SubItems(4) = IIf(Len(rst2!InvoiceNo), rst2!InvoiceNo, "")
                If Len(rst2!InvoiceDate) Then
                    Record.SubItems(5) = Format(rst2!InvoiceDate, "dd/mm/yyyy")
                End If
                Record.SubItems(6) = IIf(Len(rst2!BillNo), rst2!BillNo, "")
                tmpNonGPFee = IIf(IsNumeric(rst2!ClmNonGPFee), rst2!ClmNonGPFee, 0)
                tmpPay2CLN = IIf(IsNumeric(rst2!CLMPayableByMedix2c), rst2!CLMPayableByMedix2c, 0)
                Record.SubItems(7) = Format(tmpPay2CLN + tmpNonGPFee, "#,##0.00")
                Record.SubItems(8) = IIf(IsNumeric(rst2!CLNPaidAmt), Format(rst2!CLNPaidAmt, "#,##0.00"), 0)
                
'                If rst2!coverid = "00" Then
'                    PatientSql = "Select MBMName as PatientName from MBM where MBMNumber='" & rst2!MBMNumber & "'"
'                Else
'                    PatientSql = "Select MBMCName as PatientName from MBMCoveredPersons where MBMCNumber='" & rst2!MBMNumber & "' AND MBMCCoverID ='" & rst2!coverid & "'"
'                End If
'                PatientDB.Open PatientSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly

                If Len(rst2!PatientName) Then
                    Record.SubItems(11) = Trim(rst2!PatientName)
                End If
'                PatientDB.Close
'                Set PatientDB = Nothing
                
                If Len(rst2!CasNumber) Then
                    Record.SubItems(12) = rst2!CasNumber
                End If
                CasNumber = rst2!CasNumber
                'WSVersion = rst2!clmversion
                If Len(rst2!Bordxrefno) Then
                    Record.SubItems(13) = Trim(rst2!Bordxrefno)
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
                If Len(Trim(rst2!GRPCoCode)) Then
                    Record.SubItems(27) = Trim(rst2!GRPCoCode)
                End If
                If Len(Trim(rst2!MBMGroupCompany)) Then
                    Record.SubItems(28) = Trim(rst2!MBMGroupCompany)
                End If
                
                If Len(Trim(rst2!SUNInvNo)) Then
                    Record.SubItems(29) = Trim(rst2!SUNInvNo)
                End If
                If Len(Trim(rst2!SUNPostDate)) Then
                    Record.SubItems(30) = Trim(rst2!SUNPostDate)
                End If
                'New column added 19-10-2018
                If Len(Trim(rst2!CLMMCOProcFee)) Then
                    Record.SubItems(31) = Trim(rst2!CLMMCOProcFee)
                End If
                If Len(Trim(rst2!CLMCoPayCLNColl)) Then
                    Record.SubItems(32) = Trim(rst2!CLMCoPayCLNColl)
                End If
                If Len(Trim(rst2!CLMBordxAmt)) Then
                    Record.SubItems(33) = Trim(rst2!CLMBordxAmt)
                End If
                If Len(Trim(rst2!CLMFinalApprove)) Then
                    Record.SubItems(34) = Trim(rst2!CLMFinalApprove)
                End If
                If Len(Trim(rst2!CLMTotalActual)) Then
                    Record.SubItems(35) = Trim(rst2!CLMTotalActual)
                End If
                
'                strsqlPay = "Select * from Payment2CLN where CASNumber = '" & Trim(CasNumber) & "'"
'                PAY.Open strsqlPay, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
                INVAmt = 0
                If Trim(txtPVStatus) = "P" Then
'                    Do While Not PAY.EOF
                        INVAmt = INVAmt + rst2!CLNPaidAmt
'                    PAY.MoveNext
'                    Loop
                    If Len(INVAmt) Then
                        Record.SubItems(8) = INVAmt
                        Record.SubItems(9) = rst2!CLMPayableByMedix2c - INVAmt
                    End If
                ElseIf Trim(txtPVStatus) = "" Then
                    If Len(rst2!CLMPayableByMedix2c) Then
                        Record.SubItems(9) = Record.SubItems(7) - Record.SubItems(8)
                    End If
                End If
                Current = Current + 1
                lblCurrent = Current
            End With
'            PAY.Close
'            Set PAY = Nothing

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

Private Sub CmdSave_Click()
Const METHOD_NAME = "CmdSave_Click"
On Error GoTo debugs


Dim PaymentAmt As String
Dim i As Integer
Dim update
Dim ItemChecked As Boolean
Dim PymtIndex As String
Dim ZZ As String
Dim TempBank As String
Dim PVSql As String
Dim BankRec As New ADODB.Recordset
Dim BankSql As String
Dim TmpPVCode As String
Dim tmpCase As String
Dim tmpStatus As String
Dim tmpIndex As String
Dim tmpNonGpAmt As String
    
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

    For i = 1 To lstPaymentCLN.ListItems.Count
        If lstPaymentCLN.ListItems(i).Checked = True Then
            ItemChecked = True
        Exit For
        End If
    Next i
    
    If IsDate(txtPVDate) = False Then
        MsgBox "Invalid PV Date!", vbCritical
        txtPVDate.SetFocus
'    ElseIf txtPVDate = "__/__/____" Then
'       MsgBox "Please Enter PV Date! ", vbCritical
'        txtPVDate.SetFocus
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
    ElseIf ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    Else
        update = MsgBox("Do you want to update the record ? ", vbYesNo + vbExclamation)
        If update = vbYes Then
            PVNumber = GenerateSPVNo(TempBank, TmpPVCode)
            TxtPVNo = ""
            TxtPVNo = PVNumber
            For i = 1 To lstPaymentCLN.ListItems.Count
                 If lstPaymentCLN.ListItems(i).Checked = True Then
                    tmpCase = ""
                    tmpStatus = ""
                    PaymentAmt = ""
                    tmpIndex = ""
                    tmpCase = Trim(lstPaymentCLN.ListItems(i).SubItems(12))
                    tmpStatus = Trim(Mid(lstPaymentCLN.ListItems(i).SubItems(1), 1, 1))
                    
                    'tmpNonGpFee = IIf(IsNumeric(UPdRec!CLMNonGpFee), UPdRec!CLMNonGpFee, 0)
                    PaymentAmt = Trim(lstPaymentCLN.ListItems(i).SubItems(10))
                    tmpIndex = Trim(lstPaymentCLN.ListItems(i).SubItems(17))
                    Call UpdateRecord(tmpCase, tmpStatus, PaymentAmt, tmpIndex)
                End If
            Next i
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            lstPaymentCLN.ListItems.Clear
            LblAmt = 0
            PVSql = " Where 0=0 and Payment2CLN.PVCLNNo = '" & Trim(TxtPVNo) & "'"
            Call ClearFramePV
            lstPaymentCLN.ListItems.Clear
            Call Payment2CLNListing(PVSql)
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub UpdateRecord(caseNo As String, status As String, PaymentAmt As Variant, PymtIndex As String)
Const METHOD_NAME = "UpdateRecord"
On Error GoTo debugs

Dim PAY As New ADODB.Recordset
Dim strsqlPay As String
Dim ZZ As String
Dim TempBank As String

    ZZ = InStr(1, CboBank, "-")
    TempBank = Trim(Left(CboBank, IIf(ZZ - 1 <= 0, Len(Trim(CboBank)), ZZ - 1)))

    strsqlPay = " Select PVCLNIndex, * from Payment2CLN " & _
                " Where PVCLNIndex = '" & Trim(PymtIndex) & "'"
    PAY.Open strsqlPay, MediConnCISPayment, adOpenKeyset, adLockOptimistic
    
    If PAY.RecordCount > 0 Then
        If EditFlag = False Then
            If Trim(PAY!PVPayCLNStatus) = "P" Then
                PAY!PVPayCLNStatus = Mid(status, 1, 1)
                PAY.update
                PAY.AddNew
            End If
        End If
    Else
        PAY.AddNew
    End If
    With PAY
        !CasNumber = Trim(caseNo)
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
    PAY.update
    
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

    CmdEdit.Enabled = False
    If lstPaymentCLN.ListItems.Count Then
        If Trim(lstPaymentCLN.SelectedItem.SubItems(17)) <> "" Then
            cmdAdd.Enabled = False
            FramePV.Enabled = False
            CboBank.Enabled = False
            CmdEdit.Enabled = True
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
            txtChequeDate = Format(lstPaymentCLN.SelectedItem.SubItems(20), "dd/MM/YYYY")
                'txtChequeDate = lstPaymentCLN.SelectedItem.SubItems(20)
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
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub
Private Sub lstPaymentCLN_DblClick()
    If lstPaymentCLN.ListItems.Count > 0 Then
        If Trim(lstPaymentCLN.SelectedItem.SubItems(12)) = "" Then
            MsgBox "Please select Correct Row", vbCritical
        Else
            frmInvoiceAdj.txtCaseNo = Trim(lstPaymentCLN.SelectedItem.SubItems(12))
            Call frmInvoiceAdj.CmdGo_Click
            Load frmInvoiceAdj
            frmInvoiceAdj.Show
        End If
    End If
End Sub

Private Sub lstPaymentCLN_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    CmdSave.Enabled = True
    EditFlag = False
    Call ProcessCheckedItem(Item)
End Sub

Private Function ProcessCheckedItem(Item As ListItem) ', Optional Partial As Boolean)
    If Item.Checked = True Then
        If Mid(Item.SubItems(1), 1, 1) = "F" Then
            MsgBox "Payment has been made!", vbCritical
            Item.Checked = False
        Else
            Item.SubItems(10) = IIf(Trim(Item.SubItems(8)) = "", CDbl(Item.SubItems(7)), CDbl(Item.SubItems(9)))
            Item.SubItems(1) = "Full"
            LblAmt = LblAmt + CDbl(Format(Item.SubItems(10), "#,##0.00"))
        End If
    Else
        LblAmt = LblAmt - CDbl(Format(Item.SubItems(10), "#,##0.00"))
        Item.SubItems(10) = ""
        Item.SubItems(1) = ""
    End If
End Function

Private Sub TxtACNo_lostfocus()
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

Private Sub txtInvNo_LostFocus()
    TxtInvNo = ChkStr(TxtInvNo)
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
    Case 23
    SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(23)
    Case 24
    SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(24)
    Case 25
    SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
    Case 26
    SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
    Case 27
    SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
     Case 28
    SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
     Case 29
    SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(29)
     Case 30
    SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(30)
     Case 31
    SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(31)
     Case 32
    SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(32)
     Case 33
    SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(33)
     Case 34
    SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(34)
     Case 35
    SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(35)
     Case 36
    SortListView lstPaymentCLN, ColumnHeader.Index, ldtNumber, m_blnDirection(36)
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

Private Sub TxtWSNo_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtWSNo_LostFocus()
    TxtWSNo = ChkStr(TxtWSNo)
    CmdSearch.Default = False
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
    CboBank.Text = "CB - CITIBANK"
    txtAcNo = ""
    txtCaseNo = ""
    TxtRemark = ""
    CmdEdit.Enabled = False
End Sub

Private Sub PreviewPV()
Const METHOD_NAME = "PreviewPV"
On Error GoTo debugs

Dim REceipt As New ADODB.Recordset
Dim ReceiptNo As String
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim strSql As String
Dim PV As New ADODB.Recordset
Dim tmpCase, tmpIndex As String
Dim RecCnt As Integer
Dim tmpClinicCode As String
Dim tmpCLNName As String
Dim tmpCLNAdd1 As String
Dim tmpCLNAdd2 As String
Dim tmpCLNAdd3 As String
Dim tmpPCode As String
Dim tmpCity As String
Dim tmpState As String

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

    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\PV-Clinic.xlt")
    Set objWorksheet = objWorkbook.Sheets(1)
    objExcel.DisplayAlerts = False
    Set objWorksheet = objWorkbook.Worksheets.Item(1)
    For i = 1 To lstPaymentCLN.ListItems.Count
         If lstPaymentCLN.ListItems(i).Checked = True Then
            tmpCase = Trim(lstPaymentCLN.ListItems(i).SubItems(12))
            tmpIndex = Trim(lstPaymentCLN.ListItems(i).SubItems(17))
            strSql = "select PVCLNIndex, CASNumber,BordxRefNo, BillNo, BillDate, PVCLNNo, PVCLNChequeNo, PVCLNDate, " & _
                "ClinicCode, ClinicGrpCode, CLNName, CLNAdd1, CLNAdd2,CLNAdd3, CLNPostCode, CLNCity, CLNState, CLNPaidAmt, CLMPayableByMedix2C " & _
                "from View_Payment2CLN where CASNumber = '" & tmpCase & "'"  '& "' and PVCLNIndex='" & Trim(tmpIndex) & "'"
            PV.Open strSql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
            Do While Not PV.EOF
                tmpClinicCode = Trim(PV!ClinicCode)
                tmpCLNName = Trim(PV!CLNName)
                tmpCLNAdd1 = Trim(PV!CLNAdd1)
                tmpCLNAdd2 = Trim(PV!CLNAdd2)
                tmpCLNAdd3 = Trim(PV!CLNAdd3)
                tmpPCode = Trim(PV!CLNPostCode)
                tmpCity = Trim(PV!CLNCity)
                tmpState = Trim(PV!CLNState)
                If Trim(PV!ClinicGrpCode) <> "" Then
                    tmpClinicCode = Trim(PV!ClinicGrpCode)
                    Call ClinicHQInfo(tmpClinicCode, tmpCLNName, tmpCLNAdd1, tmpCLNAdd2, tmpCLNAdd3, tmpPCode, tmpCity, tmpState)
                End If

                objWorksheet.Cells(8, "C") = tmpCLNName & " (" & tmpClinicCode & ")"
                objWorksheet.Cells(9, "C") = tmpCLNAdd1
                objWorksheet.Cells(10, "C") = tmpCLNAdd2 & "," & tmpCLNAdd3
                objWorksheet.Cells(11, "C") = Trim(tmpPCode & " " & Trim(tmpCity) & " " & Trim(tmpState))
                objWorksheet.Cells(8, "I") = TxtPVNo
                objWorksheet.Cells(9, "I") = IIf(Len(PV!PVCLNChequeNo), PV!PVCLNChequeNo, "")
                objWorksheet.Cells(10, "I") = IIf(Len(PV!PVCLNDate), PV!PVCLNDate, "")
                objWorksheet.Cells(37, "C") = Logon
                ReceiptNo = ""
                REceipt.Open "SELECT ReceiptNo from REceiptFrPAY where CASNumber ='" & tmpCase & "'", MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
                If Not REceipt.EOF Then
                    ReceiptNo = REceipt!ReceiptNo
                Else
                    REceipt.Close
                    Set REceipt = Nothing
                    REceipt.Open "SELECT ReceiptNo from REceiptFrPAY where CASNumber ='" & tmpCase & "'", MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
                    Do While Not REceipt.EOF
                        ReceiptNo = REceipt!ReceiptNo
                        REceipt.MoveNext
                    Loop
                End If
                REceipt.Close
                Set REceipt = Nothing
                RecCnt = RecCnt + 1
                objWorksheet.Cells(15 + CDbl(RecCnt), "A") = RecCnt
                objWorksheet.Cells(15 + CDbl(RecCnt), "B") = PV!BillNo
                objWorksheet.Cells(15 + CDbl(RecCnt), "C") = PV!BillDate
                objWorksheet.Cells(15 + CDbl(RecCnt), "D") = PV!CasNumber
                objWorksheet.Cells(15 + CDbl(RecCnt), "E") = PV!Bordxrefno
                objWorksheet.Cells(15 + CDbl(RecCnt), "I") = IIf(IsNumeric(PV!CLMPayableByMedix2c), PV!CLMPayableByMedix2c, 0)
                objWorksheet.Cells(15 + CDbl(RecCnt), "G") = ReceiptNo
                PV.MoveNext
            Loop
            PV.Close
            Set PV = Nothing
        End If
    Next i
    
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

Private Sub ClinicHQInfo(tmpClinicCode, tmpCLNName, tmpCLNAdd1, tmpCLNAdd2, tmpCLNAdd3, tmpPCode, tmpCity, tmpState As String)
Const METHOD_NAME = "ClinicHQInfo"
On Error GoTo debugs

Dim StrCLNSql As String
Dim CLNRec As New ADODB.Recordset
    StrCLNSql = "SELECT CLNCode, CLNName, CLNAdd1, CLNAdd2,CLNAdd3, CLNPostCode, CLNCity, CLNState " & _
                " FROM CLN WHERE CLNCode='" & tmpClinicCode & "'"
    CLNRec.Open StrCLNSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    Do While Not CLNRec.EOF
        With CLNRec
            tmpClinicCode = Trim(!CLNCode)
            tmpCLNName = Trim(!CLNName)
            tmpCLNAdd1 = Trim(!CLNAdd1)
            tmpCLNAdd2 = Trim(!CLNAdd2)
            tmpCLNAdd3 = Trim(!CLNAdd3)
            tmpPCode = Trim(!CLNPostCode)
            tmpCity = Trim(!CLNCity)
            tmpState = Trim(!CLNState)
        End With
        CLNRec.MoveNext
    Loop
    CLNRec.Close
    Set CLNRec = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub


