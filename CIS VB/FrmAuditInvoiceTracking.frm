VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAuditInvoiceTracking 
   Caption         =   "Invoice Tracking"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   4395
      Left            =   0
      TabIndex        =   69
      Top             =   240
      Width           =   11775
      Begin VB.ComboBox cboauditind 
         Height          =   315
         Left            =   1200
         TabIndex        =   133
         Top             =   3920
         Width           =   1575
      End
      Begin VB.ComboBox CboDepartment 
         Height          =   315
         Left            =   7560
         Sorted          =   -1  'True
         TabIndex        =   131
         Top             =   3960
         Width           =   3615
      End
      Begin VB.CheckBox chkOMNIRef 
         Height          =   255
         Left            =   5475
         TabIndex        =   117
         Top             =   3390
         Width           =   255
      End
      Begin VB.CheckBox ChkPayor2 
         Height          =   255
         Left            =   5760
         TabIndex        =   110
         Top             =   1120
         Width           =   255
      End
      Begin VB.CheckBox ChkOPBatchNo 
         Height          =   255
         Left            =   11280
         TabIndex        =   57
         Top             =   2800
         Width           =   375
      End
      Begin VB.CheckBox ChkYear 
         Height          =   255
         Left            =   11280
         TabIndex        =   54
         Top             =   2570
         Width           =   375
      End
      Begin VB.CheckBox ChkMonth 
         Height          =   255
         Left            =   11280
         TabIndex        =   51
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox chkInvReg 
         Height          =   255
         Left            =   11280
         TabIndex        =   45
         Top             =   1695
         Width           =   375
      End
      Begin VB.CheckBox chkInvDate 
         Height          =   255
         Left            =   11280
         TabIndex        =   42
         Top             =   1455
         Width           =   375
      End
      Begin VB.CheckBox chkDOR 
         Height          =   255
         Left            =   11280
         TabIndex        =   39
         Top             =   1200
         Width           =   375
      End
      Begin VB.CheckBox chkDOV 
         Height          =   255
         Left            =   11280
         TabIndex        =   36
         Top             =   900
         Width           =   375
      End
      Begin VB.CheckBox ChkInvAmt 
         Height          =   255
         Left            =   11280
         TabIndex        =   48
         Top             =   1995
         Width           =   375
      End
      Begin VB.CheckBox ChkPayTo 
         Height          =   255
         Left            =   5480
         TabIndex        =   5
         Top             =   580
         Width           =   255
      End
      Begin VB.CheckBox ChkClinic 
         Height          =   255
         Left            =   5480
         TabIndex        =   7
         Top             =   850
         Width           =   255
      End
      Begin VB.CheckBox ChkPayor 
         Height          =   255
         Left            =   5480
         TabIndex        =   10
         Top             =   1120
         Width           =   255
      End
      Begin VB.CheckBox ChkMbmNo 
         Height          =   255
         Left            =   5480
         TabIndex        =   18
         Top             =   1970
         Width           =   255
      End
      Begin VB.CheckBox ChkInvNo 
         Height          =   255
         Left            =   5480
         TabIndex        =   21
         Top             =   2235
         Width           =   255
      End
      Begin VB.CheckBox ChkCaseNo 
         Height          =   255
         Left            =   5480
         TabIndex        =   3
         Top             =   280
         Width           =   255
      End
      Begin VB.CheckBox ChkBordxNo 
         Height          =   255
         Left            =   5480
         TabIndex        =   24
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox ChkName 
         Height          =   255
         Left            =   5480
         TabIndex        =   15
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox ChkPlan 
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
         Height          =   280
         Left            =   11280
         TabIndex        =   30
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtPlan1 
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
         Left            =   7560
         MaxLength       =   4
         TabIndex        =   28
         Top             =   255
         Width           =   1455
      End
      Begin VB.TextBox TxtPlan2 
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
         Left            =   9720
         MaxLength       =   4
         TabIndex        =   29
         Top             =   255
         Width           =   1455
      End
      Begin VB.CheckBox ChkICNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   11280
         TabIndex        =   33
         Top             =   570
         Width           =   375
      End
      Begin VB.TextBox TxtWSNo1 
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
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   1
         Top             =   255
         Width           =   1605
      End
      Begin VB.TextBox TxtWSNo2 
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
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   2
         Top             =   255
         Width           =   1575
      End
      Begin VB.CheckBox ChkClinic2 
         Height          =   255
         Left            =   5760
         TabIndex        =   8
         Top             =   850
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom2 
         Height          =   255
         Left            =   5760
         TabIndex        =   13
         Top             =   1400
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom 
         Height          =   255
         Left            =   5475
         TabIndex        =   12
         Top             =   1400
         Width           =   255
      End
      Begin VB.ComboBox CboPayTo 
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
         TabIndex        =   4
         Text            =   "C - Clinic"
         Top             =   540
         Width           =   4220
      End
      Begin VB.ComboBox CboClinic 
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
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   830
         Width           =   4220
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
         Left            =   1200
         TabIndex        =   9
         Top             =   1080
         Width           =   4220
      End
      Begin VB.ComboBox CboGrpCom 
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
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   1365
         Width           =   4220
      End
      Begin VB.TextBox TxtPatientName 
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
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   14
         Top             =   1635
         Width           =   4220
      End
      Begin VB.TextBox TxtMBMNo1 
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
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   16
         Top             =   1920
         Width           =   1570
      End
      Begin VB.TextBox TxtInvNo1 
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
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   19
         Top             =   2205
         Width           =   1570
      End
      Begin VB.TextBox TxtBordx1 
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
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   22
         Top             =   2490
         Width           =   1570
      End
      Begin VB.ComboBox CboCaseStatus 
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
         TabIndex        =   25
         Top             =   2775
         Width           =   1570
      End
      Begin VB.TextBox TxtMBMNo2 
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
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   17
         Top             =   1920
         Width           =   1570
      End
      Begin VB.TextBox TxtInvNo2 
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
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   20
         Top             =   2205
         Width           =   1570
      End
      Begin VB.TextBox TxtBordx2 
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
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   23
         Top             =   2490
         Width           =   1570
      End
      Begin VB.ComboBox CboClaimability 
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
         Left            =   3840
         TabIndex        =   26
         Top             =   2775
         Width           =   1570
      End
      Begin VB.TextBox TxtICNo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7560
         MaxLength       =   20
         TabIndex        =   31
         Top             =   540
         Width           =   1455
      End
      Begin VB.TextBox TxtICNo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9720
         MaxLength       =   20
         TabIndex        =   32
         Top             =   540
         Width           =   1455
      End
      Begin VB.ComboBox CboClinicType 
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
         TabIndex        =   27
         Top             =   3060
         Width           =   1570
      End
      Begin MSMask.MaskEdBox txtDOVDate1 
         Height          =   315
         Left            =   7560
         TabIndex        =   34
         Top             =   810
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox txtDORDate1 
         Height          =   315
         Left            =   7560
         TabIndex        =   37
         Top             =   1100
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox txtDOVDate2 
         Height          =   315
         Left            =   9720
         TabIndex        =   35
         Top             =   810
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox txtDORDate2 
         Height          =   315
         Left            =   9720
         TabIndex        =   38
         Top             =   1100
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   7560
         TabIndex        =   40
         Top             =   1385
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   9720
         TabIndex        =   41
         Top             =   1385
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox txtIRDDate1 
         Height          =   315
         Left            =   7560
         TabIndex        =   43
         Top             =   1670
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.TextBox txtINVAmtFr 
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
         Left            =   7560
         TabIndex        =   46
         Top             =   1955
         Width           =   1455
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
         Left            =   7560
         TabIndex        =   49
         Top             =   2210
         Width           =   1455
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
         Left            =   7560
         TabIndex        =   52
         Top             =   2495
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtIRDDate2 
         Height          =   315
         Left            =   9720
         TabIndex        =   44
         Top             =   1670
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.TextBox txtINVAmtTo 
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
         Left            =   9720
         TabIndex        =   47
         Top             =   1955
         Width           =   1455
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
         Left            =   9720
         TabIndex        =   50
         Top             =   2210
         Width           =   1455
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
         Left            =   9720
         TabIndex        =   53
         Top             =   2495
         Width           =   1455
      End
      Begin VB.TextBox TxtOPBatchNo1 
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
         Left            =   7560
         MaxLength       =   15
         TabIndex        =   55
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox TxtOPBatchNo2 
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
         Left            =   9720
         MaxLength       =   15
         TabIndex        =   56
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ComboBox cboInfoFFS 
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
         Left            =   3840
         Sorted          =   -1  'True
         TabIndex        =   111
         Top             =   3060
         Width           =   1570
      End
      Begin VB.ComboBox cboOMNIStatus 
         Height          =   315
         Left            =   7560
         TabIndex        =   113
         Top             =   3015
         Width           =   3615
      End
      Begin VB.ComboBox cboSubsidiary 
         Height          =   315
         Left            =   7560
         Sorted          =   -1  'True
         TabIndex        =   115
         Top             =   3300
         Width           =   3615
      End
      Begin VB.TextBox txtOMNIRefFr 
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
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   118
         Top             =   3350
         Width           =   1570
      End
      Begin VB.TextBox txtOMNIRefTo 
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
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   119
         Top             =   3350
         Width           =   1570
      End
      Begin VB.ComboBox cboTIMSInd 
         Height          =   315
         Left            =   1200
         TabIndex        =   122
         Top             =   3630
         Width           =   1575
      End
      Begin MSMask.MaskEdBox cmdBordxDateFr 
         Height          =   315
         Left            =   7560
         TabIndex        =   124
         Top             =   3590
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox cmdBordxDateTo 
         Height          =   315
         Left            =   9720
         TabIndex        =   125
         Top             =   3590
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.ComboBox cboOMNISent 
         Height          =   315
         Left            =   3840
         TabIndex        =   128
         Top             =   3630
         Width           =   1575
      End
      Begin VB.Label lblauditchecking 
         Caption         =   "Audited"
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
         Index           =   4
         Left            =   120
         TabIndex        =   132
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label lblDepartment 
         Caption         =   "Department"
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
         TabIndex        =   130
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "OMNI Sent"
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
         Index           =   2
         Left            =   2880
         TabIndex        =   129
         Top             =   3660
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx Date"
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
         Left            =   6360
         TabIndex        =   127
         ToolTipText     =   "Date of Visit"
         Top             =   3640
         Width           =   855
      End
      Begin VB.Label Label14 
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
         TabIndex        =   126
         Top             =   3640
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "TIMS"
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
         TabIndex        =   123
         Top             =   3660
         Width           =   855
      End
      Begin VB.Label Label47 
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
         Left            =   3240
         TabIndex        =   121
         Top             =   3420
         Width           =   255
      End
      Begin VB.Label Label27 
         Caption         =   "OMNI Ref No"
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
         TabIndex        =   120
         Top             =   3380
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "Subsidiary"
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
         TabIndex        =   116
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "OMNI"
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
         Index           =   3
         Left            =   6360
         TabIndex        =   114
         Top             =   3080
         Width           =   855
      End
      Begin VB.Label txtFFSStatus 
         Caption         =   "FFS Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   112
         Top             =   3100
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Clinic Type"
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
         TabIndex        =   109
         Top             =   3100
         Width           =   975
      End
      Begin VB.Label Label12 
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
         Left            =   6360
         TabIndex        =   108
         Top             =   2800
         Width           =   1095
      End
      Begin VB.Label Label10 
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
         TabIndex        =   107
         Top             =   2800
         Width           =   375
      End
      Begin VB.Label Label18 
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
         Left            =   11280
         TabIndex        =   105
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label45 
         Caption         =   "A/C Period Mth"
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
         TabIndex        =   104
         Top             =   2265
         Width           =   1215
      End
      Begin VB.Label Label44 
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
         TabIndex        =   103
         Top             =   2265
         Width           =   375
      End
      Begin VB.Label Label43 
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
         TabIndex        =   102
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label42 
         Caption         =   "A/C Period Yr"
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
         TabIndex        =   101
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label41 
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
         TabIndex        =   100
         Top             =   1960
         Width           =   375
      End
      Begin VB.Label Label40 
         Caption         =   "Inv Amount"
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
         TabIndex        =   99
         Top             =   1960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Claimability"
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
         Left            =   2835
         TabIndex        =   98
         Top             =   2800
         Width           =   855
      End
      Begin VB.Label Label4 
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
         TabIndex        =   97
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label6 
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
         TabIndex        =   96
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Patient Name"
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
         Top             =   1770
         Width           =   975
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
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
         Left            =   3240
         TabIndex        =   94
         Top             =   285
         Width           =   255
      End
      Begin VB.Label Label27 
         Caption         =   "Bordx No"
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
         Top             =   2550
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Inv No"
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
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label25 
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
         TabIndex        =   91
         Top             =   280
         Width           =   975
      End
      Begin VB.Label Label5 
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
         Left            =   120
         TabIndex        =   90
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Inv Date"
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
         TabIndex        =   89
         Top             =   1420
         Width           =   975
      End
      Begin VB.Label Label13 
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
         TabIndex        =   88
         Top             =   1420
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Plan Code"
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
         TabIndex        =   87
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label15 
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
         TabIndex        =   86
         Top             =   1150
         Width           =   495
      End
      Begin VB.Label Label14 
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
         TabIndex        =   85
         Top             =   880
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "DOR"
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
         TabIndex        =   84
         ToolTipText     =   "Date of Register"
         Top             =   1150
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "DOV"
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
         Left            =   6360
         TabIndex        =   83
         ToolTipText     =   "Date of Visit"
         Top             =   880
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Clinic"
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
         TabIndex        =   82
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
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
         TabIndex        =   81
         Top             =   1190
         Width           =   975
      End
      Begin VB.Label Label23 
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
         TabIndex        =   80
         Top             =   1690
         Width           =   495
      End
      Begin VB.Label Label24 
         Caption         =   "Inv Reg Date"
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
         TabIndex        =   79
         Top             =   1690
         Width           =   1140
      End
      Begin VB.Label Label11 
         Caption         =   "Case Status"
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
         TabIndex        =   78
         Top             =   2800
         Width           =   975
      End
      Begin VB.Label Label19 
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
         Left            =   3240
         TabIndex        =   77
         Top             =   2235
         Width           =   255
      End
      Begin VB.Label Label46 
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
         Left            =   3240
         TabIndex        =   76
         Top             =   1995
         Width           =   255
      End
      Begin VB.Label Label47 
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
         Left            =   3240
         TabIndex        =   75
         Top             =   2550
         Width           =   255
      End
      Begin VB.Label Label84 
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
         TabIndex        =   74
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label85 
         Caption         =   "I/C No"
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
         TabIndex        =   73
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label52 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   135
         Left            =   5520
         TabIndex        =   72
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label33 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   135
         Left            =   5760
         TabIndex        =   71
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label49 
         Caption         =   "Grp Company"
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
         TabIndex        =   70
         Top             =   1490
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstInvoiceTrackList 
      Height          =   2385
      Left            =   0
      TabIndex        =   68
      Top             =   4680
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   4207
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   25
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IC No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Case Number"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Date Visit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Med Inj Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Med Inj Days"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Med Inj Qty"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Med Charges"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Consultation Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Invoice Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Case Remark"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Audit Query"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Audit Answer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "Audit Remark"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "Cost Avoidance"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Reason Desc1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Reason Desc2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Text            =   "Reason Desc3"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Reason Desc4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Reason Desc5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Audit By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Audit Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "CNL Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "CLN Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Group Company"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Width           =   11775
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
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
         Left            =   8880
         TabIndex        =   61
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
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
         Left            =   8280
         TabIndex        =   60
         Top             =   200
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
         Left            =   7080
         TabIndex        =   58
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdPrintFile 
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
         TabIndex        =   62
         Top             =   200
         Width           =   975
      End
      Begin VB.TextBox txtNoofDay 
         Height          =   285
         Left            =   120
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
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
         TabIndex        =   64
         Top             =   200
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
         TabIndex        =   63
         Top             =   200
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
         Left            =   7680
         TabIndex        =   59
         Top             =   200
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   1200
         TabIndex        =   67
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1440
         TabIndex        =   66
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Audit Invoice Tracking"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   0
      TabIndex        =   106
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmAuditInvoiceTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim tmpPVNo As String
Dim tmpPVDate As String
Dim tmpChqNo As String
Dim tmpChqDate As String
Dim SQLqueryappend As String

Private Sub CboGrpCom_LostFocus()
Const METHOD_NAME = "CboGrpCom_LostFocus"
On Error GoTo debugs
Dim sql As String
Dim sql2 As String
Dim GRPCOMName As String
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strordergroup As String
    '=Run Sp=================================================
    sql2 = ""
    sql2 = "Exec SearchSubsidiary '" & Trim(Replace(CboGrpCom, "'", "''")) & "'"
    rst3.Open sql2, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    '===================================================
    cboSubsidiary.Clear
    Do Until rst3.EOF
        With rst3
            cboSubsidiary.AddItem IIf(Len(Trim(!MBMGRPSubsidiary)), Trim(!MBMGRPSubsidiary), "")
            rst3.MoveNext
        End With
    Loop
    rst3.Close
    Set rst3 = Nothing
    '=Run Sp=================================================
    sql2 = ""
        sql2 = "Exec SearchDepartment  '" & Trim(Replace(CboGrpCom, "'", "''")) & "'"
        rst3.Open sql2, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    '===================================================
    CboDepartment.Clear
    Do Until rst3.EOF
        With rst3
        CboDepartment.AddItem IIf(Len(Trim(!MBMDepartment)), Trim(!MBMDepartment), "")
        rst3.MoveNext
        End With
        Loop
        rst3.Close
        Set rst3 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

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

Private Sub ChkPayor2_Click()
Dim PAY As New ADODB.Recordset
Dim StrSql As String
    CboPayor.Clear
    If ChkPayor2.Value = 1 Then
        StrSql = "EXEC SearchPayor"
        PAY.Open StrSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
        
        Do Until PAY.EOF
            CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
            PAY.MoveNext
        Loop
        PAY.Close
        Set PAY = Nothing
    End If
    
End Sub
Private Sub CmdPrintFile_Click()
    If SQLqueryappend = "" And lstInvoiceTrackList.ListItems.Count = 0 Then
        MsgBox "Please search again.!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Call AuditPrintingdetails
    End If
End Sub

Private Sub AuditPrintingdetails()
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim strAppendCase As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim rst2 As New ADODB.Recordset
    Dim tmpcount As String
    Dim stintStartRow, intStartRow, intStartcell As String
    
    
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
        objExcel.Interactive = False
        objExcel.Visible = True
        objExcel.WindowState = xlMaximized
        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\InvoiceAuditTrackingNEW.xlt")
        Set objWorksheet = objWorkbook.Sheets(1)
        objExcel.DisplayAlerts = True
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
    End If
    
    strAppendCase = "1"
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "Sp_Audittraccking_2018"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, SQLqueryappend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    objWorksheet.Cells(1, "C") = Date
     intStartRow = 5
     Dim isSecondRecord As Boolean
     isSecondRecord = False
        
        Do Until rst2.EOF
            With rst2
                If rst2!RNo = 1 Then
                    
                    objWorksheet.Range("A" + CStr(intStartRow) + ":" + "A" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "A") = IIf(Len(rst2!Bordxrefno), rst2!Bordxrefno, "")
                    objWorksheet.Range("B" + CStr(intStartRow) + ":" + "B" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "B") = IIf(Len(rst2!MBMIcBcPp), rst2!MBMIcBcPp, "")
                    objWorksheet.Range("C" + CStr(intStartRow) + ":" + "C" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "C") = IIf(Len(rst2!MBMName), rst2!MBMName, "")
                    objWorksheet.Range("D" + CStr(intStartRow) + ":" + "D" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "D") = IIf(Len(rst2!CasNumber), rst2!CasNumber, "")
                    objWorksheet.Range("E" + CStr(intStartRow) + ":" + "E" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "E") = IIf(Len(rst2!DateVisit), rst2!DateVisit, "")
                    
                    objWorksheet.Range("F" + CStr(intStartRow) + ":" + "F" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "F") = IIf(Len(rst2!ReasonDesc1), rst2!ReasonDesc1, "")
                    
                    objWorksheet.Range("G" + CStr(intStartRow) + ":" + "G" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "G") = IIf(Len(rst2!ReasonDesc2), rst2!ReasonDesc2, "")
                    
                    objWorksheet.Range("H" + CStr(intStartRow) + ":" + "H" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "H") = IIf(Len(rst2!ReasonDesc3), rst2!ReasonDesc3, "")
                    
                    objWorksheet.Range("I" + CStr(intStartRow) + ":" + "I" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "I") = IIf(Len(rst2!ReasonDesc4), rst2!ReasonDesc4, "")
                    
                    If Len(rst2!CASDiagnosis5) Then
                        objWorksheet.Range("J" + CStr(intStartRow) + ":" + "J" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                        objWorksheet.Cells(intStartRow, "J") = IIf(Len(rst2!CASDiagnosis5), rst2!CASDiagnosis5, "")
                    Else
                        objWorksheet.Range("J" + CStr(intStartRow) + ":" + "J" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                        objWorksheet.Cells(intStartRow, "J") = IIf(Len(rst2!ReasonDesc5), rst2!ReasonDesc5, "")
                    End If
                    '''
                    objWorksheet.Cells(intStartRow, "K") = IIf(Len(rst2!MedInjType1), rst2!MedInjType1, "")
                    objWorksheet.Cells(intStartRow, "L") = IIf(Len(rst2!MedInjDays1), rst2!MedInjDays1, "")
                    objWorksheet.Cells(intStartRow, "M") = IIf(Len(rst2!MedInjQty1), rst2!MedInjQty1, "")
                    objWorksheet.Cells(intStartRow, "N") = IIf(Len(rst2!MedCharges1), rst2!MedCharges1, "")
                    '''
                    objWorksheet.Range("O" + CStr(intStartRow) + ":" + "O" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "O") = IIf(Len(rst2!Procedurefee), rst2!Procedurefee, "")
                    objWorksheet.Range("P" + CStr(intStartRow) + ":" + "P" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "P") = IIf(Len(rst2!SPFee), rst2!SPFee, "")
                    objWorksheet.Range("Q" + CStr(intStartRow) + ":" + "Q" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "Q") = IIf(Len(rst2!INVAmount), rst2!INVAmount, "")
                    objWorksheet.Range("R" + CStr(intStartRow) + ":" + "R" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "R") = IIf(Len(rst2!AuditQuery), rst2!AuditQuery, "")
                    objWorksheet.Range("S" + CStr(intStartRow) + ":" + "S" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "S") = IIf(Len(rst2!AuditAnswer), rst2!AuditAnswer, "")
                    objWorksheet.Range("T" + CStr(intStartRow) + ":" + "T" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "T") = IIf(Len(rst2!AuditRemarks), rst2!AuditRemarks, "")
                    objWorksheet.Range("U" + CStr(intStartRow) + ":" + "U" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "U") = IIf(Len(rst2!AuditDiscountReceived), rst2!AuditDiscountReceived, "")
                    objWorksheet.Range("V" + CStr(intStartRow) + ":" + "V" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "V") = IIf(Len(rst2!AuditBy), rst2!AuditBy, "")
                    
                    objWorksheet.Range("W" + CStr(intStartRow) + ":" + "W" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "W") = IIf(Len(rst2!CLNCode), rst2!CLNCode, "")
                    
                    objWorksheet.Range("X" + CStr(intStartRow) + ":" + "X" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "X") = IIf(Len(rst2!CLNName), rst2!CLNName, "")
                    
                    objWorksheet.Range("Y" + CStr(intStartRow) + ":" + "Y" + CStr((intStartRow + rst2!MedCount) - 1)).MergeCells = True
                    objWorksheet.Cells(intStartRow, "Y") = IIf(Len(rst2!MBMGroupCompany), rst2!MBMGroupCompany, "")
                    
                ElseIf rst2!RNo <> 1 Then
                    objWorksheet.Cells((intStartRow - 1) + rst2!RNo, "K") = IIf(Len(rst2!MedInjType1), rst2!MedInjType1, "")
                    objWorksheet.Cells((intStartRow - 1) + rst2!RNo, "L") = IIf(Len(rst2!MedInjDays1), rst2!MedInjDays1, "")
                    objWorksheet.Cells((intStartRow - 1) + rst2!RNo, "M") = IIf(Len(rst2!MedInjQty1), rst2!MedInjQty1, "")
                    objWorksheet.Cells((intStartRow - 1) + rst2!RNo, "N") = IIf(Len(rst2!MedCharges1), rst2!MedCharges1, "")
                End If
                
                If rst2!MedCount = rst2!RNo Then
                    intStartRow = intStartRow + rst2!MedCount
                End If
            End With
            rst2.MoveNext
        Loop
        
    rst2.Close
    Set rst2 = Nothing
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    Screen.MousePointer = vbDefault
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
    Call ActionLogUpdate(12, "FrmAuditInvoiceTracking")
    intExit = 0
    Call ComboListing
End Sub

'************************Start Command Button**************************
Private Sub CmdExit_Click()
    SQLqueryappend = ""
    Unload Me
    
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    lstInvoiceTrackList.ListItems.Clear
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    CmdSearch.Enabled = True
    cmdClear.Enabled = True
    CmdPrintFile.Enabled = True
    cmdExit.Enabled = True
End Sub

Public Sub cmdSearch_Click()
'On Error GoTo errRun

    CmdSearch.Enabled = False
    cmdClear.Enabled = False
    CmdPrintFile.Enabled = False
    cmdExit.Enabled = False
    Screen.MousePointer = vbHourglass
        Call SearchRecord
    Screen.MousePointer = Default
    Exit Sub
errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lstInvoiceTrackList.ListItems.Clear
    CmdSearch.Enabled = True
    cmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrintFile.Enabled = True
End Sub

Private Sub cmdBack_Click()
    lstInvoiceTrackList.Height = 3585
    lstInvoiceTrackList.Left = 0
    lstInvoiceTrackList.Top = 3480
    lstInvoiceTrackList.Width = 11775
    Frame1.Visible = True
End Sub



Private Sub CmdView_Click()
    lstInvoiceTrackList.Height = 6705
    lstInvoiceTrackList.Left = 0
    lstInvoiceTrackList.Top = 360
    lstInvoiceTrackList.Width = 11775
    Frame1.Visible = False
End Sub
'************************Start ComboListing**************************
Private Sub ComboListing()
Dim sql As String
Dim yearstart
Dim monthstart
   
    With CboPayto
        .AddItem "C - Clinic"
        .AddItem "M - Member"
    End With
    
    With cboClaimability
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
    
    cboOMNIStatus.Clear
    cboOMNIStatus.AddItem "Y" & " - " & "Yes"
    
    cboTIMSInd.Clear
    cboTIMSInd.AddItem "Y" & " - " & "Yes"
    
    cboauditind.Clear
    cboauditind.AddItem "Y - Yes"
    
    cboOMNISent.AddItem "Y - Yes"
    cboOMNISent.AddItem "N - No"
End Sub
'************************Start ComboListing**************************

'************************Start Search Record**************************

Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
'On Error GoTo debugs
Dim StrSql As String
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
        strAppend = strAppend + " AND ciscas.CASNumber >= '" & Trim(TxtWSNo1) & "'"
    End If
    
    If Len(Trim(TxtWSNo2)) Then
        strAppend = strAppend + " AND ciscas.CASNumber <= '" & Trim(TxtWSNo2) & "'"
    End If
    
    If ChkCaseNo.Value = 1 Then
        sql = ""
        sql = "AND (ciscas.CASNumber Is null OR ciscas.CASNumber = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(CboPayto)) Then
        strAppend = strAppend + " AND INVPayTo = '" & Trim(Mid(CboPayto, 1, IIf(InStr(1, CboPayto, "-") = 0, 4, InStr(1, CboPayto, "-") - 1))) & "'"
    End If
    
    If ChkPayTo.Value = 1 Then
        sql = ""
        sql = "AND (INVPayTo Is null OR INVPayTo = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(CboClinic)) Then
    Dim sArrclincicode() As String
    Dim SArrccode As String
        sArrclincicode = Split(CboClinic, "-")
        SArrccode = Trim(sArrclincicode(0))
        strAppend = strAppend + " AND INVPayee = '" & Trim(SArrccode) & "'"
        
        'strsqlHOS = "select CLNCode, CLNName from CLN where CLNCode = '" & Trim(RemStr(CboClinic)) & "'"
        '=========================trim CLNCLinic===============
        
'        strsqlHOS = "select CLNCode, CLNName from CLN where CLNCode = '" & Mid(CboClinic, 1, 4) & "'"
'        CLN.Open strsqlHOS, medixmasterconn, adOpenForwardOnly, adLockReadOnly
'
'        If CLN.EOF = False Then
'            strAppend = strAppend + " AND INVPayee = '" & Trim(CLN!CLNCode) & "'"
'        End If
'
'        CLN.Close
'        Set CLN = Nothing
    End If
    
    If ChkClinic.Value = 1 Then
        sql = ""
        sql = "AND (INVPayee Is null OR INVPayee = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And MBMCLNOthers.PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If ChkPayor.Value = 1 Then
        sql = ""
        sql = "AND (MBMCLNOthers.PAYCode Is null OR MBMCLNOthers.PAYCode = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(Replace(CboGrpCom, "'", "''"))) Then
       strAppend = strAppend + " AND MBMGroupCompany = '" & Trim(RemStr(CboGrpCom)) & "'"
    End If
    
    If ChkGrpCom.Value = 1 Then
       strAppend = strAppend + " AND (MBMGroupCompany IS NULL OR MBMGroupCompany = '')"
    End If
    
    If Len(Trim(txtPatientName)) Then
       strAppend = strAppend + " And mbmname = '" & Trim(txtPatientName) & "'"
    End If
    
    If ChkName.Value = 1 Then
        sql = ""
        sql = "AND (mbmname Is null OR mbmname = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtMBMNo1)) Then
        strAppend = strAppend + " AND CISCas.MBMNumber >= '" & Trim(TxtMBMNo1) & "'"
    End If
    
    If Len(Trim(TxtMBMNo2)) Then
        strAppend = strAppend + " AND CISCas.MBMNumber <= '" & Trim(TxtMBMNo2) & "'"
    End If
    
    If ChkMbmNo.Value = 1 Then
        sql = ""
        sql = "AND (CISCas.MBMNumber Is null OR CISCas.MBMNumber = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtInvNo1)) Then
        strAppend = strAppend + " AND CISCas.InvoiceNo >= '" & Trim(txtInvNo1) & "'"
    End If

    If Len(Trim(txtInvNo2)) Then
        strAppend = strAppend + " AND CISCas.InvoiceNo <= '" & Trim(txtInvNo2) & "'"
    End If

    If ChkInvNo.Value = 1 Then
        sql = ""
        sql = "AND (CISCas.InvoiceNo Is null OR CISCas.InvoiceNo = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtBordx1)) Then
        strAppend = strAppend + " AND Claim.BordxRefNo >= '" & Trim(TxtBordx1) & "'"
    End If
    
    If Len(Trim(TxtBordx2)) Then
        strAppend = strAppend + " AND Claim.BordxRefNo <= '" & Trim(TxtBordx2) & "'"
    End If
           
    If ChkBordxNo.Value = 1 Then
        sql = ""
        sql = "AND (Claim.BordxRefNo Is null OR Claim.BordxRefNo = '')" & ""
        strAppend = strAppend + sql
    End If
        
    If Mid(CboCaseStatus, 1, 1) = "C" Then
        strAppend = strAppend + " AND CISCas.CASStatus = 'C'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "O" Then
        strAppend = strAppend + " AND CISCas.CASStatus = 'O'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "D" Then
        strAppend = strAppend + " AND CISCas.CASStatus = 'D'"
    End If
    
    If Mid(cboClaimability, 1, 1) = "R" Then
        strAppend = strAppend + " Or CISCas.Claimability = 'R'"
    ElseIf Mid(cboClaimability, 1, 1) = "A" Then
        strAppend = strAppend + " Or CISCas.Claimability = 'A'"
    End If
    
    If Len(Trim(TxtPlan1)) Then
        strAppend = strAppend + " AND MBMCLNOthers.PLNType >= '" & Trim(TxtPlan1) & "'"
    End If
    
    If Len(Trim(TxtPlan2)) Then
        strAppend = strAppend + " AND MBMCLNOthers.PLNType <= '" & Trim(TxtPlan2) & "'"
    End If
    
    If ChkPlan.Value = 1 Then
        sql = ""
        sql = "AND (MBMCLNOthers.PLNType Is null OR MBMCLNOthers.PLNType = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(cmdBordxDateFr)) > 0 And IsDate(cmdBordxDateFr) = True Then
        strAppend = strAppend + " And Bordx.BordxDate >= '" & Format(cmdBordxDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(cmdBordxDateTo)) > 0 And IsDate(cmdBordxDateTo) = True Then
        strAppend = strAppend + " And Bordx.BordxDate <= '" & Format(cmdBordxDateTo, "mm/dd/yyyy") & "'"
    End If
        
    If Len(Trim(TxtICNo1)) Then
        strAppend = strAppend + " AND MBMIcBcPp >= '" & Trim(TxtICNo1) & "'"
    End If
    
    If Len(Trim(TxtICNo2)) Then
        strAppend = strAppend + " AND MBMIcBcPp <= '" & Trim(TxtICNo2) & "'"
    End If
    
    If ChkICNo.Value = 1 Then
      strAppend = strAppend + " AND (MBMIcBcPp Is null)"
    End If
    
    If Len(Trim(txtDOVDate1)) > 0 And IsDate(txtDOVDate1) = True Then
        strAppend = strAppend + " And CISCas.DateVisit >= '" & Format(txtDOVDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOVDate2)) > 0 And IsDate(txtDOVDate2) = True Then
        strAppend = strAppend + " And CISCas.DateVisit <= '" & Format(txtDOVDate2, "mm/dd/yyyy") & "'"
    End If
    
    If chkDOV.Value = 1 Then
        sql = ""
        sql = "AND (CISCas.DateVisit Is null OR CISCas.DateVisit = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDORDate1)) > 0 And IsDate(txtDORDate1) = True Then
        strAppend = strAppend + " And CISCas.RegisterDate >= '" & Format(txtDORDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDORDate2)) > 0 And IsDate(txtDORDate2) = True Then
        strAppend = strAppend + " And CISCas.RegisterDate <= '" & Format(txtDORDate2, "mm/dd/yyyy") & "'"
    End If
    
    If chkDOR.Value = 1 Then
        sql = ""
        sql = "AND (CISCas.RegisterDate Is null OR CISCas.RegisterDate = '')" & ""
        strAppend = strAppend + sql
    End If
        
    If Len(Trim(txtInvDate1)) > 0 And IsDate(txtInvDate1) = True Then
        strAppend = strAppend + " And CISCas.InvoiceDate >= '" & Format(txtInvDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtInvDate2)) > 0 And IsDate(txtInvDate2) = True Then
        strAppend = strAppend + " And CISCas.InvoiceDate <= '" & Format(txtInvDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(cmdBordxDateFr)) > 0 And IsDate(cmdBordxDateFr) = True Then
        strAppend = strAppend + " And Bordx.BordxDate >= '" & Format(cmdBordxDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(cmdBordxDateTo)) > 0 And IsDate(cmdBordxDateTo) = True Then
        strAppend = strAppend + " And Bordx.BordxDate <= '" & Format(cmdBordxDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If chkInvDate.Value = 1 Then
        sql = ""
        sql = "AND (CISCas.InvoiceDate Is null OR CISCas.InvoiceDate = '')" & ""
        strAppend = strAppend + sql
    End If
        
    If Len(Trim(txtIRDDate1)) > 0 And IsDate(txtIRDDate1) = True Then
        strAppend = strAppend + " And ClaimInvoice.INVEnteredDate >= '" & Format(txtIRDDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtIRDDate2)) > 0 And IsDate(txtIRDDate2) = True Then
        strAppend = strAppend + " And ClaimInvoice.INVEnteredDate <= '" & Format(txtIRDDate2, "mm/dd/yyyy") & "'"
    End If
    
    If chkInvReg.Value = 1 Then
        sql = ""
        sql = "AND (ClaimInvoice.INVEnteredDate Is null OR ClaimInvoice.INVEnteredDate = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtINVAmtFr)) Then
        strAppend = strAppend + " AND ClaimInvoice.INVAmount >= " & Trim(txtINVAmtFr) & ""
    End If
    
    If Len(Trim(txtINVAmtTo)) Then
        strAppend = strAppend + " AND ClaimInvoice.INVAmount <= " & Trim(txtINVAmtTo) & ""
    End If
    
    If ChkInvAmt.Value = 1 Then
        sql = ""
        sql = "AND (ClaimInvoice.INVAmount Is null)" & ""
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
        strAppend = strAppend & " AND CISCas.CLNType = 'S'"
    ElseIf Mid(CboClinicType, 1, 1) = 1 Then
        strAppend = strAppend & " AND CISCas.CLNType = 'G'"
    ElseIf Mid(CboClinicType, 1, 1) = 2 Then
        strAppend = strAppend & " AND (CISCas.CLNType = '' OR CISCas.CLNType IS NULL)"
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
    
    If Len(Trim(cboOMNIStatus)) Then
        strAppend = strAppend + " AND GRPOMNICLM = 'Y'"
    End If
    
    If Len(Trim(cboTIMSInd)) Then
        strAppend = strAppend + " AND GRPTIMS = 'Y'"
    End If
    
    
    
    If Len(Trim(cboSubsidiary)) Then
        strAppend = strAppend + " AND MBMOthers.Subsidiary = '" & Trim(cboSubsidiary) & "'"
    End If
    
    If Len(Trim(CboDepartment)) Then
        strAppend = strAppend + " AND MBMOthers.MBMDepartment = '" & Trim(CboDepartment) & "'"
    End If
    If chkOMNIRef.Value = 1 Then
        sql = ""
        sql = "AND (OMNIRefNo Is null OR OMNIRefNo = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtOMNIRefFr)) Then
        strAppend = strAppend + " AND OMNIRefNo >= '" & Trim(txtOMNIRefFr) & "'"
    End If

    If Len(Trim(txtOMNIRefTo)) Then
        strAppend = strAppend + " AND OMNIRefNo <= '" & Trim(txtOMNIRefTo) & "'"
    End If
    
    If Mid(cboOMNISent, 1, 1) = "Y" Then
        strAppend = strAppend & " AND CLMOMNISent = 'Y'"
    ElseIf Mid(cboOMNISent, 1, 1) = "N" Then
        strAppend = strAppend & " AND CLMOMNISent IS NULL"
    End If
    
    If Mid(cboauditind, 1, 1) = "Y" Then
        strAppend = strAppend & " AND ciscas.AuditQuery IS NOT NULL "
    End If
    
    
    SQLqueryappend = ""
    SQLqueryappend = strAppend

    Call InvoiceTrackingListing(strAppend)
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
'************************End Search Record**************************

'************************Start List Record**************************

Private Sub InvoiceTrackingListing(Optional strAppend As String)
Const METHOD_NAME = "InvoiceTrackingListing"
'On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim TotAmt As Variant
Dim TotalInvNo As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
Dim tmpGPFee As Double
Dim tmpSPFee As Double
Dim tmpGPLTMed As Double
Dim tmpSPLTMed As Double
Dim tmpGPMed As Double
Dim tmpSPMed As Double
Dim tmpGPInj As Double
Dim tmpSPInj As Double


    Current = 0
    TotAmt = 0
    
    lstInvoiceTrackList.ListItems.Clear
    strAppendCase = "1"
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "Sp_Audittraccking_2018"
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
    
    Do
        Do Until rst2.EOF
            With rst2
                Set Record = lstInvoiceTrackList.ListItems.Add(, , IIf(Len(rst2!MBMIcBcPp), rst2!MBMIcBcPp, ""))
                
                If Len(rst2!MBMName) Then
                    Record.SubItems(1) = Trim(rst2!MBMName)
                End If
                
                If Len(rst2!CasNumber) Then
                    Record.SubItems(2) = Trim(rst2!CasNumber)
                End If
                
                If rst2!DateVisit = "01/01/1900" Then
                    Record.SubItems(3) = ""
                End If
                
                If rst2!DateVisit <> "01/01/1900" Then
                    Record.SubItems(3) = Format(rst2!DateVisit, "dd/mm/yyyy")
                End If
                
                If Len(rst2!MedInjType1) Then
                   Record.SubItems(4) = rst2!MedInjType1
                End If
                
                If Len(rst2!MedInjDays1) Then
                    Record.SubItems(5) = rst2!MedInjDays1
                End If
                
                If Len(rst2!MedInjQty1) Then
                    Record.SubItems(6) = rst2!MedInjQty1
                End If
                
                If Len(rst2!MedCharges1) Then
                    Record.SubItems(7) = Format(rst2!MedCharges1, "#,##0.00;(#,##0.00)")
                End If
                
                If Len(rst2!SPFee) Then
                    Record.SubItems(8) = Format(rst2!SPFee, "#,##0.00;(#,##0.00)")
                End If
                
                If Len(rst2!INVAmount) Then
                    Record.SubItems(9) = Format(rst2!INVAmount, "#,##0.00;(#,##0.00)")
                End If
                
                 If Len(rst2!CasRemark) Then
                   Record.SubItems(10) = rst2!CasRemark
                End If
                
                If Len(rst2!AuditQuery) Then
                   Record.SubItems(11) = rst2!AuditQuery
                End If
                
                If Len(rst2!AuditAnswer) Then
                   Record.SubItems(12) = rst2!AuditAnswer
                End If
                
                If Len(rst2!AuditRemarks) Then
                   Record.SubItems(13) = rst2!AuditRemarks
                End If
                
                If Len(rst2!AuditDiscountReceived) Then
                    Record.SubItems(14) = Format(rst2!AuditDiscountReceived, "#,##0.00;(#,##0.00)")
                End If
                
                If Len(rst2!ReasonDesc1) Then
                   Record.SubItems(15) = rst2!ReasonDesc1
                End If
                
                If Len(rst2!ReasonDesc2) Then
                   Record.SubItems(16) = rst2!ReasonDesc2
                End If
                
                If Len(rst2!ReasonDesc3) Then
                   Record.SubItems(17) = rst2!ReasonDesc3
                End If
                
                If Len(rst2!ReasonDesc4) Then
                   Record.SubItems(18) = rst2!ReasonDesc4
                End If
                
                If Len(rst2!CASDiagnosis5) Then
                    Record.SubItems(19) = rst2!CASDiagnosis5
                ElseIf Len(rst2!ReasonDesc5) Then
                    Record.SubItems(19) = rst2!ReasonDesc5
                End If
                
                If Len(rst2!AuditBy) Then
                   Record.SubItems(20) = rst2!AuditBy
                End If
                
                If Len(rst2!AuditDate) Then
                   Record.SubItems(21) = rst2!AuditDate
                End If
                
                If Len(rst2!CLNCode) Then
                   Record.SubItems(22) = rst2!CLNCode
                End If
                
                If Len(rst2!CLNName) Then
                   Record.SubItems(23) = rst2!CLNName
                End If
                
                If Len(rst2!MBMGroupCompany) Then
                   Record.SubItems(24) = rst2!MBMGroupCompany
                End If
                
            End With
            rst2.MoveNext
            DoEvents
        Loop
    Loop Until Check = False
    
    If lstInvoiceTrackList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdClear.Enabled = True
        CmdPrintFile.Enabled = True
        cmdExit.Enabled = True
        CmdSearch.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    cmdClear.Enabled = True
    CmdPrintFile.Enabled = True
    cmdExit.Enabled = True
    CmdSearch.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'************************End List Record**************************

Private Sub PaymentInfo(tmpCASnumber As String)
Const METHOD_NAME = "PaymentInfo"
On Error GoTo debugs
Dim StrSql As String
Dim PayRec As New ADODB.Recordset
Dim PayList As ListItem

    
    tmpPVNo = ""
    tmpPVDate = ""
    tmpChqNo = ""
    tmpChqDate = ""
    
    StrSql = "SELECT PVNo, PVDate, PVChequeNo, PVChequeDate " & _
             "From Payment2MBM WHERE CASNumber='" & Trim(tmpCASnumber) & "'"
    PayRec.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockOptimistic
    Do While Not PayRec.EOF
        
        
        tmpPVNo = IIf(Len(PayRec!PVNo), PayRec!PVNo, "")
        tmpPVDate = IIf(Len(PayRec!PVDate), PayRec!PVDate, "")
        tmpChqNo = IIf(Len(PayRec!PVChequeNo), PayRec!PVChequeNo, "")
        tmpChqDate = IIf(Len(PayRec!PVChequeDate), PayRec!PVChequeDate, "")
        
        PayRec.MoveNext
    Loop
    PayRec.Close
    Set PayRec = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

'************************Start ComboBox Change/KeyPress**************************

Private Sub CboPayTo_Click()
    If InStr(CboPayto.Text, "null") Then
       CboPayto.Enabled = False
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
    If InStr(cboClaimability.Text, "null") Then
       cboClaimability.Enabled = False
    End If
End Sub

Private Sub CboClaimability_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboClaimability.Text, cboClaimability.SelStart) + Chr(KeyAscii) + Mid(cboClaimability.Text, cboClaimability.SelStart + cboClaimability.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboClaimability.ListCount - 1
 
        If NewText = LCase(Left(cboClaimability.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboClaimability.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboClaimability.Text = ValidValue
                cboClaimability.SelStart = 0
                cboClaimability.SelLength = Len(ValidValue)
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
    If InStr(CboPayto.Text, "null") Then
       CboPayto.Enabled = False
    End If
End Sub

Private Sub CboPayTo_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayto.Text, CboPayto.SelStart) + Chr(KeyAscii) + Mid(CboPayto.Text, CboPayto.SelStart + CboPayto.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayto.ListCount - 1
 
        If NewText = LCase(Left(CboPayto.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayto.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPayto.Text = ValidValue
                CboPayto.SelStart = 0
                CboPayto.SelLength = Len(ValidValue)
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

Private Sub lstInvoiceTrackList_DblClick()
    If lstInvoiceTrackList.ListItems.Count > 0 Then
        If Trim(lstInvoiceTrackList.SelectedItem.SubItems(2)) = "" Then
            MsgBox "Please select Correct Row", vbCritical
        Else
            frmInvoiceAdj.txtCaseNo = Trim(lstInvoiceTrackList.SelectedItem.SubItems(2))
            Call frmInvoiceAdj.CmdGo_Click
            Load frmInvoiceAdj
            frmInvoiceAdj.Show
        End If
    End If
End Sub

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
Const METHOD_NAME = "ChkClinic2_Click"
On Error GoTo debugs
Dim CLN As New ADODB.Recordset
Dim StrSql As String

    CboClinic.Clear
    If ChkClinic2.Value = 1 Then
        StrSql = "EXEC SearchCLN"
        CLN.Open StrSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic
    
        Do Until CLN.EOF
            CboClinic.AddItem CLN!CLNCode & " - " & CLN!CLNName
            CLN.MoveNext
        Loop
        CLN.Close
        Set CLN = Nothing
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub ChkGrpCom2_Click()
Const METHOD_NAME = "ChkGrpCom2_Click"
On Error GoTo debugs
Dim StrSql As String
Dim PAYGRP As New ADODB.Recordset

    CboGrpCom.Clear
    If ChkGrpCom2.Value = 1 Then
       StrSql = "EXEC SearchPayGrp '" & Mid(CboPayor, 1, 2) & "'"
        PAYGRP.Open StrSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic
        
        If Not PAYGRP.EOF Then
            Do Until PAYGRP.EOF
               CboGrpCom.AddItem Trim(PAYGRP!GrpCoName)
               PAYGRP.MoveNext
            Loop
        End If
        PAYGRP.Close
        Set PAYGRP = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
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
