VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInvoiceTracking 
   Caption         =   "Invoice Tracking"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   4875
      Left            =   0
      TabIndex        =   71
      Top             =   240
      Width           =   11775
      Begin VB.ComboBox cbopaymentstatus 
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
         Style           =   2  'Dropdown List
         TabIndex        =   154
         Top             =   4440
         Width           =   855
      End
      Begin VB.ComboBox CboFlexOP 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking.frx":0000
         Left            =   1200
         List            =   "FrmInvoiceTracking.frx":000A
         TabIndex        =   147
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkCaseStatus 
         Caption         =   "&Closed"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   4000
         TabIndex        =   146
         Tag             =   "C"
         Top             =   3195
         Width           =   975
      End
      Begin VB.CheckBox chkCaseStatus 
         Caption         =   "&Deleted"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   3000
         TabIndex        =   145
         Tag             =   "D"
         Top             =   3195
         Width           =   975
      End
      Begin VB.CheckBox chkCaseStatus 
         Caption         =   "&Pending"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   1950
         TabIndex        =   144
         Tag             =   "P"
         Top             =   3195
         Width           =   975
      End
      Begin VB.CheckBox chkCaseStatus 
         Caption         =   "&Open"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   143
         Tag             =   "O"
         Top             =   3195
         Width           =   735
      End
      Begin VB.ComboBox cboEntityType 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking.frx":0024
         Left            =   1200
         List            =   "FrmInvoiceTracking.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   142
         Top             =   4260
         Width           =   1575
      End
      Begin VB.ComboBox CboDepartment 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking.frx":004F
         Left            =   7560
         List            =   "FrmInvoiceTracking.frx":0051
         Sorted          =   -1  'True
         TabIndex        =   135
         Top             =   4170
         Width           =   3615
      End
      Begin VB.CheckBox chkOMNIRef 
         Height          =   255
         Left            =   5475
         TabIndex        =   121
         Top             =   3750
         Width           =   255
      End
      Begin VB.CheckBox ChkPayor2 
         Height          =   255
         Left            =   5760
         TabIndex        =   112
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
         Top             =   2325
         Width           =   255
      End
      Begin VB.CheckBox ChkInvNo 
         Height          =   255
         Left            =   5480
         TabIndex        =   21
         Top             =   2595
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
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox ChkName 
         Height          =   255
         Left            =   5480
         TabIndex        =   15
         Top             =   2040
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
         Top             =   1995
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
         Top             =   2280
         Width           =   1935
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
         MaxLength       =   50
         TabIndex        =   19
         Top             =   2565
         Width           =   1935
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
         Top             =   2850
         Width           =   1935
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
         Left            =   5520
         TabIndex        =   25
         Top             =   4260
         Visible         =   0   'False
         Width           =   735
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
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   17
         Top             =   2280
         Width           =   1815
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   20
         Top             =   2565
         Width           =   1815
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
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   23
         Top             =   2850
         Width           =   1815
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
         ItemData        =   "FrmInvoiceTracking.frx":0053
         Left            =   3840
         List            =   "FrmInvoiceTracking.frx":0055
         TabIndex        =   26
         Top             =   3420
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
         Top             =   3420
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
         ItemData        =   "FrmInvoiceTracking.frx":0057
         Left            =   7560
         List            =   "FrmInvoiceTracking.frx":0059
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
         ItemData        =   "FrmInvoiceTracking.frx":005B
         Left            =   7560
         List            =   "FrmInvoiceTracking.frx":005D
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
         ItemData        =   "FrmInvoiceTracking.frx":005F
         Left            =   9720
         List            =   "FrmInvoiceTracking.frx":0061
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
         ItemData        =   "FrmInvoiceTracking.frx":0063
         Left            =   9720
         List            =   "FrmInvoiceTracking.frx":0065
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
         TabIndex        =   115
         Top             =   4260
         Width           =   1570
      End
      Begin VB.ComboBox cboOMNIStatus 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking.frx":0067
         Left            =   7560
         List            =   "FrmInvoiceTracking.frx":0077
         TabIndex        =   117
         Top             =   3300
         Width           =   3615
      End
      Begin VB.ComboBox cboSubsidiary 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking.frx":00BB
         Left            =   7560
         List            =   "FrmInvoiceTracking.frx":00BD
         Sorted          =   -1  'True
         TabIndex        =   119
         Top             =   3570
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
         TabIndex        =   122
         Top             =   3705
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
         TabIndex        =   123
         Top             =   3705
         Width           =   1570
      End
      Begin MSMask.MaskEdBox cmdBordxDateFr 
         Height          =   315
         Left            =   7560
         TabIndex        =   128
         Top             =   3880
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
         TabIndex        =   129
         Top             =   3880
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
         ItemData        =   "FrmInvoiceTracking.frx":00BF
         Left            =   3840
         List            =   "FrmInvoiceTracking.frx":00CF
         TabIndex        =   132
         Top             =   3990
         Width           =   1575
      End
      Begin VB.ComboBox cboTIMSInd 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking.frx":0113
         Left            =   1200
         List            =   "FrmInvoiceTracking.frx":0123
         TabIndex        =   126
         Top             =   3990
         Width           =   1575
      End
      Begin MSMask.MaskEdBox MedopbatchDateFr 
         Height          =   315
         Left            =   7560
         TabIndex        =   150
         Top             =   3000
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
      Begin MSMask.MaskEdBox MedopbatchDateto 
         Height          =   315
         Left            =   9720
         TabIndex        =   152
         Top             =   3000
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
      Begin VB.Label Label32 
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
         Left            =   6360
         TabIndex        =   153
         Top             =   4440
         Width           =   1215
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
         Index           =   1
         Left            =   9240
         TabIndex        =   151
         Top             =   3050
         Width           =   375
      End
      Begin VB.Label lblopbatchdate 
         Caption         =   "OP Batch Date"
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
         TabIndex        =   149
         Top             =   3050
         Width           =   1095
      End
      Begin VB.Label LblFlexOP 
         Caption         =   "flexCare / OP"
         Height          =   255
         Left            =   120
         TabIndex        =   148
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label lblentitytype 
         Caption         =   "Entity Type"
         Height          =   255
         Left            =   120
         TabIndex        =   141
         Top             =   4290
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
         TabIndex        =   134
         Top             =   4170
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
         TabIndex        =   133
         Top             =   4020
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
         TabIndex        =   131
         ToolTipText     =   "Date of Visit"
         Top             =   3880
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
         Left            =   9120
         TabIndex        =   130
         Top             =   3880
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
         TabIndex        =   127
         Top             =   4020
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
         TabIndex        =   125
         Top             =   3735
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
         TabIndex        =   124
         Top             =   3735
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
         TabIndex        =   120
         Top             =   3570
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
         TabIndex        =   118
         Top             =   3300
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
         TabIndex        =   116
         Top             =   4305
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
         TabIndex        =   111
         Top             =   3465
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
         TabIndex        =   110
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
         Index           =   0
         Left            =   9240
         TabIndex        =   109
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
         TabIndex        =   107
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
         TabIndex        =   106
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
         TabIndex        =   105
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
         TabIndex        =   104
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
         TabIndex        =   103
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         TabIndex        =   100
         Top             =   3465
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
         TabIndex        =   99
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
         TabIndex        =   98
         Top             =   2400
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
         TabIndex        =   97
         Top             =   2040
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
         TabIndex        =   96
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
         TabIndex        =   95
         Top             =   2910
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
         TabIndex        =   94
         Top             =   2640
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
         TabIndex        =   93
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
         TabIndex        =   92
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
         TabIndex        =   91
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
         TabIndex        =   90
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
         TabIndex        =   89
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
         TabIndex        =   88
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
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   85
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   80
         Top             =   3165
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
         TabIndex        =   79
         Top             =   2595
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
         TabIndex        =   78
         Top             =   2355
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
         TabIndex        =   77
         Top             =   2910
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
         Top             =   1490
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstInvoiceTrackList 
      Height          =   1665
      Left            =   0
      TabIndex        =   70
      Top             =   5160
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   2937
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
      NumItems        =   82
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
         SubItemIndex    =   3
         Text            =   "BTBG Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Inv. No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Inv. Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Inv. Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Excess Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Net Amt (Inv. Amt - Excess Amt)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "CoyPayment"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Approved Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "BTB Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Non GP Fee"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "FFS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "DOR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "DOV"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "No of Days (DOR - DOV)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Payee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Clinic Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   19
         Text            =   "Clinic Type"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Membership No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   22
         Text            =   "Plan Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   23
         Text            =   "Insured Type"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Group Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   25
         Text            =   "Inv. Reg"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Inv. Entry By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "OP Batch No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   28
         Text            =   "Case Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   29
         Text            =   "Claimability"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   30
         Text            =   "Bordx Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "Reason for Treatment"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   32
         Text            =   "A/C Period"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "MC Days"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "Reason for Treatment1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "Reason for Treatment2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "Reason for Treatment3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Text            =   "Reason for Treatment4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Text            =   "Reason for Treatment5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Text            =   "OP Enter By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   40
         Text            =   "FFS Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Text            =   "Subsidiary"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Text            =   "Last Update User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Text            =   "Last Update Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Text            =   "PV No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Text            =   "PV Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Text            =   "Cheque No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Text            =   "Cheque Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Text            =   "OMNI Ref No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Text            =   "Medication 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Text            =   "Medication 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Text            =   "Medication 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   52
         Text            =   "Medication 4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   53
         Text            =   "Medication 5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   54
         Text            =   "Medication 6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   55
         Text            =   "Medication 7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   56
         Text            =   "Med Days 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   57
         Text            =   "Med Days 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   58
         Text            =   "Med Days 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   59
         Text            =   "Med Days 4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   60
         Text            =   "Med Days 5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   61
         Text            =   "Med Qty 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   62
         Text            =   "Med Qty 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(64) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   63
         Text            =   "Med Qty 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(65) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   64
         Text            =   "Med Qty 4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(66) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   65
         Text            =   "Med Qty 5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(67) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   66
         Text            =   "Consultation Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(68) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   67
         Text            =   "Medication Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(69) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   68
         Text            =   "Injection Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(70) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   69
         Text            =   "OMNI Sent"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(71) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   70
         Text            =   "CAS Mgmt Notes"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(72) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   71
         Text            =   "Cas Remark"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(73) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   72
         Text            =   "Doctor Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(74) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   73
         Text            =   "Patient IC No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(75) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   74
         Text            =   "CoverID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(76) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   75
         Text            =   "OP Batch Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(77) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   76
         Text            =   "Recording PIC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(78) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   77
         Text            =   "Preliminary Diagnosis"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(79) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   78
         Text            =   "Claim subsidiary"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(80) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   79
         Text            =   "Claim cost center"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(81) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   80
         Text            =   "Claim Dept Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(82) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   81
         Text            =   "LG Internal Remarks"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   7000
      Width           =   11775
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   138
         Top             =   195
         Width           =   1095
      End
      Begin VB.TextBox LblAppAmt 
         Height          =   285
         Left            =   3480
         TabIndex        =   137
         Top             =   195
         Width           =   1095
      End
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   5880
         TabIndex        =   136
         Top             =   195
         Width           =   1215
      End
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
         Top             =   2040
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
      Begin VB.TextBox txtTotalBTBGAmt 
         Height          =   285
         Left            =   3480
         TabIndex        =   140
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "Total BTB Amt"
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
         Left            =   2400
         TabIndex        =   139
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "Total App Amt"
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
         Left            =   2400
         TabIndex        =   114
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   840
         TabIndex        =   69
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label29 
         Caption         =   "Total Inv Amt"
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
         Left            =   4680
         TabIndex        =   68
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1680
         TabIndex        =   67
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
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
         Left            =   1560
         TabIndex        =   66
         Top             =   195
         Width           =   735
      End
   End
   Begin VB.Label Label16 
      Caption         =   "*Double-Click to see claim details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   113
      Top             =   6840
      Width           =   6015
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Invoice  Tracking"
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
      TabIndex        =   108
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmInvoiceTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Private m_blnDirection(1 To 82) As Boolean
Dim tmpPVNo As String
Dim tmpPVDate As String
Dim tmpChqNo As String
Dim tmpChqDate As String



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

    'PAY.Open "Select PAYCode, PAYCompanyName from PAY", MediNeptuneConnCISDB, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
    CboPayor.Clear
    If ChkPayor2.Value = 1 Then
    
       ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
       ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
        StrSql = "EXEC SearchPayor"
        PAY.Open StrSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
        
        Do Until PAY.EOF
            CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
            PAY.MoveNext
        Loop
        PAY.Close
        Set PAY = Nothing
        
       ' MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
       ' MediConnCISWS.Close
       ' Set MediConnCISWS = Nothing
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
    Call ActionLogUpdate(12, "FrmInvoiceTracking")
    intExit = 0
    lblCurrent = 0
    LblTotalAmt = 0
    Call ComboListing
    cboEntityType.ListIndex = 0
    '# Added by Ravi
    CboFlexOP.Enabled = False
    LblFlexOP.Enabled = False
    '# Added by Ravi
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
    CboFlexOP.Enabled = False
    LblFlexOP.Enabled = False
    
End Sub

Private Sub cmdStop_Click()
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
    
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
        Call SearchRecord
    
   ' MediConnCISDB.Close
  '  Set MediConnCISDB = Nothing
    'medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' MediConnCISWS.Close
   ' Set MediConnCISWS = Nothing
    
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
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
    Case 12
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    Case 13
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
    Case 14
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
    Case 15
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(15)
    Case 16
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(16)
    Case 17
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
    Case 18
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 19
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(19)
    Case 20
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(21)
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
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(28)
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
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(34)
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
    Case 40
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(40)
    Case 41
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(41)
    Case 42
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(42)
    Case 43
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(43)
    Case 44
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(44)
    Case 45
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(45)
    Case 46
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(46)
    Case 47
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(47)
    Case 48
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(48)
    Case 49
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(49)
    Case 50
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(50)
    Case 51
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(51)
    Case 52
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(52)
    Case 53
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(53)
    Case 54
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(54)
    Case 55
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(55)
    Case 56
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(56)
    Case 57
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(57)
    Case 58
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(58)
    Case 59
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(59)
    Case 60
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(60)
    Case 61
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(61)
    Case 62
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(62)
    Case 63
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(63)
    Case 64
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(64)
    Case 65
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(65)
    Case 66
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(66)
    Case 67
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(67)
    Case 68
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(68)
    Case 69
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(69)
    Case 70
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(70)
    Case 71
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(71)
    Case 72
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(72)
    Case 73
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(73)
    Case 74
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(74)
    Case 75
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(75)
    Case 76
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(76)
    Case 77
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(77)
    Case 78
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(78)
        
    Case 79
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(79)
    Case 80
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(80)
    Case 81
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(81)
    Case 82
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(82)
    End Select
End Sub
'************************End sorted lstInvoiceTrackList**************************

'************************Start ComboListing**************************
Private Sub ComboListing()
Dim sql As String
Dim yearstart
Dim monthstart
   
    With CboPayto
        .AddItem "C - Clinic"
        .AddItem "M - Member"
    End With
    
    With cbopaymentstatus
        .AddItem ""
        .AddItem "Yes"
        .AddItem "No"
    End With
    
    With cboClaimability
        .AddItem "A - Accepted"
        .AddItem "R - Rejected"
    End With
    
    With CboCaseStatus
        .AddItem "O - Open"
        .AddItem "C - Close"
        .AddItem "D - Delete"
    'Ravi
        .AddItem "P - Pending"
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
    
    cboOMNISent.AddItem "Y - Yes"
    cboOMNISent.AddItem "N - No"
End Sub
'************************Start ComboListing**************************

'************************Start Search Record**************************

Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs
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
    strAppend = strAppend + " AND ClinicCode = '" & Trim(SArrccode) & "'"
        'strsqlHOS = "select CLNCode, CLNName from CLN where CLNCode = '" & Trim(RemStr(CboClinic)) & "'"
        '=========================trim CLNCLinic===============
        
'        strsqlHOS = "select CLNCode, CLNName from CLN where CLNCode = '" & Mid(CboClinic, 1, 4) & "'"
'        CLN.Open strsqlHOS, medixmasterconn, adOpenForwardOnly, adLockReadOnly
'        If CLN.EOF = False Then
'            strAppend = strAppend + " AND ClinicCode = '" & Trim(CLN!CLNCode) & "'" 'INVPayee
'        End If
'        CLN.Close
'        Set CLN = Nothing


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
    
    If Len(Trim(Replace(CboGrpCom, "'", "''"))) Then
       strAppend = strAppend + " AND MBMGroupCompany = '" & Trim(RemStr(CboGrpCom)) & "'"
    End If
    
    If ChkGrpCom.Value = 1 Then
       strAppend = strAppend + " AND (MBMGroupCompany IS NULL OR MBMGroupCompany = '')"
    End If
    
    If Len(Trim(txtPatientName)) Then
       strAppend = strAppend + " And PatientName = '" & Trim(txtPatientName) & "'"
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
    
    If Len(Trim(txtInvNo1)) Then
        strAppend = strAppend + " AND BillNo >= '" & Trim(txtInvNo1) & "'"
    End If

    If Len(Trim(txtInvNo2)) Then
        strAppend = strAppend + " AND BillNo <= '" & Trim(txtInvNo2) & "'"
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
        
    '*******START: Edited by SRI, Date: 26/04/2019 3:30PM*********'
    '*Case Status selection changed to Checkboxes to enable multiple selection*'
    '*CboCaseStatus is temporarily invisible, once UAT succesful, we can delete the old code and CboCaseStatus control*'
    
    'Start: @Previous Code'
    'If Mid(CboCaseStatus, 1, 1) = "C" Then
    '    strAppend = strAppend + " AND CASStatus = 'C'"
    'ElseIf Mid(CboCaseStatus, 1, 1) = "O" Then
    '    strAppend = strAppend + " AND CASStatus = 'O'"
    'ElseIf Mid(CboCaseStatus, 1, 1) = "D" Then
    '    strAppend = strAppend + " AND CASStatus = 'D'"
        'Ravi
    'ElseIf Mid(CboCaseStatus, 1, 1) = "P" Then
    '    strAppend = strAppend + " AND Pending = 'Y'"
    'End If
    'End: @Previous Code'
    
    Dim CaseStatusIdx As Integer
    Dim CaseStatusSelected As String
    For CaseStatusIdx = 0 To 2
        If ChkCaseStatus(CaseStatusIdx).Value = 1 Then
            CaseStatusSelected = CaseStatusSelected + "'" + ChkCaseStatus(CaseStatusIdx).Tag + "',"
        End If
    Next CaseStatusIdx
    
    If (CaseStatusSelected <> "") Then
        CaseStatusSelected = Mid(CaseStatusSelected, 1, Len(CaseStatusSelected) - 1)
        CaseStatusSelected = " AND CASStatus in (" + CaseStatusSelected + ") "
        strAppend = strAppend + CaseStatusSelected
    End If
    
    If ChkCaseStatus(3).Value = 1 Then
        strAppend = strAppend + " AND Pending = 'Y'"
    End If
    
    '*******END: Edited by SRI, Date: 26/04/2019 4:16PM*********'
    
    If Mid(cboClaimability, 1, 1) = "R" Then
        strAppend = strAppend + " Or Claimability = 'R'"
    ElseIf Mid(cboClaimability, 1, 1) = "A" Then
        strAppend = strAppend + " Or Claimability = 'A'"
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
    
    If Len(Trim(cmdBordxDateFr)) > 0 And IsDate(cmdBordxDateFr) = True Then
        strAppend = strAppend + " And BordxDate >= '" & Format(cmdBordxDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(cmdBordxDateTo)) > 0 And IsDate(cmdBordxDateTo) = True Then
        strAppend = strAppend + " And BordxDate <= '" & Format(cmdBordxDateTo, "mm/dd/yyyy") & "'"
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
        strAppend = strAppend + " And cast(RegisterDate as date) >= '" & Format(txtDORDate1, "YYYY-MM-DD") & "'"
    End If
    
    If Len(Trim(txtDORDate2)) > 0 And IsDate(txtDORDate2) = True Then
        strAppend = strAppend + " And cast(RegisterDate as date) <= '" & Format(txtDORDate2, "YYYY-MM-DD") & "'"
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
    
    If Len(Trim(cmdBordxDateFr)) > 0 And IsDate(cmdBordxDateFr) = True Then
        strAppend = strAppend + " And BordxDate >= '" & Format(cmdBordxDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(cmdBordxDateTo)) > 0 And IsDate(cmdBordxDateTo) = True Then
        strAppend = strAppend + " And BordxDate <= '" & Format(cmdBordxDateTo, "mm/dd/yyyy") & "'"
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
    
    If Len(Trim(MedopbatchDateFr)) And IsDate(MedopbatchDateFr) = True Then
        strAppend = strAppend + " AND OPBatchDate >= '" & Format(MedopbatchDateFr, "mm/dd/yyyy") & "'"
    End If
    '
    If Len(Trim(MedopbatchDateto)) And IsDate(MedopbatchDateto) = True Then
        strAppend = strAppend + " AND OPBatchDate <= '" & Format(MedopbatchDateto, "mm/dd/yyyy") & "'"
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
        ''strAppend = strAppend + " AND Subsidiary = '" & Trim(cboSubsidiary) & "'"
        strAppend = strAppend + " AND clmsubsidiary = '" & Trim(cboSubsidiary) & "'"
    End If
    
    If Len(Trim(CboDepartment)) Then
        ''strAppend = strAppend + " AND MBMDepartment = '" & Trim(CboDepartment) & "'"
        strAppend = strAppend + " AND CLMDeptName = '" & Trim(CboDepartment) & "'"
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
        'strAppend = strAppend & " AND CLMOMNISent = 'N' OR (CLMOMNISent = '' OR CLMOMNISent IS NULL)"
        strAppend = strAppend & " AND CLMOMNISent IS NULL"
    End If
    
     If cboEntityType.ListIndex = 0 Then
        strAppend = strAppend + " AND Substring(CasNumber, 1, 2) <>'SJ' "
    ElseIf cboEntityType.ListIndex = 1 Then
        strAppend = strAppend + " AND Substring(CasNumber, 1, 2) ='SJ' "
    End If
    
    '# Added by Ravi
    If CboFlexOP.ListIndex = 0 Then
        strAppend = strAppend + " AND (FBEmployee is null or (FBFamily =1 and CoverID<>'00'))"
    ElseIf CboFlexOP.ListIndex = 1 Then
        strAppend = strAppend + " AND (FBEmployee is not null and CoverID='00') "
    End If
    '# Added by Ravi
    
    If Mid(cbopaymentstatus, 1, 1) = "Y" And Mid(CboPayto, 1, 1) = "C" Then
       strAppend = strAppend + " AND PVCLNNo IS NOT NULL "
    ElseIf Mid(cbopaymentstatus, 1, 1) = "Y" And Mid(CboPayto, 1, 1) = "M" Then
       strAppend = strAppend + " AND PVNo IS NOT NULL "
    ElseIf Mid(cbopaymentstatus, 1, 1) = "N" And Mid(CboPayto, 1, 1) = "C" Then
        strAppend = strAppend + " AND PVCLNNo IS NULL "
    ElseIf Mid(cbopaymentstatus, 1, 1) = "N" And Mid(CboPayto, 1, 1) = "M" Then
        strAppend = strAppend + " AND PVNo IS NULL "
    End If
    
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
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotAmt As Variant
Dim TotBTBGAmt As Variant
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
    LblTotalAmt = ""
    LblAppAmt = 0
    lstInvoiceTrackList.ListItems.Clear
    
    strAppendCase = "25"
            
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimBordxModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 900
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
            
            'lblTotal = rst2.RecordCount
            'Start: Murali comment this part
'            If Trim(rst2!INVPayTo) = "M" Then
'                Call PaymentInfo(rst2!CasNumber)
'            Else
'
'                tmpPVNo = IIf(Len(rst2!PVCLNNo), rst2!PVCLNNo, "")
'                tmpPVDate = IIf(Len(rst2!PVCLNDate), rst2!PVCLNDate, "")
'                tmpChqNo = IIf(Len(rst2!PVCLNChequeNo), rst2!PVCLNChequeNo, "")
'                tmpChqDate = IIf(Len(rst2!PVCLNChequeDate), rst2!PVCLNChequeDate, "")
'            End If
            'End: Murali comment this part
            
            With rst2
                
                Set Record = lstInvoiceTrackList.ListItems.Add(, , rst2!CasNumber)
                
                If Len(rst2!Bordxrefno) Then
                    Record.SubItems(1) = Trim(rst2!Bordxrefno)
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(2) = Trim(rst2!BordxDate)
                End If
                If Len(rst2!BTBGBordxRefNo) Then
                    Record.SubItems(3) = Trim(rst2!BTBGBordxRefNo)
                End If
                If Len(rst2!BillNo) Then
                    Record.SubItems(4) = Trim(rst2!BillNo)
                End If
                If Len(rst2!BillDate) Then
                   Record.SubItems(5) = Format(rst2!BillDate, "dd/mm/yyyy")
                End If
                If Len(rst2!INVAmount) Then
                    Record.SubItems(6) = Format(rst2!INVAmount, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                If Len(rst2!INVPaidByMem) Then
                    Record.SubItems(7) = Format(rst2!INVPaidByMem, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(7) = "0.00"
                End If
                
                If Len(rst2!NetAmt) Then
                    Record.SubItems(8) = Format(rst2!NetAmt, "#,##0.00")
                Else
                    Record.SubItems(8) = "0.00"
                End If
                
                If Len(rst2!CLMCoPayCLNColl) Then
                    Record.SubItems(9) = Format(rst2!CLMCoPayCLNColl, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(9) = "0.00"
                End If

                'If Len(rst2!CLMFinalApprove) Then
                '    Record.SubItems(9) = Format(rst2!CLMFinalApprove, "#,##0.00;(#,##0.00)")
                'Else
                '    Record.SubItems(9) = "0.00"
                'End If
                                
                
                If Len(rst2!CLMBordxAmt) Then
                    Record.SubItems(10) = Format(rst2!CLMBordxAmt, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(10) = "0.00"
                End If
                
                
                
                If Len(rst2!CLMBTBGAmt) Then
                    Record.SubItems(11) = Format(rst2!CLMBTBGAmt, "#,##0.00")
                Else
                    Record.SubItems(11) = "0.00"
                End If
                
                'Dim Copay As Double
                'Dim FFS As Double
                'If Len(rst2!CLMFinalApprove) Then
                '  If Len(rst2!CLMCoPayCLNColl) Then
                '
                '    If InStr(1, rst2!CLMCoPayCLNColl, "-") Then
                '        Copay = rst2!CLMCoPayCLNColl
                '    Else
                '        Copay = (CDbl(rst2!CLMCoPayCLNColl) * -1)
                '    End If
                '
                '    If Len(rst2!CLMMCOProcFee) Then
                ''        FFS = CDbl(rst2!CLMMCOProcFee) * -1
                 '   Else
                 '       FFS = 0
                 '   End If
                 '
                 '   If Len(rst2!INVAmount) Then
                 '    Record.SubItems(9) = Format(CDbl(rst2!INVAmount) + Copay + FFS, "#,##0.00;(#,##0.00)")
                 '   End If
                 '
                 'Else
                 'Record.SubItems(9) = Format(rst2!CLMFinalApprove, "#,##0.00;(#,##0.00)")
                 'End If
                'else
                '    Record.SubItems(9) = "0.00"
                'End If
                
                
                If Len(rst2!ClmNonGPFee) Then
                    Record.SubItems(12) = Format(rst2!ClmNonGPFee, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(12) = "0.00"
                End If
                
                If Len(rst2!CLMMCOProcFee) Then
                    Record.SubItems(13) = Format(rst2!CLMMCOProcFee, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(13) = "0.00"
                End If
                
                If Len(rst2!RegisterDate) Then
                    Record.SubItems(14) = Format(rst2!RegisterDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!DateVisit) Then
                    Record.SubItems(15) = Format(rst2!DateVisit, "dd/mm/yyyy")
                End If
                
                If rst2!NoofDays = "0" Then
                    Record.SubItems(16) = "1"
                ElseIf Len(rst2!NoofDays) Then
                    Record.SubItems(16) = rst2!NoofDays
                End If
                
                If Len(rst2!INVPayee) Then
                    Record.SubItems(17) = rst2!INVPayee
                End If
                                                                      
                If Trim(rst2!INVPayTo) = "C" Then
                    If Len(rst2!CLNName) Then
                        Record.SubItems(18) = Trim(rst2!CLNName)
                    End If
                Else
                
                    If Len(rst2!NPClinic) Then
                        Record.SubItems(18) = Trim(rst2!NPClinic)
                    End If
                End If
                    
                
                If Len(rst2!CLNType) Then
                    Record.SubItems(19) = Trim(rst2!CLNType)
                End If
                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(20) = Trim(rst2!MBMNumber)
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(21) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!PLNCode) Then
                    Record.SubItems(22) = Trim(rst2!PLNCode)
                End If
                
                If Len(rst2!INSCode) Then
                    Record.SubItems(23) = Trim(rst2!INSCode)
                End If
                
                If Len(rst2!MBMGroupCompany) Then
                    Record.SubItems(24) = Trim(rst2!MBMGroupCompany)
                End If
                
                If Len(rst2!INVEnteredDate) Then
                    Record.SubItems(25) = Format(rst2!INVEnteredDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!INVEnteredBy) Then
                    Record.SubItems(26) = Trim(rst2!INVEnteredBy)
                End If
                
                If Len(rst2!OPBatchNo) Then
                    Record.SubItems(27) = Trim(rst2!OPBatchNo)
                End If
                
                If Len(rst2!CASStatus) Then
                    Record.SubItems(28) = Trim(rst2!CASStatus)
                End If
                
                If Len(rst2!Claimability) Then
                    Record.SubItems(29) = Trim(rst2!Claimability)
                End If
                
                If Len(rst2!ClaimBordxStatus) Then
                    Record.SubItems(30) = Trim(rst2!ClaimBordxStatus)
                End If
                
                'If Len(rst2!ReasonDesc) Then
                '    Record.SubItems(31) = Trim(rst2!ReasonDesc)
                'End If
                
                If Len(rst2!AccPeriodMonth) Or (Len(rst2!AccPeriodYear)) Then
                    Record.SubItems(32) = rst2!AccPeriodMonth & "/" & rst2!AccPeriodYear
                End If
                 
                If Len(rst2!MCDays) Then
                    Record.SubItems(33) = Trim(rst2!MCDays)
                End If
                
                If Len(rst2!ReasonDesc1) Then
                    Record.SubItems(34) = Trim(rst2!ReasonDesc1)
                End If
                
                If Len(rst2!ReasonDesc2) Then
                    Record.SubItems(35) = Trim(rst2!ReasonDesc2)
                End If
                
                If Len(rst2!ReasonDesc3) Then
                    Record.SubItems(36) = Trim(rst2!ReasonDesc3)
                End If
                
                If Len(rst2!ReasonDesc4) Then
                    Record.SubItems(37) = Trim(rst2!ReasonDesc4)
                End If
                
                If Len(rst2!CASTreatReason5) Then
                    If Trim(rst2!CASTreatReason5) <> "199" And Trim(rst2!CASTreatReason5) <> "299" Then
                        If Len(Trim(rst2!ReasonDesc5)) Then
                        Record.SubItems(38) = Trim(rst2!ReasonDesc5)
                        End If
                    Else
                        Record.SubItems(38) = Trim(rst2!CASDiagnosis5)
                    End If
                End If
                
                Record.SubItems(39) = IIf(Len(rst2!OPNoEnterUser), rst2!OPNoEnterUser, "")
                
                Record.SubItems(40) = IIf(Len(rst2!CLNFFSStatus), rst2!CLNFFSStatus, "")
                
                Record.SubItems(41) = IIf(Len(rst2!Subsidiary), rst2!Subsidiary, "")
                
                Record.SubItems(42) = IIf(Len(rst2!ClaimLastUpdateUser), rst2!ClaimLastUpdateUser, "")
                
                Record.SubItems(43) = IIf(Len(rst2!ClaimLastUpdateDate), rst2!ClaimLastUpdateDate, "")
                                
                Record.SubItems(44) = IIf(Len(rst2!PVCLNNo), Trim(rst2!PVCLNNo), "") 'IIf(Len(tmpPVNo), tmpPVNo, "")
                
                Record.SubItems(45) = IIf(Len(rst2!PVCLNDate), Format(PVCLNDate, "dd/mm/yyyy"), "") 'IIf(Len(tmpPVDate), Format(tmpPVDate, "dd/mm/yyyy"), "")
                
                Record.SubItems(46) = IIf(Len(rst2!PVCLNChequeNo), Trim(rst2!PVCLNChequeNo), "") 'IIf(Len(tmpChqNo), tmpChqNo, "")
                
                Record.SubItems(47) = IIf(Len(rst2!PVCLNChequeDate), Format(rst2!PVCLNChequeDate, "dd/mm/yyyy"), "") 'IIf(Len(tmpChqDate), Format(tmpChqDate, "dd/mm/yyyy"), "")
                
                'If rst2!coverid = "00" Then
                Record.SubItems(48) = IIf(Len(rst2!OMNIRefNo), rst2!OMNIRefNo, "")
                'Else
                '    Record.SubItems(47) = IIf(Len(rst2!MBMAXANumber), rst2!MBMAXANumber, "")
                'End If
                
                Record.SubItems(49) = IIf(Len(rst2!MedInjType1), rst2!MedInjType1, "")
                Record.SubItems(50) = IIf(Len(rst2!MedInjType2), rst2!MedInjType2, "")
                Record.SubItems(51) = IIf(Len(rst2!MedInjType3), rst2!MedInjType3, "")
                Record.SubItems(52) = IIf(Len(rst2!MedInjType4), rst2!MedInjType4, "")
                Record.SubItems(53) = IIf(Len(rst2!MedInjType5), rst2!MedInjType5, "")
                Record.SubItems(54) = IIf(Len(rst2!MedInjType6), rst2!MedInjType6, "")
                Record.SubItems(55) = IIf(Len(rst2!MedInjType7), rst2!MedInjType7, "")
                
                Record.SubItems(56) = IIf(Len(rst2!MedInjDays1), rst2!MedInjDays1, "")
                Record.SubItems(57) = IIf(Len(rst2!MedInjDays2), rst2!MedInjDays2, "")
                Record.SubItems(58) = IIf(Len(rst2!MedInjDays3), rst2!MedInjDays3, "")
                Record.SubItems(59) = IIf(Len(rst2!MedInjDays4), rst2!MedInjDays4, "")
                Record.SubItems(60) = IIf(Len(rst2!MedInjDays5), rst2!MedInjDays5, "")
                
                Record.SubItems(61) = IIf(Len(rst2!MedInjQty1), rst2!MedInjQty1, "")
                Record.SubItems(62) = IIf(Len(rst2!MedInjQty2), rst2!MedInjQty2, "")
                Record.SubItems(62) = IIf(Len(rst2!MedInjQty3), rst2!MedInjQty3, "")
                Record.SubItems(64) = IIf(Len(rst2!MedInjQty4), rst2!MedInjQty4, "")
                Record.SubItems(65) = IIf(Len(rst2!MedInjQty5), rst2!MedInjQty5, "")
                tmpGPFee = IIf(IsNumeric(rst2!GPFee), rst2!GPFee, 0)
                tmpSPFee = IIf(IsNumeric(rst2!SPFee), rst2!SPFee, 0)
                tmpGPLTMed = IIf(IsNumeric(rst2!GPLTMed), rst2!GPLTMed, 0)
                tmpSPLTMed = IIf(IsNumeric(rst2!SPLTMed), rst2!SPLTMed, 0)
                tmpGPMed = IIf(IsNumeric(rst2!GPMed), rst2!GPMed, 0)
                tmpSPMed = IIf(IsNumeric(rst2!SPMed), rst2!SPMed, 0)
                tmpGPInj = IIf(IsNumeric(rst2!GPInj), rst2!GPInj, 0)
                tmpSPInj = IIf(IsNumeric(rst2!SPInj), rst2!SPInj, 0)
                
                Record.SubItems(66) = CDbl(tmpGPFee) + CDbl(tmpSPFee)
                Record.SubItems(67) = CDbl(tmpGPLTMed) + CDbl(tmpSPLTMed) + CDbl(tmpGPMed) + CDbl(tmpSPMed)
                Record.SubItems(68) = CDbl(tmpGPInj) + CDbl(tmpSPInj)
                Record.SubItems(69) = IIf(Len(rst2!CLMOMNISent), rst2!CLMOMNISent, "")
                Record.SubItems(70) = IIf(Len(rst2!CASMgmtNotes), rst2!CASMgmtNotes, "")
                Record.SubItems(71) = IIf(Len(rst2!CasRemark), rst2!CasRemark, "")
                Record.SubItems(72) = IIf(Len(rst2!DoctorName), rst2!DoctorName, "")
                Record.SubItems(73) = IIf(Len(rst2!IcNo), rst2!IcNo, "")
                Record.SubItems(74) = IIf(Len(rst2!coverid), rst2!coverid, "")
                Record.SubItems(75) = IIf(Len(rst2!OPBatchDate), Format(rst2!OPBatchDate, "dd/mm/yyyy"), "")
                Record.SubItems(76) = IIf(Len(rst2!recordingpic), Trim(rst2!recordingpic), "")
                Record.SubItems(77) = IIf(Len(rst2!PreliminaryDiag), Trim(rst2!PreliminaryDiag), "")
                
                Record.SubItems(78) = IIf(Len(rst2!clmsubsidiary), Trim(rst2!clmsubsidiary), "")
                Record.SubItems(79) = IIf(Len(rst2!clmcostcenter), Trim(rst2!clmcostcenter), "")
                Record.SubItems(80) = IIf(Len(rst2!CLMDeptName), Trim(rst2!CLMDeptName), "")
                Record.SubItems(81) = IIf(Len(rst2!InternalRemarks), Trim(rst2!InternalRemarks), "")
                
                If Current = 249 Then
                Dim str As String
                str = 250
                End If
                
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
                
                If Len(rst2!CLMBTBGAmt) Then
                    TotBTBGAmt = TotBTBGAmt + rst2!CLMBTBGAmt
                End If
                txtTotalBTBGAmt.Text = Format(TotBTBGAmt, "#,##0.00")
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
    
    'strsql = "SELECT PVCLNNo, PVCLNDate,  PVCLNChequeNo, PVCLNChequeDate " & _
    '         " From Payment2CLN WHERE CASNumber='" & Trim(tmpCASNumber) & "'"
    'PayRec.Open strsql, MediConnCISPayment, adOpenForwardOnly, adLockOptimistic
    'Do While Not PayRec.EOF
        
    '    tmpPVNo = IIf(Len(PayRec!PVCLNNo), PayRec!PVCLNNo, "")
    ''    tmpPVDate = IIf(Len(PayRec!PVCLNDate), PayRec!PVCLNDate, "")
    '    tmpChqNo = IIf(Len(PayRec!PVCLNChequeNo), PayRec!PVCLNChequeNo, "")
    '    tmpChqDate = IIf(Len(PayRec!PVCLNChequeDate), PayRec!PVCLNChequeDate, "")
    '
    '    PayRec.MoveNext
    'Loop
    'PayRec.Close
    'Set PayRec = Nothing
    
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
        frmInvoiceAdj.txtCaseNo = Trim(lstInvoiceTrackList.SelectedItem.Text)
        Call frmInvoiceAdj.CmdGo_Click
        Load frmInvoiceAdj
        frmInvoiceAdj.Show
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
        
       ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
       ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        'CLN.Open "Select CLNCode from CLN", medixmasterconn, adOpenKeyset, adLockOptimistic
        StrSql = "EXEC SearchCLN"
        CLN.Open StrSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic
    
        Do Until CLN.EOF
            CboClinic.AddItem CLN!CLNCode & " - " & CLN!CLNName
            CLN.MoveNext
        Loop
        CLN.Close
        Set CLN = Nothing
    
     '   MediConnCISDB.Close
      '  Set MediConnCISDB = Nothing
       ' medixmasterconn.Close
      '  Set medixmasterconn = Nothing
       ' MediConnCISWS.Close
      '  Set MediConnCISWS = Nothing
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
       ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
       ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        
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
        
      '  MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
      '  medixmasterconn.Close
       ' Set medixmasterconn = Nothing
      '  MediConnCISWS.Close
      '  Set MediConnCISWS = Nothing
        
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

'# Added by Ravi
Private Sub CboGrpCom_Click()
    If (CboGrpCom.Text = "UNITED OVERSEAS BANK (MALAYSIA) BHD") Or (CboPayor.Text = "UOB PROPERTIES BHD") Or (CboPayor.Text = "UNITED OVERSEAS BANK LIMITED") Then
       CboFlexOP.Enabled = True
       LblFlexOP.Enabled = True
    Else
       CboFlexOP.Enabled = False
       LblFlexOP.Enabled = False
    End If
End Sub
'# Added by Ravi


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


