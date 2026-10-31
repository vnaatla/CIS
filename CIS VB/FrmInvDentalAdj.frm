VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmInvDentalAdj 
   Caption         =   "Invoice Adjustment - Dental"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameCAS 
      Height          =   615
      Left            =   5760
      TabIndex        =   139
      Top             =   225
      Width           =   6080
      Begin VB.ComboBox cboStatus 
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
         Left            =   600
         TabIndex        =   4
         Text            =   "O - Open"
         Top             =   200
         Width           =   1095
      End
      Begin VB.ComboBox cboClaimability 
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
         Left            =   2760
         TabIndex        =   5
         Text            =   "A - Accepted"
         Top             =   200
         Width           =   1455
      End
      Begin VB.ComboBox CboPending 
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
         Left            =   5040
         TabIndex        =   6
         Text            =   "N - No"
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label107 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   142
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label25 
         Caption         =   "Claimability"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   141
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label24 
         Caption         =   "Pending"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   140
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame FrameDis 
      Enabled         =   0   'False
      Height          =   1550
      Left            =   8800
      TabIndex        =   72
      Top             =   765
      Width           =   3015
      Begin VB.TextBox txtDentStatus 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   18
         Top             =   1200
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox txtPolStatus 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   920
         TabIndex        =   14
         Top             =   220
         Width           =   615
      End
      Begin VB.TextBox TxtClnContract 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   920
         TabIndex        =   15
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox TxtBalLimit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2320
         TabIndex        =   19
         Top             =   180
         Width           =   615
      End
      Begin VB.TextBox txtLTM 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   920
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TxtCoPay 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   920
         TabIndex        =   17
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox TxtVisitLimit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2320
         TabIndex        =   20
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox txtDentLmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2320
         TabIndex        =   21
         Top             =   680
         Width           =   615
      End
      Begin VB.TextBox txtSpecLmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2320
         TabIndex        =   22
         Top             =   920
         Width           =   615
      End
      Begin VB.TextBox txtMepLmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2320
         TabIndex        =   23
         Top             =   1160
         Width           =   615
      End
      Begin VB.Label Label43 
         Caption         =   "MEP Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1620
         TabIndex        =   81
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label39 
         Caption         =   "Spec Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1620
         TabIndex        =   80
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label42 
         Caption         =   "Dent Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1620
         TabIndex        =   79
         Top             =   740
         Width           =   735
      End
      Begin VB.Label Label69 
         BackColor       =   &H80000013&
         Caption         =   "LTM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   135
         Left            =   120
         TabIndex        =   78
         Top             =   820
         Width           =   375
      End
      Begin VB.Label Label68 
         Caption         =   "Co-Pay"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   77
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label34 
         Caption         =   "Vst Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1620
         TabIndex        =   76
         Top             =   490
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000013&
         Caption         =   "Bal. Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1620
         TabIndex        =   75
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label40 
         Caption         =   "Card Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   560
         Width           =   855
      End
      Begin VB.Label Label38 
         Caption         =   "Pol. Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   260
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   0
      TabIndex        =   137
      Top             =   225
      Width           =   5745
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4980
         TabIndex        =   71
         Top             =   200
         Width           =   660
      End
      Begin VB.TextBox txtCaseNo 
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
         Left            =   1080
         TabIndex        =   0
         Top             =   200
         Width           =   2175
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   1
         Top             =   200
         Width           =   540
      End
      Begin VB.CommandButton CmdPrev 
         Caption         =   "Pre&v"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3900
         TabIndex        =   2
         Top             =   200
         Width           =   540
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   3
         Top             =   200
         Width           =   540
      End
      Begin VB.Label Label23 
         Caption         =   "Case No."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   138
         Top             =   255
         Width           =   735
      End
   End
   Begin VB.Frame FrameMBMCLN 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   0
      TabIndex        =   83
      Top             =   735
      Width           =   5040
      Begin VB.TextBox txtMBMNo 
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
         Left            =   1080
         TabIndex        =   7
         Top             =   195
         Width           =   3855
      End
      Begin VB.TextBox txtMBMName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   450
         Width           =   3855
      End
      Begin VB.ComboBox CboClinic 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         TabIndex        =   10
         Top             =   700
         Width           =   2895
      End
      Begin VB.TextBox TxtClinicCode 
         Height          =   315
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   9
         Top             =   740
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "MBM Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Clinic"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   740
         Width           =   495
      End
      Begin VB.Label Label54 
         Caption         =   "MBM No."
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         Top             =   210
         Width           =   1095
      End
   End
   Begin VB.Frame FrameInvoice 
      Height          =   1100
      Left            =   5040
      TabIndex        =   87
      Top             =   765
      Width           =   3735
      Begin VB.TextBox txtInvNo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         MaxLength       =   20
         TabIndex        =   11
         Top             =   195
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtInvDate 
         Height          =   300
         Left            =   2640
         TabIndex        =   12
         Top             =   200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
      Begin VB.TextBox TxtTotInvoiceAmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   13
         Top             =   450
         Width           =   975
      End
      Begin VB.ComboBox CboPayto 
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
         Left            =   840
         TabIndex        =   143
         Text            =   "D - Dental"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtNPClinic 
         Height          =   285
         Left            =   840
         TabIndex        =   144
         Top             =   735
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label LblNP 
         Caption         =   "NP Clinic"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   146
         Top             =   765
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label109 
         Caption         =   "Pay To"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   145
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label67 
         Caption         =   "Inv. Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   90
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label35 
         Caption         =   "Inv. No"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Inv. Amt"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   88
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame FrameInv 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   0
      TabIndex        =   91
      Top             =   1785
      Width           =   8780
      Begin VB.TextBox txtBillNo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   24
         Top             =   150
         Width           =   1935
      End
      Begin VB.TextBox txtMC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7800
         MaxLength       =   3
         TabIndex        =   27
         Top             =   150
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtVisitDate 
         Height          =   315
         Left            =   3960
         TabIndex        =   25
         Top             =   150
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
      Begin MSMask.MaskEdBox txtVisitTime 
         Height          =   315
         Left            =   5880
         TabIndex        =   26
         Top             =   150
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         Caption         =   "Bill No"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "MC Days"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   95
         Top             =   200
         Width           =   735
      End
      Begin VB.Label Label66 
         Caption         =   "Visit Time"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   94
         Top             =   195
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Visit Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   93
         Top             =   195
         Width           =   735
      End
   End
   Begin VB.Frame FrameDent 
      Height          =   4650
      Left            =   6480
      TabIndex        =   82
      Top             =   2235
      Width           =   5415
      Begin MSFlexGridLib.MSFlexGrid FgridDent 
         CausesValidation=   0   'False
         Height          =   3900
         Left            =   60
         TabIndex        =   63
         Top             =   120
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   6879
         _Version        =   393216
         Rows            =   14
         Cols            =   4
         FixedCols       =   2
         BackColorFixed  =   -2147483638
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame FrameDet 
      Height          =   2205
      Left            =   0
      TabIndex        =   107
      Top             =   2235
      Width           =   6450
      Begin VB.CheckBox ChkToothRef1 
         Height          =   255
         Index           =   6
         Left            =   6000
         TabIndex        =   123
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox ChkToothRef1 
         Height          =   255
         Index           =   5
         Left            =   6000
         TabIndex        =   122
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox ChkToothRef1 
         Height          =   255
         Index           =   4
         Left            =   6000
         TabIndex        =   121
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox ChkToothRef1 
         Height          =   255
         Index           =   3
         Left            =   6000
         TabIndex        =   120
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox ChkToothRef1 
         Height          =   255
         Index           =   2
         Left            =   6000
         TabIndex        =   119
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox ChkToothRef1 
         Height          =   255
         Index           =   1
         Left            =   6000
         TabIndex        =   118
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox ChkToothRef1 
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   117
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox ChkProcCat1 
         Height          =   255
         Index           =   0
         Left            =   2835
         TabIndex        =   116
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox ChkProcCat1 
         Height          =   255
         Index           =   1
         Left            =   2835
         TabIndex        =   115
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox ChkProcCat1 
         Height          =   255
         Index           =   2
         Left            =   2835
         TabIndex        =   114
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox ChkProcCat1 
         Height          =   255
         Index           =   3
         Left            =   2835
         TabIndex        =   113
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox ChkProcCat1 
         Height          =   255
         Index           =   4
         Left            =   2835
         TabIndex        =   112
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox ChkProcCat1 
         Height          =   255
         Index           =   5
         Left            =   2835
         TabIndex        =   111
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox ChkProcCat1 
         Height          =   255
         Index           =   6
         Left            =   2835
         TabIndex        =   110
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox txtTabMBMNo 
         Height          =   285
         Left            =   3120
         TabIndex        =   109
         Text            =   "Text1"
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtTabCoverID 
         Height          =   285
         Left            =   3120
         TabIndex        =   108
         Text            =   "Text1"
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox cboProcCat1 
         Height          =   315
         Index           =   0
         Left            =   315
         TabIndex        =   28
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox cboProcCat1 
         Height          =   315
         Index           =   1
         Left            =   315
         TabIndex        =   30
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox cboProcCat1 
         Height          =   315
         Index           =   2
         Left            =   315
         TabIndex        =   32
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cboProcCat1 
         Height          =   315
         Index           =   3
         Left            =   315
         TabIndex        =   34
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ComboBox cboProcCat1 
         Height          =   315
         Index           =   4
         Left            =   315
         TabIndex        =   36
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox cboProcCat1 
         Height          =   315
         Index           =   5
         Left            =   315
         TabIndex        =   38
         Top             =   1560
         Width           =   2415
      End
      Begin VB.ComboBox cboProcCat1 
         Height          =   315
         Index           =   6
         Left            =   315
         TabIndex        =   40
         Top             =   1800
         Width           =   2415
      End
      Begin VB.ComboBox cboToothRef1 
         Height          =   315
         Index           =   0
         Left            =   3480
         TabIndex        =   29
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox cboToothRef1 
         Height          =   315
         Index           =   1
         Left            =   3480
         TabIndex        =   31
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox cboToothRef1 
         Height          =   315
         Index           =   2
         Left            =   3480
         TabIndex        =   33
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cboToothRef1 
         Height          =   315
         Index           =   3
         Left            =   3480
         TabIndex        =   35
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ComboBox cboToothRef1 
         Height          =   315
         Index           =   4
         Left            =   3480
         TabIndex        =   37
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox cboToothRef1 
         Height          =   315
         Index           =   5
         Left            =   3480
         TabIndex        =   39
         Top             =   1560
         Width           =   2415
      End
      Begin VB.ComboBox cboToothRef1 
         Height          =   315
         Index           =   6
         Left            =   3480
         TabIndex        =   41
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Procedures"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   134
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   133
         Top             =   1180
         Width           =   135
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000013&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   940
         Width           =   135
      End
      Begin VB.Label Label6 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   680
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   130
         Top             =   400
         Width           =   135
      End
      Begin VB.Label Label10 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label12 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   128
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label13 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   127
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label19 
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2940
         TabIndex        =   126
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label20 
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   6045
         TabIndex        =   125
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label21 
         Caption         =   "Tooth Ref."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   124
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame FrameMed 
      Caption         =   "Medication && Injection"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   0
      TabIndex        =   98
      Top             =   4440
      Width           =   6450
      Begin VB.TextBox txtMedType1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   315
         MaxLength       =   50
         TabIndex        =   42
         Top             =   405
         Width           =   4695
      End
      Begin VB.TextBox txtMedType1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   315
         TabIndex        =   45
         Top             =   690
         Width           =   4695
      End
      Begin VB.TextBox txtMedType1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   315
         TabIndex        =   48
         Top             =   945
         Width           =   4695
      End
      Begin VB.TextBox txtMedType1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   315
         TabIndex        =   51
         Top             =   1230
         Width           =   4695
      End
      Begin VB.TextBox txtMedType1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   315
         TabIndex        =   54
         Top             =   1485
         Width           =   4695
      End
      Begin VB.TextBox txtMedDay1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4980
         MaxLength       =   3
         TabIndex        =   43
         Top             =   405
         Width           =   615
      End
      Begin VB.TextBox txtMedDay1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4980
         MaxLength       =   3
         TabIndex        =   46
         Top             =   690
         Width           =   615
      End
      Begin VB.TextBox txtMedDay1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4980
         MaxLength       =   3
         TabIndex        =   49
         Top             =   945
         Width           =   615
      End
      Begin VB.TextBox txtMedDay1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   4980
         MaxLength       =   3
         TabIndex        =   52
         Top             =   1230
         Width           =   615
      End
      Begin VB.TextBox txtMedDay1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4980
         MaxLength       =   3
         TabIndex        =   55
         Top             =   1485
         Width           =   615
      End
      Begin VB.TextBox txtMedQty1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   5560
         MaxLength       =   3
         TabIndex        =   44
         Top             =   405
         Width           =   615
      End
      Begin VB.TextBox txtMedQty1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   5560
         MaxLength       =   3
         TabIndex        =   47
         Top             =   690
         Width           =   615
      End
      Begin VB.TextBox txtMedQty1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   5560
         MaxLength       =   3
         TabIndex        =   50
         Top             =   945
         Width           =   615
      End
      Begin VB.TextBox txtMedQty1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   5560
         MaxLength       =   3
         TabIndex        =   53
         Top             =   1230
         Width           =   615
      End
      Begin VB.TextBox txtMedQty1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   5560
         MaxLength       =   3
         TabIndex        =   56
         Top             =   1485
         Width           =   615
      End
      Begin VB.TextBox txtMedType1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   315
         TabIndex        =   57
         Top             =   1750
         Width           =   4695
      End
      Begin VB.TextBox txtMedDay1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4980
         MaxLength       =   3
         TabIndex        =   58
         Top             =   1750
         Width           =   615
      End
      Begin VB.TextBox txtMedQty1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   5560
         MaxLength       =   3
         TabIndex        =   59
         Top             =   1750
         Width           =   615
      End
      Begin VB.TextBox txtMedType1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   315
         TabIndex        =   60
         Top             =   2040
         Width           =   4695
      End
      Begin VB.TextBox txtMedDay1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   4980
         MaxLength       =   3
         TabIndex        =   61
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtMedQty1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   5560
         MaxLength       =   3
         TabIndex        =   62
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label27 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   152
         Top             =   2080
         Width           =   135
      End
      Begin VB.Label Label26 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   151
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label50 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   1515
         Width           =   135
      End
      Begin VB.Label Label41 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   405
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   705
         Width           =   375
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000013&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   101
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label Label72 
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5700
         TabIndex        =   100
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label71 
         Caption         =   "Days"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5080
         TabIndex        =   99
         Top             =   135
         Width           =   375
      End
   End
   Begin VB.Frame Frame7 
      Height          =   555
      Left            =   0
      TabIndex        =   96
      Top             =   6840
      Width           =   11895
      Begin VB.CommandButton cmdCompute 
         Caption         =   "&Compute"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7200
         TabIndex        =   64
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11040
         TabIndex        =   69
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton CmdClearAll 
         Caption         =   "&Clear All"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10200
         TabIndex        =   68
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdMBMInfo 
         Caption         =   "&MBM Info"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8640
         TabIndex        =   66
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8040
         TabIndex        =   65
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9600
         TabIndex        =   67
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9600
         TabIndex        =   97
         Top             =   160
         Width           =   615
      End
      Begin VB.Label LblBordx 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   735
         TabIndex        =   150
         Top             =   195
         Width           =   720
      End
      Begin VB.Label Label108 
         Caption         =   "Bordx ?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   149
         Top             =   195
         Width           =   615
      End
      Begin VB.Label Label33 
         Caption         =   "Batch No."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   148
         Top             =   195
         Width           =   735
      End
      Begin VB.Label LblBatchNo 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   147
         Top             =   195
         Width           =   840
      End
   End
   Begin MSComctlLib.ListView lstInvoice 
      Height          =   225
      Left            =   0
      TabIndex        =   70
      Top             =   7560
      Visible         =   0   'False
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   397
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   518
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Batch No"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Case No."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Member No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cov ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Patient Name"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Payor"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ClnPln"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Cln Ins"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "HltCode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "BalLmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Vst Lmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ClnContract"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "VstInd"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "AnnLmtInd"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "CoPay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "PolNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "ICNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Effdate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "ExpDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Payto"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "payee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "SuppLmtType"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "DentStatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "SpecStatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "MEPStatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "DentLmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "SpecLmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "MepLmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "DentLmtStatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "SpecLmtStatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "DentExtLmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "1stSpecLmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "PolStatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   40
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Text            =   "VisitDate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Text            =   "Visit Time"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   52
         Text            =   "Dental"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   53
         Text            =   "MC"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   54
         Text            =   "MedInjType1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   55
         Text            =   "MedInjType2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   56
         Text            =   "MedInjType3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   57
         Text            =   "MedInjType4"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   58
         Text            =   "MedInjType5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   59
         Text            =   "MedInjDay1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   60
         Text            =   "MedInjDay2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   61
         Text            =   "MedInjDay3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   62
         Text            =   "MedInjDay4"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(64) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   63
         Text            =   "MedInjDay5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(65) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   64
         Text            =   "MedInjQty1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(66) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   65
         Text            =   "MedInjQty2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(67) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   66
         Text            =   "MedInjQty3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(68) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   67
         Text            =   "MedInjQty4"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(69) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   68
         Text            =   "MedInjQty5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(70) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   69
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(71) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   70
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(72) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   71
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(73) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   72
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(74) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   73
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(75) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   74
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(76) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   75
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(77) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   76
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(78) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   77
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(79) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   78
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(80) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   79
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(81) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   80
         Text            =   "BillNo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(82) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   81
         Text            =   "Bill Amt"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(83) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   82
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(84) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   83
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(85) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   84
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(86) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   85
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(87) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   86
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(88) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   87
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(89) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   88
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(90) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   89
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(91) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   90
         Text            =   "Procedures1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(92) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   91
         Text            =   "Procedures2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(93) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   92
         Text            =   "Procedures3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(94) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   93
         Text            =   "Procedures4"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(95) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   94
         Text            =   "Procedures5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(96) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   95
         Text            =   "Procedures6"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(97) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   96
         Text            =   "Procedures7"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(98) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   97
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(99) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   98
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(100) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   99
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(101) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   100
         Text            =   "ToothRef1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(102) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   101
         Text            =   "ToothRef2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(103) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   102
         Text            =   "ToothRef3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(104) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   103
         Text            =   "ToothRef4"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(105) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   104
         Text            =   "ToothRef5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(106) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   105
         Text            =   "ToothRef6"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(107) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   106
         Text            =   "ToothRef7"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(108) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   107
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(109) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   108
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(110) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   109
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(111) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   110
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(112) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   111
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(113) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   112
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(114) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   113
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(115) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   114
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(116) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   115
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(117) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   116
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(118) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   117
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(119) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   118
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(120) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   119
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(121) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   120
         Text            =   "NoOfTeeth1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(122) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   121
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(123) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   122
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(124) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   123
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(125) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   124
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(126) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   125
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(127) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   126
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(128) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   127
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(129) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   128
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(130) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   129
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(131) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   130
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(132) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   131
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(133) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   132
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(134) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   133
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(135) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   134
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(136) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   135
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(137) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   136
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(138) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   137
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(139) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   138
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(140) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   139
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(141) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   140
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(142) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   141
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(143) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   142
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(144) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   143
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(145) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   144
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(146) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   145
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(147) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   146
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(148) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   147
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(149) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   148
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(150) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   149
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(151) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   150
         Text            =   "Amt(teeth)1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(152) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   151
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(153) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   152
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(154) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   153
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(155) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   154
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(156) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   155
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(157) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   156
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(158) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   157
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(159) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   158
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(160) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   159
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(161) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   160
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(162) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   161
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(163) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   162
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(164) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   163
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(165) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   164
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(166) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   165
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(167) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   166
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(168) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   167
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(169) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   168
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(170) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   169
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(171) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   170
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(172) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   171
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(173) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   172
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(174) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   173
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(175) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   174
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(176) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   175
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(177) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   176
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(178) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   177
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(179) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   178
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(180) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   179
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(181) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   180
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(182) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   181
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(183) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   182
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(184) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   183
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(185) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   184
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(186) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   185
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(187) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   186
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(188) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   187
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(189) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   188
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(190) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   189
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(191) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   190
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(192) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   191
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(193) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   192
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(194) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   193
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(195) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   194
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(196) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   195
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(197) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   196
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(198) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   197
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(199) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   198
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(200) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   199
         Text            =   "Rsv for inv 1~200"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(201) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   200
         Text            =   "CLMOpenDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(202) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   201
         Text            =   "CLMOpenBy"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(203) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   202
         Text            =   "CLMUpdateBy"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(204) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   203
         Text            =   "CLMUpdateDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(205) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   204
         Text            =   "CLMBordxStatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(206) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   205
         Text            =   "CLMStatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(207) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   206
         Text            =   "CLMCloseBy"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(208) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   207
         Text            =   "CLMCloseDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(209) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   208
         Text            =   "TotFCLN"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(210) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   209
         Text            =   "TotCLNExc"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(211) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   210
         Text            =   "TotFinalAct"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(212) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   211
         Text            =   "TotFinalAPP"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(213) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   212
         Text            =   "TotFinalMem"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(214) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   213
         Text            =   "CLMBalLmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(215) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   214
         Text            =   "CLMPrevBal"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(216) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   215
         Text            =   "CLMPaidByM2M"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(217) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   216
         Text            =   "CLMPaidByM2C"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(218) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   217
         Text            =   "CLMActGPBMed"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(219) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   218
         Text            =   "CLMActGPMedInj"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(220) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   219
         Text            =   "CLMActGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(221) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   220
         Text            =   "CLMActSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(222) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   221
         Text            =   "CLMActSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(223) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   222
         Text            =   "CLMActSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(224) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   223
         Text            =   "CLMActXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(225) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   224
         Text            =   "CLMActUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(226) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   225
         Text            =   "CLMActECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(227) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   226
         Text            =   "CLMActDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(228) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   227
         Text            =   "CLMActNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(229) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   228
         Text            =   "CLMActID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(230) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   229
         Text            =   "CLMActTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(231) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   230
         Text            =   "CLMActMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(232) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   231
         Text            =   "CLMActPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(233) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   232
         Text            =   "CLMActBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(234) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   233
         Text            =   "CLMActBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(235) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   234
         Text            =   "CLMActUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(236) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   235
         Text            =   "CLMActPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(237) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   236
         Text            =   "CLMActLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(238) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   237
         Text            =   "CLMActDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(239) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   238
         Text            =   "CLMActDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(240) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   239
         Text            =   "CLMActHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(241) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   240
         Text            =   "CLMActMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(242) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   241
         Text            =   "CLMActIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(243) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   242
         Text            =   "CLMActDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(244) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   243
         Text            =   "CLMActMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(245) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   244
         Text            =   "CLMActPRVCheckup"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(246) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   245
         Text            =   "CLMActVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(247) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   246
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(248) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   247
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(249) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   248
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(250) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   249
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(251) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   250
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(252) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   251
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(253) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   252
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(254) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   253
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(255) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   254
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(256) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   255
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(257) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   256
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(258) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   257
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(259) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   258
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(260) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   259
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(261) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   260
         Text            =   "CLMContGPBMed"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(262) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   261
         Text            =   "CLMContGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(263) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   262
         Text            =   "CLMContGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(264) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   263
         Text            =   "CLMContSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(265) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   264
         Text            =   "CLMContSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(266) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   265
         Text            =   "CLMContSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(267) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   266
         Text            =   "CLMContXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(268) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   267
         Text            =   "CLMContUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(269) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   268
         Text            =   "CLMContECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(270) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   269
         Text            =   "CLMContDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(271) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   270
         Text            =   "CLMContNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(272) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   271
         Text            =   "CLMContID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(273) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   272
         Text            =   "CLMContTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(274) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   273
         Text            =   "CLMContMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(275) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   274
         Text            =   "CLMContPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(276) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   275
         Text            =   "CLMContBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(277) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   276
         Text            =   "CLMContBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(278) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   277
         Text            =   "CLMContUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(279) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   278
         Text            =   "CLMContPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(280) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   279
         Text            =   "CLMContLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(281) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   280
         Text            =   "CLMContDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(282) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   281
         Text            =   "CLMContDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(283) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   282
         Text            =   "CLMContHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(284) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   283
         Text            =   "CLMContMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(285) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   284
         Text            =   "CLMContIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(286) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   285
         Text            =   "CLMContDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(287) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   286
         Text            =   "CLMContMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(288) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   287
         Text            =   "CLMContPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(289) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   288
         Text            =   "CLMContVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(290) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   289
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(291) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   290
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(292) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   291
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(293) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   292
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(294) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   293
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(295) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   294
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(296) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   295
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(297) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   296
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(298) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   297
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(299) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   298
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(300) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   299
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(301) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   300
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(302) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   301
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(303) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   302
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(304) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   303
         Text            =   "CLMAppGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(305) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   304
         Text            =   "CLMAppGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(306) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   305
         Text            =   "CLMAppGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(307) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   306
         Text            =   "CLMAppSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(308) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   307
         Text            =   "CLMAppSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(309) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   308
         Text            =   "CLMAppSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(310) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   309
         Text            =   "CLMAppXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(311) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   310
         Text            =   "CLMAppUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(312) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   311
         Text            =   "CLMAppECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(313) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   312
         Text            =   "CLMAppDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(314) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   313
         Text            =   "CLMAppNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(315) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   314
         Text            =   "CLMAppID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(316) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   315
         Text            =   "CLMAppTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(317) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   316
         Text            =   "CLMAppMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(318) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   317
         Text            =   "CLMAppPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(319) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   318
         Text            =   "CLMAppBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(320) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   319
         Text            =   "CLMAppBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(321) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   320
         Text            =   "CLMAppUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(322) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   321
         Text            =   "CLMAppPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(323) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   322
         Text            =   "CLMAppLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(324) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   323
         Text            =   "CLMAppDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(325) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   324
         Text            =   "CLMAppDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(326) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   325
         Text            =   "CLMAppHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(327) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   326
         Text            =   "CLMAppMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(328) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   327
         Text            =   "CLMAppIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(329) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   328
         Text            =   "CLMAppDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(330) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   329
         Text            =   "CLMAppMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(331) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   330
         Text            =   "CLMAppPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(332) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   331
         Text            =   "CLMAppVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(333) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   332
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(334) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   333
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(335) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   334
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(336) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   335
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(337) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   336
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(338) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   337
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(339) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   338
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(340) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   339
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(341) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   340
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(342) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   341
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(343) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   342
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(344) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   343
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(345) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   344
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(346) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   345
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(347) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   346
         Text            =   "CLMBnfGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(348) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   347
         Text            =   "CLMBnfGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(349) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   348
         Text            =   "CLMBnfGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(350) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   349
         Text            =   "CLMBnfSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(351) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   350
         Text            =   "CLMBnfSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(352) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   351
         Text            =   "CLMBnfSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(353) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   352
         Text            =   "CLMBnfXray"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(354) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   353
         Text            =   "CLMBnfUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(355) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   354
         Text            =   "CLMBnfECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(356) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   355
         Text            =   "CLMBnfDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(357) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   356
         Text            =   "CLMBnfNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(358) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   357
         Text            =   "CLMBnfID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(359) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   358
         Text            =   "CLMBnfTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(360) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   359
         Text            =   "CLMBnfMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(361) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   360
         Text            =   "CLMBnfPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(362) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   361
         Text            =   "CLMBnfBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(363) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   362
         Text            =   "CLMBnfBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(364) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   363
         Text            =   "CLMBnfUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(365) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   364
         Text            =   "CLMBnfPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(366) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   365
         Text            =   "CLMBnfLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(367) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   366
         Text            =   "CLMBnfDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(368) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   367
         Text            =   "CLMBnfDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(369) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   368
         Text            =   "CLMBnfHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(370) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   369
         Text            =   "CLMBnfMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(371) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   370
         Text            =   "CLMBnfIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(372) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   371
         Text            =   "CLMBnfDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(373) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   372
         Text            =   "CLMBnfMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(374) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   373
         Text            =   "CLMBnfPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(375) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   374
         Text            =   "CLMBnfVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(376) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   375
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(377) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   376
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(378) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   377
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(379) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   378
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(380) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   379
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(381) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   380
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(382) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   381
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(383) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   382
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(384) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   383
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(385) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   384
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(386) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   385
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(387) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   386
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(388) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   387
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(389) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   388
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(390) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   389
         Text            =   "CLMMBMExcGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(391) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   390
         Text            =   "CLMMBMExcGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(392) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   391
         Text            =   "CLMMBMExcGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(393) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   392
         Text            =   "CLMMBMExcSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(394) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   393
         Text            =   "CLMMBMExcSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(395) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   394
         Text            =   "CLMMBMExcSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(396) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   395
         Text            =   "CLMMBMExcXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(397) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   396
         Text            =   "CLMMBMExcUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(398) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   397
         Text            =   "CLMMBMExcECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(399) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   398
         Text            =   "CLMMBMExcDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(400) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   399
         Text            =   "CLMMBMExcNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(401) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   400
         Text            =   "CLMMBMExcID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(402) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   401
         Text            =   "CLMMBMExcTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(403) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   402
         Text            =   "CLMMBMExcMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(404) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   403
         Text            =   "CLMMBMExcPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(405) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   404
         Text            =   "CLMMBMExcBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(406) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   405
         Text            =   "CLMMBMExcBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(407) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   406
         Text            =   "CLMMBMExcUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(408) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   407
         Text            =   "CLMMBMExcPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(409) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   408
         Text            =   "CLMMBMExcLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(410) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   409
         Text            =   "CLMMBMExcDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(411) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   410
         Text            =   "CLMMBMExcDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(412) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   411
         Text            =   "CLMMBMExcHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(413) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   412
         Text            =   "CLMMBMExcMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(414) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   413
         Text            =   "CLMMBMExcIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(415) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   414
         Text            =   "CLMMBMExcDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(416) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   415
         Text            =   "CLMMBMExcMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(417) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   416
         Text            =   "CLMMBMExcPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(418) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   417
         Text            =   "CLMMBMExcVision"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(419) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   418
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(420) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   419
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(421) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   420
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(422) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   421
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(423) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   422
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(424) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   423
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(425) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   424
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(426) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   425
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(427) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   426
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(428) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   427
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(429) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   428
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(430) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   429
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(431) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   430
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(432) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   431
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(433) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   432
         Text            =   "CLMNPayCLNGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(434) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   433
         Text            =   "CLMNPayCLNGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(435) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   434
         Text            =   "CLMNPayCLNGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(436) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   435
         Text            =   "CLMNPayCLNSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(437) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   436
         Text            =   "CLMNPayCLNSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(438) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   437
         Text            =   "CLMNPayCLNSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(439) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   438
         Text            =   "CLMNPayCLNXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(440) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   439
         Text            =   "CLMNPayCLNUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(441) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   440
         Text            =   "CLMNPayCLNECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(442) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   441
         Text            =   "CLMNPayCLNDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(443) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   442
         Text            =   "CLMNPayCLNNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(444) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   443
         Text            =   "CLMNPayCLNID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(445) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   444
         Text            =   "CLMNPayCLNTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(446) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   445
         Text            =   "CLMNPayCLNMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(447) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   446
         Text            =   "CLMNPayCLNPThers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(448) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   447
         Text            =   "CLMNPayCLNBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(449) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   448
         Text            =   "CLMNPayCLNBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(450) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   449
         Text            =   "CLMNPayCLNUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(451) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   450
         Text            =   "CLMNPayCLNPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(452) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   451
         Text            =   "CLMNPayCLNLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(453) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   452
         Text            =   "CLMNPayCLNDTPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(454) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   453
         Text            =   "CLMNPayCLNDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(455) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   454
         Text            =   "CLMNPayCLNHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(456) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   455
         Text            =   "CLMNPayCLNMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(457) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   456
         Text            =   "CLMNPayCLNIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(458) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   457
         Text            =   "CLMNPayCLNDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(459) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   458
         Text            =   "CLMNPayCLNMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(460) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   459
         Text            =   "CLMNPayCLNPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(461) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   460
         Text            =   "CLMNPayCLNVision"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(462) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   461
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(463) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   462
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(464) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   463
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(465) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   464
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(466) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   465
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(467) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   466
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(468) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   467
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(469) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   468
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(470) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   469
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(471) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   470
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(472) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   471
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(473) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   472
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(474) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   473
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(475) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   474
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(476) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   475
         Text            =   "CLMPayCLNGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(477) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   476
         Text            =   "CLMPayCLNGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(478) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   477
         Text            =   "CLMPayCLNGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(479) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   478
         Text            =   "CLMPayCLNSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(480) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   479
         Text            =   "CLMPayCLNSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(481) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   480
         Text            =   "CLMPayCLNSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(482) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   481
         Text            =   "CLMPayCLNXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(483) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   482
         Text            =   "CLMPayCLNUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(484) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   483
         Text            =   "CLMPayCLNECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(485) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   484
         Text            =   "CLMPayCLNDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(486) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   485
         Text            =   "CLMPayCLNNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(487) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   486
         Text            =   "CLMPayCLNID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(488) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   487
         Text            =   "CLMPayCLNTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(489) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   488
         Text            =   "CLMPayCLNMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(490) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   489
         Text            =   "CLMPayCLNPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(491) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   490
         Text            =   "CLMPayCLNBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(492) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   491
         Text            =   "CLMPayCLNBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(493) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   492
         Text            =   "CLMPayCLNUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(494) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   493
         Text            =   "CLMPayCLNPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(495) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   494
         Text            =   "CLMPayCLNLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(496) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   495
         Text            =   "CLMPayCLNDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(497) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   496
         Text            =   "CLMPayCLNDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(498) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   497
         Text            =   "CLMPayCLNHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(499) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   498
         Text            =   "CLMPayCLNMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(500) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   499
         Text            =   "CLMPayCLNIOhters"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(501) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   500
         Text            =   "CLMPayCLNDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(502) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   501
         Text            =   "CLMPayCLNMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(503) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   502
         Text            =   "CLMPayCLNPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(504) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   503
         Text            =   "CLMPayCLNVision"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(505) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   504
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(506) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   505
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(507) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   506
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(508) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   507
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(509) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   508
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(510) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   509
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(511) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   510
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(512) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   511
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(513) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   512
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(514) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   513
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(515) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   514
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(516) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   515
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(517) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   516
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(518) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   517
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstErr 
      Height          =   255
      Left            =   0
      TabIndex        =   106
      Top             =   7560
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "MBM No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cov"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Policy No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Patient IC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Eff Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Exp. Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bal. Limit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Visit Lmt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Err Msg"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblInvoice 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Invoice Adjustment - Dental"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   0
      TabIndex        =   136
      Top             =   0
      Width           =   11895
   End
   Begin VB.Label Label22 
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
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   135
      Top             =   7800
      Width           =   1935
   End
End
Attribute VB_Name = "FrmInvDentalAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
