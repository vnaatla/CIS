VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMBMCLNEnq 
   Caption         =   "Membership Enquiry - Clinical"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame16 
      Height          =   540
      Left            =   100
      TabIndex        =   16
      Top             =   240
      Width           =   11730
      Begin VB.TextBox txtMBMNumber 
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
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   0
         Top             =   160
         Width           =   2055
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Go"
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
         Left            =   3720
         TabIndex        =   1
         Top             =   160
         Width           =   810
      End
      Begin VB.CommandButton cmdSearchMBM 
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
         Height          =   285
         Left            =   6150
         TabIndex        =   4
         Top             =   160
         Width           =   810
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
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
         Left            =   5340
         TabIndex        =   3
         Top             =   160
         Width           =   810
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Pre"
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
         Left            =   4530
         TabIndex        =   2
         Top             =   160
         Width           =   810
      End
      Begin VB.Label Label76 
         Caption         =   "Member Number"
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
         TabIndex        =   17
         Top             =   200
         Width           =   1275
      End
   End
   Begin VB.Frame FrameEndt 
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   11730
      Begin VB.TextBox txtMBMName 
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
         Height          =   330
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   5
         Top             =   165
         Width           =   4935
      End
      Begin VB.TextBox txtMBMPolNo 
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
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   6
         Top             =   465
         Width           =   2295
      End
      Begin VB.ComboBox cboStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmMBMCLNEnq.frx":0000
         Left            =   7440
         List            =   "FrmMBMCLNEnq.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   165
         Width           =   1320
      End
      Begin VB.TextBox txtMBMIcBcPP 
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
         Left            =   4440
         MaxLength       =   30
         TabIndex        =   7
         Top             =   460
         Width           =   1820
      End
      Begin VB.ComboBox cboDataStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmMBMCLNEnq.frx":0004
         Left            =   7440
         List            =   "FrmMBMCLNEnq.frx":0006
         TabIndex        =   127
         Text            =   "A - Active"
         Top             =   440
         Width           =   1320
      End
      Begin VB.ComboBox txtPAYCode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8160
         Style           =   1  'Simple Combo
         TabIndex        =   10
         Top             =   440
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.ComboBox cboExpStatus 
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
         ItemData        =   "FrmMBMCLNEnq.frx":0008
         Left            =   7440
         List            =   "FrmMBMCLNEnq.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   128
         Top             =   725
         Width           =   1320
      End
      Begin VB.Label lblLimitStatus 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8775
         TabIndex        =   133
         Top             =   165
         Width           =   2835
      End
      Begin VB.Label Label55 
         Caption         =   "Pol. Status"
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
         Left            =   6480
         TabIndex        =   131
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label89 
         Caption         =   "Exp. Status"
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
         Left            =   6480
         TabIndex        =   130
         Top             =   735
         Width           =   915
      End
      Begin VB.Label Label62 
         Caption         =   "Data Status"
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
         Left            =   6480
         TabIndex        =   129
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label64 
         Caption         =   "IC. Num"
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
         Left            =   3720
         TabIndex        =   13
         Top             =   520
         Width           =   615
      End
      Begin VB.Label Label73 
         Caption         =   "Policy Number"
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
         TabIndex        =   12
         Top             =   495
         Width           =   1140
      End
      Begin VB.Label Label80 
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
         TabIndex        =   11
         Top             =   210
         Width           =   1140
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5835
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10292
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Principal &Details"
      TabPicture(0)   =   "FrmMBMCLNEnq.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePrincDet"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Supplementary"
      TabPicture(1)   =   "FrmMBMCLNEnq.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstCovAdd"
      Tab(1).Control(1)=   "FrameSupp"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Ajustment History"
      TabPicture(2)   =   "FrmMBMCLNEnq.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstAdjustmentCis"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "MBM History"
      TabPicture(3)   =   "FrmMBMCLNEnq.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstMBMHis"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Case"
      TabPicture(4)   =   "FrmMBMCLNEnq.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lstCaseList"
      Tab(4).Control(1)=   "lstSpecialistCases"
      Tab(4).Control(2)=   "Label36(8)"
      Tab(4).Control(3)=   "Label36(5)"
      Tab(4).Control(4)=   "lblTotalSpeVisits"
      Tab(4).Control(5)=   "Label36(7)"
      Tab(4).Control(6)=   "Label36(6)"
      Tab(4).Control(7)=   "lblTotalSpecBordx"
      Tab(4).Control(8)=   "lblTotalSpeBill"
      Tab(4).Control(9)=   "Label36(4)"
      Tab(4).Control(10)=   "lblTotalVisit"
      Tab(4).Control(11)=   "Label36(3)"
      Tab(4).Control(12)=   "lblBillAmt"
      Tab(4).Control(13)=   "lblTotalBordx"
      Tab(4).Control(14)=   "lblTotalFFS"
      Tab(4).Control(15)=   "Label36(2)"
      Tab(4).Control(16)=   "Label36(1)"
      Tab(4).Control(17)=   "Label36(0)"
      Tab(4).ControlCount=   18
      TabCaption(5)   =   "Benefit Schedule"
      TabPicture(5)   =   "FrmMBMCLNEnq.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lstPolicyWording"
      Tab(5).ControlCount=   1
      Begin VB.ListBox lstPolicyWording 
         Height          =   5130
         ItemData        =   "FrmMBMCLNEnq.frx":00B4
         Left            =   -74880
         List            =   "FrmMBMCLNEnq.frx":00B6
         TabIndex        =   196
         ToolTipText     =   "right click to randomize contents"
         Top             =   480
         Width           =   4455
      End
      Begin VB.Frame FrameSupp 
         Height          =   3045
         Left            =   -74880
         TabIndex        =   70
         Top             =   360
         Width           =   11460
         Begin VB.TextBox txtCovIcBcPp 
            Enabled         =   0   'False
            Height          =   330
            Left            =   8400
            MaxLength       =   30
            TabIndex        =   103
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox txtCovName 
            Enabled         =   0   'False
            Height          =   330
            Left            =   3720
            MaxLength       =   60
            TabIndex        =   102
            Top             =   240
            Width           =   3495
         End
         Begin VB.Frame Frame3 
            Height          =   1335
            Left            =   120
            TabIndex        =   71
            Top             =   1560
            Width           =   11145
            Begin VB.TextBox txtCovID 
               Height          =   315
               Left            =   8700
               TabIndex        =   146
               Top             =   1500
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label lblSpecCovAnn 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6765
               TabIndex        =   180
               Top             =   180
               Width           =   1275
            End
            Begin VB.Label Label27 
               Caption         =   "Spec. Annual Limit"
               Height          =   255
               Index           =   1
               Left            =   5040
               TabIndex        =   179
               Top             =   195
               Width           =   1695
            End
            Begin VB.Label lblSpecCovBal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6765
               TabIndex        =   178
               Top             =   900
               Width           =   1275
            End
            Begin VB.Label Label13 
               Caption         =   "Spec. Balance Limit"
               Height          =   255
               Index           =   1
               Left            =   5040
               TabIndex        =   177
               Top             =   930
               Width           =   1695
            End
            Begin VB.Label Label37 
               Caption         =   "Spec. Accum Amt"
               Height          =   255
               Index           =   1
               Left            =   5040
               TabIndex        =   176
               Top             =   555
               Width           =   1575
            End
            Begin VB.Label lblSpecCovAcc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6765
               TabIndex        =   175
               Top             =   540
               Width           =   1275
            End
            Begin VB.Label lblCovTopUp 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9540
               TabIndex        =   145
               Top             =   540
               Width           =   1200
            End
            Begin VB.Label lblCovFFS 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9540
               TabIndex        =   144
               Top             =   180
               Width           =   1200
            End
            Begin VB.Label Label14 
               Caption         =   "Top Up"
               Height          =   255
               Index           =   4
               Left            =   8340
               TabIndex        =   143
               Top             =   540
               Width           =   1095
            End
            Begin VB.Label Label14 
               Caption         =   "FFS"
               Height          =   255
               Index           =   3
               Left            =   8340
               TabIndex        =   142
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "MEP Limit"
               Height          =   255
               Index           =   0
               Left            =   2880
               TabIndex        =   101
               Top             =   2610
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label6 
               Caption         =   "1st Spec. Lmt"
               Height          =   255
               Left            =   300
               TabIndex        =   100
               Top             =   2370
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.Label lblCov1stSpecLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1425
               TabIndex        =   99
               Top             =   2340
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label Label9 
               Caption         =   "Dental Ext Lmt"
               Height          =   255
               Left            =   300
               TabIndex        =   98
               Top             =   1620
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.Label lblCovDentExtLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1425
               TabIndex        =   97
               Top             =   1620
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Line Line4 
               X1              =   120
               X2              =   10920
               Y1              =   840
               Y2              =   840
            End
            Begin VB.Label lblCovMepLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3960
               TabIndex        =   96
               Top             =   2580
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblCovSpecLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1425
               TabIndex        =   95
               Top             =   1980
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblCovDentLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1305
               TabIndex        =   94
               Top             =   2595
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblCovMCO 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   93
               Top             =   180
               Width           =   1035
            End
            Begin VB.Label lblCovPrem 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   92
               Top             =   1500
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Label lblCovPremRed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   91
               Top             =   1860
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label Label25 
               Caption         =   "MCO Fees"
               Height          =   255
               Left            =   240
               TabIndex        =   90
               Top             =   195
               Width           =   870
            End
            Begin VB.Label lblCovVisitLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9525
               TabIndex        =   89
               Top             =   900
               Width           =   1275
            End
            Begin VB.Label Label30 
               Caption         =   "Premium"
               Height          =   255
               Left            =   195
               TabIndex        =   88
               Top             =   1500
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label Label31 
               Caption         =   "Prem Refund"
               Height          =   255
               Left            =   195
               TabIndex        =   87
               Top             =   1845
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Label lblCovMCORed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   86
               Top             =   540
               Width           =   1035
            End
            Begin VB.Label Label33 
               Caption         =   "MCO Refund"
               Height          =   255
               Left            =   240
               TabIndex        =   85
               Top             =   555
               Width           =   975
            End
            Begin VB.Label Label34 
               Caption         =   "Dental Limit"
               Height          =   255
               Left            =   180
               TabIndex        =   84
               Top             =   2625
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Label Label35 
               Caption         =   "Spec. Limit"
               Height          =   255
               Left            =   300
               TabIndex        =   83
               Top             =   2010
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.Label lblCovAccumAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3765
               TabIndex        =   82
               Top             =   540
               Width           =   1035
            End
            Begin VB.Line Line5 
               X1              =   120
               X2              =   10920
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label Label37 
               Caption         =   "Accum Amt"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   81
               Top             =   555
               Width           =   975
            End
            Begin VB.Label Label39 
               Caption         =   "Visit Limit"
               Height          =   255
               Left            =   8400
               TabIndex        =   80
               Top             =   915
               Width           =   900
            End
            Begin VB.Label Label8 
               Caption         =   "Effective MCO"
               Height          =   255
               Left            =   240
               TabIndex        =   79
               Top             =   930
               Width           =   1095
            End
            Begin VB.Label lblCovMCOEff 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   78
               Top             =   900
               Width           =   1035
            End
            Begin VB.Label Label12 
               Caption         =   "Effective Prem"
               Height          =   255
               Left            =   195
               TabIndex        =   77
               Top             =   2250
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.Label lblCovPremEff 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   76
               Top             =   2220
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label Label13 
               Caption         =   "Balance Limit"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   75
               Top             =   930
               Width           =   975
            End
            Begin VB.Label lblCovBalLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3765
               TabIndex        =   74
               Top             =   900
               Width           =   1035
            End
            Begin VB.Label Label27 
               Caption         =   "Annual Limit"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   73
               Top             =   195
               Width           =   975
            End
            Begin VB.Label lblCovAnnLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3765
               TabIndex        =   72
               Top             =   180
               Width           =   1035
            End
         End
         Begin VB.ComboBox cboCovRelationship 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1140
            TabIndex        =   104
            Text            =   "W -Wife"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtSuppPLNCode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1140
            MaxLength       =   4
            TabIndex        =   106
            Top             =   530
            Width           =   1455
         End
         Begin VB.TextBox txtCovExclusion 
            Height          =   795
            Left            =   3720
            MaxLength       =   2500
            MultiLine       =   -1  'True
            TabIndex        =   126
            Top             =   540
            Width           =   3495
         End
         Begin VB.TextBox txtCOVCLNRemarks 
            Height          =   795
            Left            =   8400
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   105
            Top             =   540
            Width           =   2895
         End
         Begin MSMask.MaskEdBox txtSuppDOB 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   1140
            TabIndex        =   135
            Top             =   780
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox CboCovCardType 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "FrmMBMCLNEnq.frx":00B8
            Left            =   1140
            List            =   "FrmMBMCLNEnq.frx":00BA
            TabIndex        =   107
            Top             =   1020
            Width           =   1455
         End
         Begin MSMask.MaskEdBox txtCovEffDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   3720
            TabIndex        =   209
            Top             =   1300
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
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
         Begin MSMask.MaskEdBox txtCovExpDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   6000
            TabIndex        =   210
            Top             =   1300
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
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
         Begin VB.Label Label15 
            Caption         =   "Eff Date"
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
            Left            =   2880
            TabIndex        =   212
            Top             =   1320
            Width           =   1275
         End
         Begin VB.Label Label16 
            Caption         =   "Exp Date"
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
            Left            =   5160
            TabIndex        =   211
            Top             =   1320
            Width           =   915
         End
         Begin VB.Label Label47 
            Caption         =   "Supp. Warranty"
            Height          =   495
            Left            =   2880
            TabIndex        =   117
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label52 
            Caption         =   "Relationship"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   285
            Width           =   1005
         End
         Begin VB.Label Label41 
            Caption         =   "IC Num"
            Height          =   255
            Left            =   7440
            TabIndex        =   115
            Top             =   280
            Width           =   645
         End
         Begin VB.Label Label38 
            Caption         =   "Name"
            Height          =   255
            Left            =   2880
            TabIndex        =   114
            Top             =   280
            Width           =   660
         End
         Begin VB.Label Label92 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   7440
            TabIndex        =   113
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label66 
            Caption         =   "Plan Code"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   560
            Width           =   1125
         End
         Begin VB.Label Label4 
            Caption         =   "Date of Birth"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   830
            Width           =   975
         End
         Begin VB.Label lblCovAgeBand 
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   1140
            TabIndex        =   110
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label17 
            Caption         =   "AgeBand"
            Height          =   255
            Left            =   120
            TabIndex        =   109
            Top             =   1260
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label23 
            Caption         =   "Card Type"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   1080
            Width           =   915
         End
      End
      Begin VB.Frame FramePrincDet 
         Height          =   5370
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   11445
         Begin VB.TextBox txtDept 
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
            Left            =   6000
            MaxLength       =   150
            TabIndex        =   218
            Top             =   1560
            Width           =   5355
         End
         Begin VB.TextBox txtsplremark 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1545
            Left            =   5400
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   215
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox txtMBMAdd1 
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
            Left            =   840
            MaxLength       =   40
            TabIndex        =   150
            Top             =   240
            Width           =   4065
         End
         Begin VB.TextBox txtHISRemark 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1545
            Left            =   7890
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   136
            Top             =   2160
            Width           =   3495
         End
         Begin VB.TextBox txtCISRemarks 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1545
            Left            =   2880
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox txtMBMExclusion 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1545
            Left            =   120
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   2160
            Width           =   2580
         End
         Begin VB.Frame Frame2 
            Height          =   1320
            Left            =   120
            TabIndex        =   20
            Top             =   3960
            Width           =   11265
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8640
               TabIndex        =   207
               Top             =   915
               Width           =   600
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8640
               TabIndex        =   206
               Top             =   165
               Width           =   600
            End
            Begin VB.Label Label14 
               Caption         =   "Dental Ann Lmt"
               Height          =   255
               Index           =   6
               Left            =   7320
               TabIndex        =   205
               Top             =   165
               Width           =   1455
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8640
               TabIndex        =   204
               Top             =   525
               Width           =   600
            End
            Begin VB.Label lblMatBal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6360
               TabIndex        =   201
               Top             =   915
               Width           =   600
            End
            Begin VB.Label lblMeternityLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6360
               TabIndex        =   200
               Top             =   165
               Width           =   600
            End
            Begin VB.Label Label14 
               Caption         =   "Maternity Ann Lmt"
               Height          =   255
               Index           =   5
               Left            =   4920
               TabIndex        =   199
               Top             =   165
               Width           =   1455
            End
            Begin VB.Label lblMatAcc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6360
               TabIndex        =   198
               Top             =   525
               Width           =   600
            End
            Begin VB.Label Label26 
               Caption         =   "Maternity Bal Lmt"
               Height          =   255
               Index           =   2
               Left            =   4920
               TabIndex        =   197
               Top             =   915
               Width           =   1455
            End
            Begin VB.Label lblSpecAccumAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4150
               TabIndex        =   172
               Top             =   540
               Width           =   720
            End
            Begin VB.Label lblSpecAnnLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4150
               TabIndex        =   170
               Top             =   180
               Width           =   720
            End
            Begin VB.Label lblSpecBalLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4150
               TabIndex        =   169
               Top             =   915
               Width           =   720
            End
            Begin VB.Label Label5 
               Caption         =   "SP Acc Amt"
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   168
               Top             =   525
               Width           =   1455
            End
            Begin VB.Label Label94 
               Caption         =   "Dental Ext Limit"
               Height          =   255
               Left            =   2640
               TabIndex        =   149
               Top             =   2625
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.Label lblDentalExtLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   148
               Top             =   2595
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.Label lblMBMFFS 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   10260
               TabIndex        =   141
               Top             =   165
               Width           =   720
            End
            Begin VB.Label Label14 
               Caption         =   "FFS"
               Height          =   255
               Index           =   1
               Left            =   9360
               TabIndex        =   140
               Top             =   165
               Width           =   1095
            End
            Begin VB.Label lblTopupAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   10260
               TabIndex        =   138
               Top             =   525
               Width           =   720
            End
            Begin VB.Label Label11 
               Caption         =   "Visit Limit"
               Height          =   210
               Left            =   9360
               TabIndex        =   48
               Top             =   915
               Width           =   900
            End
            Begin VB.Line Line6 
               X1              =   120
               X2              =   10920
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label lblBalLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2400
               TabIndex        =   46
               Top             =   915
               Width           =   720
            End
            Begin VB.Label Label77 
               Caption         =   "Spec. Limit"
               Height          =   255
               Left            =   75
               TabIndex        =   45
               Top             =   2010
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.Label lblMBMMCORed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   42
               Top             =   540
               Width           =   465
            End
            Begin VB.Label Label58 
               Caption         =   "Prem Refund"
               Height          =   255
               Left            =   120
               TabIndex        =   41
               Top             =   1875
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Label lblPremium 
               Caption         =   "Premium"
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   1545
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label lblMBMVisitLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   10260
               TabIndex        =   39
               Top             =   915
               Width           =   720
            End
            Begin VB.Label lblMCOFeesEff 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   38
               Top             =   915
               Width           =   465
            End
            Begin VB.Label Label85 
               Caption         =   "Eff MCO"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   930
               Width           =   1095
            End
            Begin VB.Label lblMBMPremiumRed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   34
               Top             =   1845
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblMBMPremiumEff 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   33
               Top             =   2235
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblMBMPremium 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   32
               Top             =   1500
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Label lblMCOFees 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   31
               Top             =   180
               Width           =   465
            End
            Begin VB.Label lblDentalLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   30
               Top             =   2595
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.Label lblMBMSpecLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1365
               TabIndex        =   29
               Top             =   1965
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.Label lblMEPLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6720
               TabIndex        =   28
               Top             =   2595
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.Line Line1 
               X1              =   120
               X2              =   10920
               Y1              =   840
               Y2              =   840
            End
            Begin VB.Label lblFirstSpecLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1365
               TabIndex        =   27
               Top             =   2355
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.Label Label102 
               Caption         =   "First Spec. Limit"
               Height          =   255
               Left            =   75
               TabIndex        =   26
               Top             =   2385
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.Label Label105 
               Caption         =   "MEP Limit"
               Height          =   255
               Left            =   5400
               TabIndex        =   25
               Top             =   2610
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label lblAnnLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2400
               TabIndex        =   24
               Top             =   180
               Width           =   720
            End
            Begin VB.Label Label14 
               Caption         =   "Ann Lmt"
               Height          =   255
               Index           =   0
               Left            =   1680
               TabIndex        =   23
               Top             =   165
               Width           =   1095
            End
            Begin VB.Label lblAccumAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2400
               TabIndex        =   22
               Top             =   540
               Width           =   720
            End
            Begin VB.Label Label26 
               Caption         =   "Bal Lmt"
               Height          =   255
               Index           =   0
               Left            =   1680
               TabIndex        =   21
               Top             =   915
               Width           =   1095
            End
            Begin VB.Label Label79 
               Caption         =   "Dental Limit"
               Height          =   255
               Left            =   120
               TabIndex        =   44
               Top             =   2625
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.Label Label83 
               Caption         =   "Effective Prem "
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   2250
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.Label Label61 
               Caption         =   "MCO Ref"
               Height          =   255
               Left            =   195
               TabIndex        =   43
               Top             =   555
               Width           =   1095
            End
            Begin VB.Label lblMCOFee 
               Caption         =   "MCO"
               Height          =   255
               Left            =   195
               TabIndex        =   35
               Top             =   225
               Width           =   870
            End
            Begin VB.Label Label5 
               Caption         =   "Accu Amt"
               Height          =   255
               Index           =   0
               Left            =   1680
               TabIndex        =   47
               Top             =   540
               Width           =   1095
            End
            Begin VB.Label Label14 
               Caption         =   "SP Ann Lmt"
               Height          =   255
               Index           =   2
               Left            =   3240
               TabIndex        =   171
               Top             =   165
               Width           =   1455
            End
            Begin VB.Label Label26 
               Caption         =   "SP Bal Lmt"
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   173
               Top             =   915
               Width           =   1455
            End
            Begin VB.Label Label5 
               Caption         =   "Maternity Acc Amt"
               Height          =   255
               Index           =   2
               Left            =   4920
               TabIndex        =   202
               Top             =   525
               Width           =   1455
            End
            Begin VB.Label Label5 
               Caption         =   "Dental Acc Amt"
               Height          =   255
               Index           =   3
               Left            =   7320
               TabIndex        =   208
               Top             =   525
               Width           =   1455
            End
            Begin VB.Label Label26 
               Caption         =   "Dental Bal Lmt"
               Height          =   255
               Index           =   3
               Left            =   7320
               TabIndex        =   203
               Top             =   915
               Width           =   1455
            End
            Begin VB.Label Label24 
               Caption         =   "Top Up Amt"
               Height          =   255
               Left            =   9360
               TabIndex        =   139
               Top             =   525
               Width           =   1020
            End
         End
         Begin VB.TextBox txtInvoiceNo 
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
            Left            =   10320
            TabIndex        =   124
            Top             =   240
            Width           =   1035
         End
         Begin MSMask.MaskEdBox txtMBMPayorEffDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   8280
            TabIndex        =   52
            Top             =   240
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
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
         Begin MSMask.MaskEdBox txtMBMPayorExpDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   8280
            TabIndex        =   53
            Top             =   490
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
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
         Begin VB.ComboBox cboPLNCode 
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
            Left            =   5880
            Sorted          =   -1  'True
            TabIndex        =   54
            Top             =   195
            Width           =   1305
         End
         Begin VB.ComboBox cboINSCode 
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
            Left            =   5880
            Sorted          =   -1  'True
            TabIndex        =   55
            Text            =   "I"
            Top             =   480
            Width           =   1305
         End
         Begin VB.ComboBox CboCardType 
            Height          =   315
            ItemData        =   "FrmMBMCLNEnq.frx":00BC
            Left            =   4680
            List            =   "FrmMBMCLNEnq.frx":00BE
            TabIndex        =   56
            Top             =   4320
            Visible         =   0   'False
            Width           =   1665
         End
         Begin MSMask.MaskEdBox txtDOB 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   5880
            TabIndex        =   134
            Top             =   765
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
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
         Begin VB.TextBox txtMBMAdd2 
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
            Left            =   840
            MaxLength       =   40
            TabIndex        =   151
            Top             =   495
            Width           =   4065
         End
         Begin VB.TextBox txtMBMAdd3 
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
            Left            =   840
            MaxLength       =   40
            TabIndex        =   152
            Top             =   750
            Width           =   4065
         End
         Begin VB.TextBox txtMBMPostCode 
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
            Left            =   840
            MaxLength       =   6
            TabIndex        =   153
            Top             =   1005
            Width           =   1335
         End
         Begin VB.ComboBox cboCTYCode 
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
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   154
            Top             =   1005
            Width           =   1545
         End
         Begin VB.ComboBox cboSTACode 
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
            TabIndex        =   155
            Top             =   1260
            Width           =   4065
         End
         Begin MSMask.MaskEdBox txtInvoiceDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   10320
            TabIndex        =   122
            Top             =   490
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
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
         Begin VB.TextBox txtMBMGroupCompany 
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
            Left            =   5880
            MaxLength       =   150
            TabIndex        =   51
            Top             =   1000
            Width           =   5475
         End
         Begin MSMask.MaskEdBox txtMBMBordxDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   10080
            TabIndex        =   189
            Top             =   2160
            Visible         =   0   'False
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
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
         Begin MSMask.MaskEdBox txtMBMDateJoin 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   10080
            TabIndex        =   190
            Top             =   2160
            Visible         =   0   'False
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
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
         Begin VB.TextBox txtMBMPosition 
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
            Left            =   5880
            TabIndex        =   192
            Top             =   1260
            Width           =   1305
         End
         Begin VB.TextBox cboMBMSex 
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
            Left            =   9960
            MaxLength       =   30
            TabIndex        =   191
            Top             =   2520
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtMBMSubsidiary 
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
            Left            =   8160
            TabIndex        =   194
            Top             =   1260
            Width           =   3225
         End
         Begin VB.TextBox txtProduct 
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
            Left            =   840
            MaxLength       =   40
            TabIndex        =   213
            Top             =   1550
            Width           =   4065
         End
         Begin VB.Label Label43 
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
            Left            =   5040
            TabIndex        =   217
            Top             =   1560
            Width           =   1185
         End
         Begin VB.Label Label42 
            Caption         =   "Special Remarks"
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
            Left            =   5400
            TabIndex        =   216
            Top             =   1920
            Width           =   1590
         End
         Begin VB.Label Label1 
            Caption         =   "Product"
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
            Index           =   5
            Left            =   120
            TabIndex        =   214
            Top             =   1580
            Width           =   1140
         End
         Begin VB.Label Label1 
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
            Index           =   4
            Left            =   7320
            TabIndex        =   195
            Top             =   1320
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "Position"
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
            Index           =   31
            Left            =   5040
            TabIndex        =   193
            Top             =   1320
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "PostCode"
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
            Left            =   120
            TabIndex        =   159
            Top             =   1030
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "State"
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
            Left            =   120
            TabIndex        =   158
            Top             =   1300
            Width           =   465
         End
         Begin VB.Label Label57 
            Caption         =   "City/Town"
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
            Left            =   2280
            TabIndex        =   157
            Top             =   1030
            Width           =   1020
         End
         Begin VB.Label Label1 
            Caption         =   "Address"
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
            TabIndex        =   156
            Top             =   285
            Width           =   1140
         End
         Begin VB.Label Label2 
            Caption         =   "MBM Remarks"
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
            TabIndex        =   137
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label32 
            Caption         =   "Invoice No "
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
            Left            =   9360
            TabIndex        =   125
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label29 
            Caption         =   "Invoice Date"
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
            Left            =   9360
            TabIndex        =   123
            Top             =   480
            Width           =   915
         End
         Begin VB.Label Label16 
            Caption         =   "Exp Date"
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
            Left            =   7320
            TabIndex        =   69
            Top             =   525
            Width           =   915
         End
         Begin VB.Label Label15 
            Caption         =   "Eff Date"
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
            Left            =   7320
            TabIndex        =   68
            Top             =   270
            Width           =   1275
         End
         Begin VB.Label Label18 
            Caption         =   "Grp. Com"
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
            Left            =   5040
            TabIndex        =   67
            Top             =   1030
            Width           =   1185
         End
         Begin VB.Label Label21 
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
            Left            =   5040
            TabIndex        =   66
            Top             =   255
            Width           =   825
         End
         Begin VB.Label Label20 
            Caption         =   "Ins Type"
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
            Left            =   5040
            TabIndex        =   65
            Top             =   500
            Width           =   840
         End
         Begin VB.Label Label84 
            Caption         =   "Clinical Remarks"
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
            Left            =   2880
            TabIndex        =   64
            Top             =   1920
            Width           =   1590
         End
         Begin VB.Label Label40 
            Caption         =   "Principal Warranty "
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
            TabIndex        =   63
            Top             =   1920
            Width           =   1590
         End
         Begin VB.Label Label3 
            Caption         =   "Card Type"
            Height          =   255
            Left            =   4080
            TabIndex        =   62
            Top             =   4680
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label22 
            Caption         =   "DOB"
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
            Left            =   5040
            TabIndex        =   61
            Top             =   780
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "AgeBand"
            Height          =   255
            Left            =   1440
            TabIndex        =   60
            Top             =   4800
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblAgeBand 
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   4200
            TabIndex        =   59
            Top             =   4125
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label lblAdultChild 
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   4200
            TabIndex        =   58
            Top             =   4380
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label19 
            Caption         =   "Adult/Child"
            Height          =   255
            Left            =   1320
            TabIndex        =   57
            Top             =   4920
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView lstCovAdd 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   118
         Top             =   3480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   4048
         SortKey         =   4
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   32
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CovID"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "IC No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rel"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Warranty/Exclusion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Remarks"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "DOB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "SuppPlan"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "CurrAdmentStatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Prem"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "PremRefund"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "PremEff"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "MCO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "MCORefund"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "EffMCO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "AnnLmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "AccumLmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Bal Limit"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   19
            Text            =   "Visit Limit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Mep Lmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   21
            Text            =   "Dental Limit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   22
            Text            =   "Dent Ext Lmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Spec Lmt"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "1st Spec Lmt"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "Ageband"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "Adult/Child"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "SpecAnnLmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "SpecAccumAmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "SpecBalLmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "Eff Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "Exp Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstAdjustmentCis 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   119
         Top             =   360
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   9340
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Endt Bordx Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Adjustment Date"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Endt Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Endt Eff Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Adjusted By"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstMBMHis 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   120
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9340
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Member No"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Policy No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Effective Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Expired Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Cancel Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Data Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Policy Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Bordx Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Endt Bordx Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "MBM Reg Date"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstCaseList 
         Height          =   2115
         Left            =   -74880
         TabIndex        =   132
         Top             =   600
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   3731
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Mem No"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Patient"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Rel"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "DOV"
            Object.Width           =   1939
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Claimability"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Treatment Reason"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Case Status"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Bill Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "BordxAmt"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "FFS"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView lstSpecialistCases 
         Height          =   1875
         Left            =   -74880
         TabIndex        =   174
         Top             =   3480
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   3307
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Mem No"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Patient"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Rel"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "DOV"
            Object.Width           =   1939
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Claimability"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Treatment Reason"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Case Status"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Bill Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "BordxAmt"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "FFS"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Label Label36 
         Caption         =   "SP Listing"
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
         Index           =   8
         Left            =   -74880
         TabIndex        =   188
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "GP Listing"
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
         Index           =   5
         Left            =   -74880
         TabIndex        =   187
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblTotalSpeVisits 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
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
         Left            =   -68760
         TabIndex        =   181
         Top             =   5400
         Width           =   735
      End
      Begin VB.Label Label36 
         Caption         =   "Total Bill"
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
         Index           =   7
         Left            =   -67800
         TabIndex        =   186
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label Label36 
         Caption         =   "Total Bordx Amt"
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
         Index           =   6
         Left            =   -65640
         TabIndex        =   185
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label lblTotalSpecBordx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
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
         Left            =   -64440
         TabIndex        =   184
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label lblTotalSpeBill 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
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
         Left            =   -67080
         TabIndex        =   183
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "Total Visits"
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
         Left            =   -69600
         TabIndex        =   182
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label lblTotalVisit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
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
         Left            =   -70680
         TabIndex        =   167
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label36 
         Caption         =   "Total Visits"
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
         Left            =   -71520
         TabIndex        =   166
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblBillAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
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
         Left            =   -69000
         TabIndex        =   165
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label lblTotalBordx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
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
         Left            =   -66360
         TabIndex        =   164
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblTotalFFS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
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
         Left            =   -64440
         TabIndex        =   163
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "Total FFS"
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
         Left            =   -65160
         TabIndex        =   162
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label36 
         Caption         =   "Total Bordx Amt"
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
         Left            =   -67560
         TabIndex        =   161
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "Total Bill"
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
         Left            =   -69720
         TabIndex        =   160
         Top             =   2880
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      Height          =   435
      Left            =   120
      TabIndex        =   14
      Top             =   7800
      Width           =   11655
      Begin VB.CommandButton cmdExclusions 
         Caption         =   "&Exclusions"
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
         Left            =   9550
         TabIndex        =   147
         Top             =   120
         Width           =   960
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   10500
         TabIndex        =   15
         Top             =   120
         Width           =   960
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Member Enquiry - Clinical"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   0
      TabIndex        =   121
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "FrmMBMCLNEnq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADORSTMBM As New ADODB.Recordset
Dim tmpPolicyWording As String
Dim PictureFileName As String
Dim tmpLimitStatus As String
Dim tmpAccMaternity As Double
Dim sharelmtstatus As Boolean
Dim ASOInd As String
Dim DepID As String
Private mbActive                As Boolean
Private mlCurThumb              As Long
Private Const SRCCOPY           As Long = &HCC0020
Private Const STRETCH_HALFTONE  As Long = &H4&
Private Const SW_RESTORE        As Long = &H9&

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lpPt As PointAPI) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub ComboListing()

Const METHOD_NAME = "ComboListing"
On Error GoTo debugs

Dim REL As New ADODB.Recordset
Dim CheckItems As New ADODB.Recordset
        CheckItems.Open "SELECT CatName FROM CheckItems WHERE CatCode='CSMD'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not CheckItems.EOF
            CboCardType.AddItem CheckItems!CatName
            CboCovCardType.AddItem CheckItems!CatName
            CheckItems.MoveNext
        Loop
        CheckItems.Close
        Set CheckItems = Nothing
        
        REL.Open "select RELCode, RELDescription from  REL", medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do While Not REL.EOF
            cboCovRelationship.AddItem ChkStr(REL!RelCode) '& " - " & REL!RELdescription
            REL.MoveNext
        Loop
        REL.Close
        Set REL = Nothing
        
        cboStatus.AddItem "A" & " - " & "Active"
        cboStatus.AddItem "C" & " - " & "Cancel"
        cboStatus.AddItem "D" & " - " & "Deleted"
        cboStatus.AddItem "S" & " - " & "Suspended"
                    
        cboExpStatus.AddItem "Y" & " - " & "Yes"
        cboExpStatus.AddItem "N" & " - " & "No"
                    
        cboDataStatus.AddItem "A" & " - " & "Active"
        cboDataStatus.AddItem "D" & " - " & "Deleted"
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ClearEntries()
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is ComboBox Then
            cboINSCode.Text = ""
            cboPLNCode.Text = ""
            CboCardType.Text = ""
        ElseIf TypeOf ctl Is ListView Then
            ctl.ListItems.Clear
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.mask = ""
            ctl.Text = ""
            ctl.mask = "##/##/####"
        ElseIf TypeOf ctl Is Label Then
            lblTopupAmt = ""
            lblMBMPremium = ""
            lblMBMPremiumRed = ""
            lblMBMPremiumEff = ""
            lblMCOFees = ""
            lblMBMMCORed = ""
            lblMCOFeesEff = ""
            lblMBMVisitLimit = ""
            lblDentalLimit = ""
            lblMBMSpecLimit = ""
            lblBalLimit = ""
            lblAnnLmt = ""
            lblAccumAmt = ""
            lblMEPLimit = ""
            lblDentalExtLimit = ""
            lblFirstSpecLimit = ""
            lblSpecCovAnn = ""
            lblSpecCovAcc = ""
            lblSpecCovBal = ""
            
        End If
    Next ctl
    Call ClearCoveredEntry
End Sub


Private Sub ClearCoveredEntry()
    txtCovName = ""
    txtCovIcBcPp = ""
    txtSuppPLNCode = ""
    txtSuppDOB.mask = ""
    txtSuppDOB.Text = ""
    txtSuppDOB.mask = "##/##/####"
    txtCovExclusion = ""
    txtCOVCLNRemarks = ""
    cboCovRelationship.Text = "S - Son"
    CboCovCardType.Text = ""
    lblCovMCO = ""
    lblCovMCORed = ""
    lblCovMCOEff = ""
    lblCovPrem = ""
    lblCovPremRed = ""
    lblCovPremEff = ""
    lblCovMepLmt = ""
    lblCovBalLmt = ""
    lblCovAnnLmt = ""
    lblCovAccumAmt = ""
    lblCovVisitLmt = ""
    lblCovDentLmt = ""
    lblCovDentExtLmt = ""
    lblCovSpecLmt = ""
    lblCov1stSpecLmt = ""
End Sub


Private Sub cmdExclusions_Click()
    If Trim(txtMBMGroupCompany) = "" Then
        MsgBox "Please Check Group Company! ", vbOKOnly
    Else
        Load frmGeneralBenefitExcEnq
        frmGeneralBenefitExcEnq.cboGRPCompany = Trim(txtMBMGroupCompany)
        frmGeneralBenefitExcEnq.cmdSearch_Click
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrev_Click()
Const METHOD_NAME = "CmdPrev_Click"
On Error GoTo debugs

Dim MBM As New ADODB.Recordset
Dim sql As String
        sql = " select Top 1 MBMNumber from MBM WHERE HSCLNInd <> 'H' AND MBMNumber < '" & RemStr(txtMBMNumber) & "' " & _
              " Order by MBMnumber DESC "
        MBM.MaxRecords = 1
        MBM.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        If MBM.EOF = False Then
            txtMBMNumber.Text = MBM!MBMNumber
            MBM.Close
            Set MBM = Nothing
            cmdShow_Click
        Else
            MBM.Close
            Set MBM = Nothing
            MsgBox "First Record Reached! " & vbNewLine & "Not a valid Membership Number Format! ", vbCritical + vbOKOnly
        End If
        
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If


End Sub


Private Sub cmdSearchMBM_Click()
    Call ClearEntries
    frmSearchMBM.Source = 6
    frmSearchMBM.Show
End Sub

Private Sub Form_Load()
    Call ComboListing
    SSTab1.Tab = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub
Private Sub lstAdjustmentCis_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstAdjustmentCis.SortKey = ColumnHeader.Index - 1
    lstAdjustmentCis.Sorted = True
    If lstAdjustmentCis.SortOrder = lvwAscending Then
        lstAdjustmentCis.SortOrder = lvwDescending
    Else
        lstAdjustmentCis.SortOrder = lvwAscending
    End If
End Sub

Private Sub lstCaseList_DblClick()
    If lstCaseList.ListItems.Count Then
        frmInvoiceAdj.txtCaseNo = lstCaseList.SelectedItem.Text
        Call frmInvoiceAdj.CmdGo_Click
        frmInvoiceAdj.Show
    End If
End Sub

Private Sub lstCovAdd_Click()
Const METHOD_NAME = "lstCovAdd_Click"
On Error GoTo debugs

Dim MBMRef As New ADODB.Recordset
Dim strMBMRef As String
Dim TopupRec As New ADODB.Recordset
Dim tmpTopupAmt As Double
Dim tmpSuppFFS As Double
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim tmpMBMFFS As Double
Dim tempSuppCovID As String
Dim rst2 As New ADODB.Recordset
Dim strAppend2 As String

Dim StrPlnSql As String
Dim PLnRec As New ADODB.Recordset
Dim tmpSuppLmtStatus As String
tmpSuppFFS = 0
strAppend2 = ""
    If lstCovAdd.ListItems.Count > 0 Then
        txtCovName = lstCovAdd.SelectedItem.Text
        txtCovID = lstCovAdd.SelectedItem.SubItems(1)
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        cboCovRelationship.Text = lstCovAdd.SelectedItem.SubItems(3)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(5)
        txtCOVCLNRemarks = lstCovAdd.SelectedItem.SubItems(6)
        If Trim(lstCovAdd.SelectedItem.SubItems(7)) = "" Then
            txtSuppDOB = ""
        Else
            If IsDate(lstCovAdd.SelectedItem.SubItems(7)) Then
                txtSuppDOB = Format(lstCovAdd.SelectedItem.SubItems(7), "mm/dd/yyyy")
            Else
                txtSuppDOB = lstCovAdd.SelectedItem.SubItems(7)
            End If
        End If
        
        txtSuppPLNCode = lstCovAdd.SelectedItem.SubItems(8)
        CboCovCardType.Text = lstCovAdd.SelectedItem.SubItems(9)
        lblCovPrem = Format(lstCovAdd.SelectedItem.SubItems(10), "#,##0.00")
        lblCovPremRed = Format(lstCovAdd.SelectedItem.SubItems(11), "#,##0.00")
        lblCovPremEff = Format(lstCovAdd.SelectedItem.SubItems(12), "#,##0.00")
        lblCovMCO = Format(lstCovAdd.SelectedItem.SubItems(13), "#,##0.00")
        lblCovMCORed = Format(lstCovAdd.SelectedItem.SubItems(14), "#,##0.00")
        lblCovMCOEff = Format(lstCovAdd.SelectedItem.SubItems(15), "#,##0.00")
        lblCovAnnLmt = Format(lstCovAdd.SelectedItem.SubItems(16), "#,##0.00")
        lblCovAccumAmt = Format(lstCovAdd.SelectedItem.SubItems(17), "#,##0.00")
        lblCovBalLmt = Format(lstCovAdd.SelectedItem.SubItems(18), "#,##0.00")
        lblCovVisitLmt = Format(lstCovAdd.SelectedItem.SubItems(19), "#,##0.00")
        lblCovMepLmt = Format(lstCovAdd.SelectedItem.SubItems(20), "#,##0.00")
        lblCovDentLmt = Format(lstCovAdd.SelectedItem.SubItems(21), "#,##0.00")
        lblCovDentExtLmt = Format(lstCovAdd.SelectedItem.SubItems(22), "#,##0.00")
        lblCovSpecLmt = Format(lstCovAdd.SelectedItem.SubItems(23), "#,##0.00")
        lblCov1stSpecLmt = Format(lstCovAdd.SelectedItem.SubItems(24), "#,##0.00")
        lblSpecCovAnn = Format(lstCovAdd.SelectedItem.SubItems(27), "#,##0.00")
        lblSpecCovAcc = Format(lstCovAdd.SelectedItem.SubItems(28), "#,##0.00")
        lblSpecCovBal = Format(lstCovAdd.SelectedItem.SubItems(29), "#,##0.00")
        txtCovEffDate = Format(lstCovAdd.SelectedItem.SubItems(30), "dd/mm/yyyy")
        txtCovExpDate = Format(lstCovAdd.SelectedItem.SubItems(31), "dd/mm/yyyy")
        
        StrPlnSql = "Select SuppLimitStatus from AnnualLimit Where INSCode ='" & Trim(Mid(cboINSCode, 1, 1)) & "' And PLNCode ='" & Trim(Mid(cboPLNCode, 1, 4)) & _
                 "' AND AnnEffDate <= '" & Format(txtMBMPayorEffDate.Text, "mm/dd/yyyy") & "' order by AnnEffDate  desc"
        PLnRec.Open StrPlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        tmpSuppLmtStatus = ""
        Do While Not PLnRec.EOF
            tmpSuppLmtStatus = IIf(Len(Trim(PLnRec!SuppLimitStatus)), Trim(PLnRec!SuppLimitStatus), "A")
            PLnRec.MoveNext
        Loop
        PLnRec.Close
        Set PLnRec = Nothing
        
        If Len(txtMBMNumber) Then
           If Trim(tmpSuppLmtStatus) <> "A" Then
                    strAppend2 = ""
                If tmpSuppLmtStatus = "B" And Trim(txtCovID) = "01" Then
                    strAppend2 = strAppend2 & " CISCas. MBMnumber= '" & Trim(txtMBMNumber) & "' AND CoverID <> '00' "
                ElseIf tmpSuppLmtStatus = "C" Or tmpSuppLmtStatus = "" Then
                    strAppend2 = strAppend2 & "  CISCas.MBMnumber= '" & Trim(txtMBMNumber) & "' AND CoverID = '" & txtCovID & "'"
                ElseIf tmpSuppLmtStatus = "D" And Trim(txtCovID) = "01" Then
                    strAppend2 = strAppend2 & " CISCas. MBMnumber= '" & Trim(txtMBMNumber) & "' And CoverID  <> '00' "
                ElseIf tmpSuppLmtStatus = "E" Then
                    strAppend2 = strAppend2 & " CISCas. MBMnumber= '" & Trim(txtMBMNumber) & "' And CoverID  ='" + txtCovID + "' "
                End If
                If Trim(strAppend2) = "" Then
                    lblCovFFS = Format("0", "#,##0.00")
                Else
                    Set cmd.ActiveConnection = MediConnCISDB
                    cmd.CommandText = "CIS_FFS2013"
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandTimeout = 300
                    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend2)
                    cmd.Parameters.Append Prm1
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd.Execute
                    Do While Not rst2.EOF
                        tmpSuppFFS = CDbl(tmpSuppFFS) + IIf(IsNumeric(rst2!FFS), rst2!FFS, 0)
                        rst2.MoveNext
                    Loop
                    Set rst2 = Nothing
                    Set cmd = Nothing
                    Set Prm1 = Nothing
                    lblCovFFS = Format(tmpSuppFFS, "#,###.00")
                End If
            Else
                lblCovFFS = Format("0", "#,##0.00")
            End If
           If (tmpSuppLmtStatus = "B" Or tmpSuppLmtStatus = "D") And Trim(txtCovID) = "01" Then
                TopupRec.Open " SELECT SUM(TopupAmt) AS sumOfTopupAmt FROM MBMCISTopupHistory WHERE (MBMNumber = '" & txtMBMNumber & "') AND (MBMCovID = '01') GROUP BY MBMNumber ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                With TopupRec
                    If TopupRec.EOF = False Then
                        tmpTopupAmt = IIf(IsNumeric(TopupRec!sumOfTopupAmt), TopupRec!sumOfTopupAmt, "")
                        lblCovTopUp = Format(tmpTopupAmt, "#,##0.00")
                    Else
                        tmpTopupAmt = 0
                        lblCovTopUp = Format("0", "#,##0.00")
                    End If
                End With
                TopupRec.Close
                Set TopupRec = Nothing
                If IsNumeric(lblCovAnnLmt) Then
                    If IsNumeric(lblCovTopUp) Then
                        lblCovBalLmt = Format(CDbl(lblCovAnnLmt) + CDbl(tmpTopupAmt) - CDbl(lblCovAccumAmt), "#,##0.00")
                    Else
                        lblCovBalLmt = Format(CDbl(lblCovAnnLmt) - CDbl(lblCovAccumAmt), "#,##0.00")
                    End If
                Else
                    lblCovBalLmt = "N/A"
                End If
            ElseIf tmpSuppLmtStatus = "C" Or tmpSuppLmtStatus = "" Then
                TopupRec.Open " SELECT SUM(TopupAmt) AS sumOfTopupAmt FROM MBMCISTopupHistory WHERE (MBMNumber = '" & txtMBMNumber & "') AND (MBMCovID = '" & txtCovID & "') GROUP BY MBMNumber ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                With TopupRec
                    If TopupRec.EOF = False Then
                        tmpTopupAmt = IIf(IsNumeric(TopupRec!sumOfTopupAmt), TopupRec!sumOfTopupAmt, "")
                        lblCovTopUp = Format(tmpTopupAmt, "#,##0.00")
                    Else
                        tmpTopupAmt = 0
                        lblCovTopUp = Format("0", "#,##0.00")
                    End If
                End With
                TopupRec.Close
                Set TopupRec = Nothing
                If IsNumeric(lblCovAnnLmt) Then
                    If IsNumeric(lblCovTopUp) Then
                        lblCovBalLmt = Format(CDbl(lblCovAnnLmt) + CDbl(tmpTopupAmt) - CDbl(lblCovAccumAmt), "#,##0.00")
                    Else
                        lblCovBalLmt = Format(CDbl(lblCovAnnLmt) - CDbl(lblCovAccumAmt), "#,##0.00")
                    End If
                Else
                    lblCovBalLmt = "N/A"
                End If
            Else
                tmpTopupAmt = 0
                lblCovTopUp = Format("0", "#,##0.00")
            End If
        End If
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstMBMHis_DblClick()
    If lstMBMHis.ListItems.Count Then
        txtMBMNumber = lstMBMHis.SelectedItem.Text
        Call cmdShow_Click
    End If
End Sub

Private Sub lstPolicyWording_Click()
Const METHOD_NAME = "lstPolicyWording_Click"
On Error GoTo debugs

Dim lRet    As Long
Dim sFile   As String
    
    PictureFileName = SRVPath & "CLNPolicyWording\" & tmpPolicyWording & ""
    crt PictureFileName & "\" & lstPolicyWording.List(lstPolicyWording.ListIndex)
    sFile = PictureFileName & "\" & lstPolicyWording.List(lstPolicyWording.ListIndex)
    lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Sub crt(pth As String)
Dim f
f = FreeFile
Open App.path + "\temp.html" For Output As #f
    Print #f, "<html><body><center><img src=""" & pth & """></center></body></html>"
Close #f
DoEvents
'WebBrowser1.Navigate App.Path & "\temp.html"
End Sub

Private Sub lstSpecialistCases_DblClick()
    If lstSpecialistCases.ListItems.Count Then
        frmInvoiceAdj.txtCaseNo = lstSpecialistCases.SelectedItem.Text
        Call frmInvoiceAdj.CmdGo_Click
        frmInvoiceAdj.Show
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Const METHOD_NAME = "SSTab1_Click"
On Error GoTo debugs
    Screen.MousePointer = vbHourglass
    If SSTab1.Tab = 1 Then
        If Mid(cboINSCode, 1, 1) = "I" Or Mid(cboINSCode, 1, 1) = "G" Then
            MsgBox "This Plan does not have supplementary", vbOKOnly + vbCritical, DialogBoxTitle
            SSTab1.Tab = 0
        Else
            SSTab1.Tab = 1
        End If
    ElseIf SSTab1.Tab = 2 Then
            lstAdjustmentCis.ListItems.Clear
            If lstAdjustmentCis.ListItems.Count = 0 Then
                Call ADJListing
            End If
            lstAdjustmentCis.SetFocus
    ElseIf SSTab1.Tab = 3 Then
        If lstMBMHis.ListItems.Count = 0 Then
            Call MBMHISListing
        End If
    ElseIf SSTab1.Tab = 4 Then
        If Trim(txtMBMNumber) <> "" And lstCaseList.ListItems.Count = 0 Then
            Call ListCases
            Call ListSpecialistCases
        End If
    ElseIf SSTab1.Tab = 5 Then
        If Trim(tmpPolicyWording) <> "" Then
            Call LoadDoc
        End If
    End If
    Screen.MousePointer = Default
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub LoadDoc()
On Error Resume Next
Dim B As String, f As Integer, d As Integer, e As Integer
    Call addTheFiles
Exit Sub
End Sub

Sub addTheFiles()
Dim B As String
lstPolicyWording.Clear

    PictureFileName = SRVPath & "CLNPolicyWording\" & tmpPolicyWording & ""
    '\Medical Rpt"
    B = Dir(PictureFileName & "\*.*")
    If B <> "" Then
        If Right(B, 4) = ".jpg" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".tif" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".gif" Then
            lstPolicyWording.AddItem B
        End If
        
        If Right(B, 4) = ".JPG" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".TIF" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".GIF" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".pdf" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".PDF" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".txt" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".TXT" Then
            lstPolicyWording.AddItem B
        End If
        
        If Right(B, 4) = "docx" Then
                lstPolicyWording.AddItem B
        End If
        
        If Right(B, 4) = "xlsx" Then
            lstPolicyWording.AddItem B
        End If
        
        If Right(B, 4) = "XLSX" Then
            lstPolicyWording.AddItem B
        End If
        
    Else
        Exit Sub
    End If
    Do: DoEvents
        B = Dir
        If B <> "" Then
            If Right(B, 4) = ".jpg" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = ".tif" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = ".gif" Then
                lstPolicyWording.AddItem B
            End If
            
            If Right(B, 4) = ".JPG" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = ".TIF" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = ".GIF" Then
                lstPolicyWording.AddItem B
            End If
            
            If Right(B, 4) = ".PDF" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = ".pdf" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = ".txt" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = ".TXT" Then
                lstPolicyWording.AddItem B
            End If
            
            If Right(B, 4) = "xlsx" Then
                lstPolicyWording.AddItem B
            End If
            
            If Right(B, 4) = "docx" Then
                lstPolicyWording.AddItem B
            End If
            
            If Right(B, 4) = "XLSX" Then
                lstPolicyWording.AddItem B
            End If
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub ListCases()
Const METHOD_NAME = "ListCases"
On Error GoTo debugs

Dim CASList As ListItem
Dim StrSql As String
Dim CAS As New ADODB.Recordset


    lblTotalVisit = 0
    lblBillAmt = 0
    lblTotalBordx = 0
    lblTotalFFS = 0
    
    StrSql = " Select CASNumber, MBMCName, " & _
             " ReasonDesc, CoverID, CLMMCOProcFee, " & _
             " DateVisit, MBMCRELCode, CLMTotalActual,CLMBordxAmt, " & _
             " CASStatus, MBMName, MBMNumber, Claimability  "
    StrSql = StrSql & " From VIEW_CISMBMInquiry_GPCASE where 0=0"
    StrSql = StrSql & " AND MBMNumber = '" & RemStr(txtMBMNumber) & "' "
            
    CAS.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If CAS.RecordCount <> 0 Then
        Screen.MousePointer = vbHourglass
        lstCaseList.ListItems.Clear
        Do Until CAS.EOF
            Set CASList = lstCaseList.ListItems.Add(, , CAS!CasNumber)
                CASList.SubItems(1) = CAS!MBMNumber
                
                If CAS!coverid = "00" Then
                    If Len(CAS!MBMName) Then
                        CASList.SubItems(2) = Trim(CAS!MBMName)
                    End If
                    CASList.SubItems(3) = "P"
                Else
                    If Len(CAS!MBMCName) Then
                        CASList.SubItems(2) = Trim(CAS!MBMCName)
                    End If
                    If Len(CAS!MBMCRELCode) Then
                        CASList.SubItems(3) = CAS!MBMCRELCode
                    End If
                End If
                
                If Len(Trim(CAS!DateVisit)) Then
                    CASList.SubItems(4) = CAS!DateVisit
                End If
                
                If Len(Trim(CAS!Claimability)) Then
                    CASList.SubItems(5) = CAS!Claimability
                End If
                If Trim(CAS!ReasonDesc) <> "" Then
                    CASList.SubItems(6) = Trim(CAS!ReasonDesc)
                End If
                If Len(CAS!CASStatus) Then
                    CASList.SubItems(7) = CAS!CASStatus
                End If
                
                If IsNumeric(Trim(CAS!CLMTotalActual)) Then
                    CASList.SubItems(8) = Format(CAS!CLMTotalActual, "#,##0.00")
                Else
                    CASList.SubItems(8) = Format(0, "#,##0.00")
                End If
                
                If IsNumeric(Trim(CAS!CLMBordxAmt)) Then
                    CASList.SubItems(9) = Format(CAS!CLMBordxAmt, "#,##0.00")
                Else
                    CASList.SubItems(9) = Format(0, "#,##0.00")
                End If
                
                If IsNumeric(Trim(CAS!CLMMCOProcFee)) Then
                    CASList.SubItems(10) = Format(CAS!CLMMCOProcFee, "#,##0.00")
                Else
                    CASList.SubItems(10) = Format(0, "#,##0.00")
                End If
                lblTotalVisit = CDbl(lblTotalVisit) + 1
                lblBillAmt = CDbl(lblBillAmt) + CDbl(CAS!CLMTotalActual)
                
                If (CAS!CASStatus <> "D") Then
                    If IsNumeric(Trim(CAS!CLMBordxAmt)) Then
                    lblTotalBordx = CDbl(lblTotalBordx) + CDbl(CAS!CLMBordxAmt)
                    Else
                    lblTotalBordx = CDbl(lblTotalBordx) + CDbl(0)
                    End If
                End If
                
                lblTotalFFS = CDbl(lblTotalFFS) + CDbl(CAS!CLMMCOProcFee)
                
                CAS.MoveNext
        Loop
        Screen.MousePointer = vbDefault
    End If
    CAS.Close
    Set CAS = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ListSpecialistCases()
Const METHOD_NAME = "ListSpecialistCases"
On Error GoTo debugs

Dim CASList As ListItem
Dim StrSql As String
Dim CAS As New ADODB.Recordset


    lblTotalSpeVisits = 0
    lblTotalSpeBill = 0
    lblTotalSpecBordx = 0
    'lblTotalFFS = 0
    
    StrSql = " Select CASNumber, MBMCName, " & _
             " ReasonDesc, CoverID, CLMMCOProcFee, " & _
             " DateVisit, MBMCRELCode, CLMTotalActual,CLMBordxAmt, " & _
             " CASStatus, MBMName, MBMNumber, Claimability  "
    StrSql = StrSql & " From VIEW_CISMBMInquiry_Spe_CASE where 0=0"
    StrSql = StrSql & " AND MBMNumber = '" & RemStr(txtMBMNumber) & "' "
            
    CAS.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If CAS.RecordCount <> 0 Then
        Screen.MousePointer = vbHourglass
        lstSpecialistCases.ListItems.Clear
        Do Until CAS.EOF
            Set CASList = lstSpecialistCases.ListItems.Add(, , CAS!CasNumber)
                CASList.SubItems(1) = CAS!MBMNumber
                
                If CAS!coverid = "00" Then
                    If Len(CAS!MBMName) Then
                        CASList.SubItems(2) = Trim(CAS!MBMName)
                    End If
                    CASList.SubItems(3) = "P"
                Else
                    If Len(CAS!MBMCName) Then
                        CASList.SubItems(2) = Trim(CAS!MBMCName)
                    End If
                    If Len(CAS!MBMCRELCode) Then
                        CASList.SubItems(3) = CAS!MBMCRELCode
                    End If
                End If
                
                If Len(Trim(CAS!DateVisit)) Then
                    CASList.SubItems(4) = CAS!DateVisit
                End If
                
                If Len(Trim(CAS!Claimability)) Then
                    CASList.SubItems(5) = CAS!Claimability
                End If
                If Trim(CAS!ReasonDesc) <> "" Then
                    CASList.SubItems(6) = Trim(CAS!ReasonDesc)
                End If
                If Len(CAS!CASStatus) Then
                    CASList.SubItems(7) = CAS!CASStatus
                End If
                
                If IsNumeric(Trim(CAS!CLMTotalActual)) Then
                    CASList.SubItems(8) = Format(CAS!CLMTotalActual, "#,##0.00")
                Else
                    CASList.SubItems(8) = Format(0, "#,##0.00")
                End If
                
                'If IsNumeric(Trim(CAS!CLMBordxAmt)) Then
                '    CASList.SubItems(9) = Format(CAS!CLMBordxAmt, "#,##0.00")
                'Else
                '    CASList.SubItems(9) = Format(0, "#,##0.00")
                'End If
                If IsNumeric(Trim(CAS!CLMTotalActual)) Then
                    CASList.SubItems(9) = Format(CAS!CLMTotalActual, "#,##0.00")
                Else
                    CASList.SubItems(9) = Format(0, "#,##0.00")
                End If
                
                If IsNumeric(Trim(CAS!CLMMCOProcFee)) Then
                    CASList.SubItems(10) = Format(CAS!CLMMCOProcFee, "#,##0.00")
                Else
                    CASList.SubItems(10) = Format(0, "#,##0.00")
                End If
                lblTotalSpeVisits = CDbl(lblTotalSpeVisits) + 1
                
                Dim TtlActual As Double
                TtlActual = 0
                If IsNumeric(CAS!CLMTotalActual) Then
                    TtlActual = CDbl(CAS!CLMTotalActual)
                Else
                    TtlActual = 0
                End If
                
                lblTotalSpeBill = CDbl(lblTotalSpeBill) + TtlActual
                'lblTotalSpecBordx = CDbl(lblTotalSpecBordx) + CDbl(CAS!CLMBordxAmt)
                lblTotalSpecBordx = CDbl(lblTotalSpecBordx) + TtlActual

                CAS.MoveNext
        Loop
        Screen.MousePointer = vbDefault
    End If
    CAS.Close
    Set CAS = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub txtCOVCLNRemarks_LostFocus()
    txtCOVCLNRemarks = ChkStr(txtCOVCLNRemarks)
End Sub


Private Sub txtMBMNumber_GotFocus()
    cmdShow.Default = True
End Sub

Private Sub txtMBMNumber_LostFocus()
    txtMBMNumber = ChkStr(txtMBMNumber)
    cmdShow.Default = False
End Sub

Private Sub txtCovExclusion_Click()
    KeyPreview = False
End Sub

Private Sub txtCovExclusion_LostFocus()
    KeyPreview = True
    txtCovExclusion = ChkStr(txtCovExclusion)
End Sub

Private Sub CmdNext_Click()
Const METHOD_NAME = "CmdNext_Click"
On Error GoTo debugs

    Dim MBM As New ADODB.Recordset
    Dim sql As String
    sql = " Select Top 1 MBMNumber from MBM Where HSCLNInd <> 'H' AND MBMNumber > '" & RemStr(txtMBMNumber) & "' " & _
        " Order by MBMnumber "
    MBM.MaxRecords = 1
    MBM.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    If MBM.EOF = False Then
        txtMBMNumber.Text = MBM!MBMNumber
            
        MBM.Close
        Set MBM = Nothing
        cmdShow_Click
    Else
        MBM.Close
        Set MBM = Nothing
        MsgBox " Last Record Reached! ", vbCritical + vbOKOnly
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Public Sub cmdShow_Click()
Const METHOD_NAME = "cmdShow_Click"
On Error GoTo debugs
Dim ScanCode As String
Dim MBMNumber As String
Dim tmpRecCnt As Integer
        MBMNumber = txtMBMNumber
        sharelmtstatus = False
        Call ClearEntries
        SSTab1.Tab = 0
        txtMBMNumber = MBMNumber
        If txtMBMNumber = "" Or Len(txtMBMNumber) < 8 Then
            MsgBox "Please Enter Membership No. Atleast first 8 characters required.", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMNumber.SetFocus
        'New Con
        ElseIf (Mid(Trim(txtMBMNumber), 1, 2)) <> "DC" And (Mid(Trim(txtMBMNumber), 1, 2) <> "SL") And (Len(txtMBMNumber) <= 11 And Mid(Trim(txtMBMNumber), 12, 1) <> "*") Then
            tmpRecCnt = 1
            frmSearchMBM05.Source = 3
            frmSearchMBM05.Show
            tmpRecCnt = frmSearchMBM05.SearchRecCnt
            If frmSearchMBM05.SearchRecCnt < 1 Then
                Unload frmSearchMBM05
            ElseIf frmSearchMBM05.SearchRecCnt = 1 Then
                txtMBMNumber = frmSearchMBM05.tmpMBMNo
                Unload frmSearchMBM05
                Call FindMBM
            End If
        'New Con
        Else
            ScanCode = Trim(txtMBMNumber)
                Screen.MousePointer = vbHourglass
                Call FindMBM
                Screen.MousePointer = Default
                
            Dim MBMNo As String
            If InStr(MBMNumber, "-") Then
                Dim sArr() As String
                sArr = Split(MBMNumber, "-")
                MBMNo = sArr(0)
            Else
                MBMNo = MBMNumber
            End If
            Dim StrSqlMsgProm As String
            Dim MsgProm As ADODB.Recordset
            Set MsgProm = New ADODB.Recordset
            StrSqlMsgProm = "Select PLNPolicyWording,productname,PLNRBType,PLNDescription  from MBM inner join PLN on PLN.PLNCode = MBM.PLNCode where MBMNumber ='" + MBMNo + "' "
            MsgProm.Open StrSqlMsgProm, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            If MsgProm.EOF = False Then
                txtProduct = IIf(Trim(Len(MsgProm!ProductName)), MsgProm!ProductName, "") & "(" & IIf(Trim(Len(MsgProm!PLNDescription)), MsgProm!PLNDescription, "") & ")"
            Else
                txtProduct = ""
            End If
            MsgProm.Close
            Set MsgProm = Nothing
    End If
    If Len(txtsplremark) Then
        txtsplremark.BackColor = vbYellow
        MsgBox txtsplremark, vbOKOnly + vbInformation
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub FindMBM()
Const METHOD_NAME = "FindMBM"
On Error GoTo debugs


Dim sql As String
        
    sql = "Exec SearchCISMemberProc '" & Trim(txtMBMNumber) & "'"
    
    ADORSTMBM.Open sql, medixmasterconn, adOpenForwardOnly, adLockOptimistic
    
    If Not ADORSTMBM.EOF Then
        Call showMBM
        If Trim(ADORSTMBM!INSType) = "F" Or Trim(ADORSTMBM!INSType) = "H" Then
            If Len(Trim(ADORSTMBM!MBMCCoverID)) And Trim(ADORSTMBM!MBMCCoverID) <> "00" Then
                Call showMBMCov
            End If
        End If
    Else
        MsgBox "No record found!", vbCritical
    End If
    ADORSTMBM.Close
    Set ADORSTMBM = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub showMBM()
Const METHOD_NAME = "showMBM"
On Error GoTo debugs

Dim MBMRef As New ADODB.Recordset
Dim strMBMRef As String
Dim tmpAnnLmt As String
Dim MBMO As New ADODB.Recordset
Dim tmpVisitLmt As String
Dim AnnLmtRec As New ADODB.Recordset
Dim SpecLmtRec As New ADODB.Recordset
Dim MBMOTwo As New ADODB.Recordset
Dim MBMOThree As New ADODB.Recordset
Dim TopupRec As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim tmpTopupAmt As String
Dim tmpMBMCoverId As String
Dim strAppend As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim tmpMBMFFS As Double
Dim tmpSpecMBMFFS As Double
Dim tmpMBMCLMAmt As Double
Dim tempSuppCovID As String
Dim tmpSpecLimitStatus As String
Dim tmpSpecAccLmt As Double
Dim tmpSpecLmt As String
Dim SpecLmt As New ADODB.Recordset
Dim strsqlSpec As String
Dim tmpMaternityInd As String
Dim tmpDentalInd As String
Dim dblMeternity As Double
Dim dblDental As Double
Dim OtherDed As Double
Dim thyroid As Double
Dim AccMaternity As Double
Dim sMBMNo As String
Dim SCovId As String
Dim sArr() As String
Dim slblStatus As String
Dim BordxAmt As Double
Dim isPrinDepGPShared As Boolean

    BordxAmt = 0
    AccMaternity = 0
    tmpAccMaternity = 0
    tmpMBMFFS = 0
    tmpSpecMBMFFS = 0
    lblMeternityLmt = 0
    tmpSpecAccLmt = 0
    With ADORSTMBM
    tmpMaternityInd = ""
    
        If Trim(!MBMCLNStatus) = "A" Then
            cboStatus.ListIndex = 0
        ElseIf Trim(!MBMCLNStatus) = "C" Then
            cboStatus.ListIndex = 1
        ElseIf Trim(!MBMCLNStatus) = "D" Then
            cboStatus.ListIndex = 2
        ElseIf Trim(!MBMCLNStatus) = "S" Then
            cboStatus.ListIndex = 3
        End If
        If Trim(!MBMCLNExpiryStatus) = "N" Then
            cboExpStatus.ListIndex = 0
        Else
            cboExpStatus.ListIndex = 1
        End If
        If Trim(!MBMCLNDataStatus) = "A" Then
            cboDataStatus.ListIndex = 0
        ElseIf Trim(!MBMCLNDataStatus) = "D" Then
            cboDataStatus.ListIndex = 1
        End If
        txtPAYCode = !PAYCode
        txtMBMName = !MBMName
        txtMBMPolNo = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
        txtMBMIcBcPP = IIf(Len(!MBMIcBcPp), !MBMIcBcPp, "")
        cboINSCode = !INSType
        cboPLNCode = !PLNType
        ASOInd = IIf(Len(!ASOInd), !ASOInd, "N")
        tmpMBMCoverId = !MBMCOVID
        txtMBMAdd1 = IIf(Len(!MBMAddress1), !MBMAddress1, "")
        txtMBMAdd2 = IIf(Len(!MBMAddress2), !MBMAddress2, "")
        txtMBMAdd3 = IIf(Len(!MBMAddress3), !MBMAddress3, "")
        cboCTYCode = IIf(Len(!CTYCode), !CTYCode, "")
        txtMBMPostCode = IIf(Len(!MBMPostCode), !MBMPostCode, "")
        cboSTACode = IIf(Len(!STACode), !STACode, "")
        txtMBMSubsidiary = IIf(Len(!Subsidiary), !Subsidiary, "")
        txtDept = IIf(Len(!MBMDepartment), !MBMDepartment, "")
        txtMBMGroupCompany = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
        tmpPolicyWording = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
        txtHISRemark = IIf(Len(!MBMRemarks), !MBMRemarks, "")
        txtsplremark = IIf(Len(!PLNSpecialRemark), !PLNSpecialRemark, "")
        txtDOB = IIf(Len(!MBMDOB), !MBMDOB, "__/__/____")
        lblLimitStatus = "N"
        tmpLimitStatus = ""
        lblAdultChild = !MBMAdultChild
        AnnLmtRec.Open "SELECT Top 1 SuppLimitStatus FROM AnnualLimit WHERE PLNCode='" & Trim(!PLNType) & _
        "' AND INSCode='" & Trim(!INSType) & "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        If AnnLmtRec.EOF = False Then
            lblLimitStatus = IIf(Len(AnnLmtRec!SuppLimitStatus), AnnLmtRec!SuppLimitStatus, "A")
            tmpLimitStatus = IIf(Len(AnnLmtRec!SuppLimitStatus), AnnLmtRec!SuppLimitStatus, "A")
        End If
        AnnLmtRec.Close
        Set AnnLmtRec = Nothing
        slblStatus = lblLimitStatus
        
        If Trim(lblLimitStatus) = "N" Then
            lblLimitStatus = "Unlimited Limit"
        ElseIf Trim(lblLimitStatus) = "A" Then
            lblLimitStatus = "Share Limit For Family"
        ElseIf Trim(lblLimitStatus) = "B" Then
            lblLimitStatus = "Share Limit For Supplementary"
        ElseIf Trim(lblLimitStatus) = "C" Then
            lblLimitStatus = "Separate Limit For Supplementary and Principle"
        ElseIf Trim(lblLimitStatus) = "D" Then
            lblLimitStatus = "Share For Supplementary and Full For Principle"
        ElseIf Trim(lblLimitStatus) = "E" Then
            lblLimitStatus = "Separate For Supplementary and Full For Principle"
        ElseIf Trim(lblLimitStatus) = "F" Then
            lblLimitStatus = "Principle-Seperate" & vbNewLine & "Spouse-Seperate" & vbNewLine & "Children-Seperate for each" & vbNewLine
        End If
        
        tmpTopupAmt = 0
        lblTopupAmt = "N/A"
        If tmpLimitStatus = "A" Then
            TopupRec.Open "SELECT Top 1 sum(TopupAmt) as sumOfTopupAmt from MBMCISTopupHistory where MBMNumber='" & txtMBMNumber & "' group by MBMNumber", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            With TopupRec
                If TopupRec.EOF = False Then
                    tmpTopupAmt = IIf(IsNumeric(TopupRec!sumOfTopupAmt), TopupRec!sumOfTopupAmt, "")
                    lblTopupAmt = Format(tmpTopupAmt, "#,##0.00")
                End If
            End With
            TopupRec.Close
            Set TopupRec = Nothing
        ElseIf tmpLimitStatus = "B" Then
            TopupRec.Open " SELECT Top 1 SUM(TopupAmt) AS sumOfTopupAmt FROM MBMCISTopupHistory WHERE (MBMNumber = '" & txtMBMNumber & "') AND (MBMCovID = '00') GROUP BY MBMNumber ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            With TopupRec
                If TopupRec.EOF = False Then
                    tmpTopupAmt = IIf(IsNumeric(TopupRec!sumOfTopupAmt), TopupRec!sumOfTopupAmt, "")
                    lblTopupAmt = Format(tmpTopupAmt, "#,##0.00")
                End If
            End With
            TopupRec.Close
            Set TopupRec = Nothing
        ElseIf tmpLimitStatus = "C" Then
            TopupRec.Open " SELECT Top 1 SUM(TopupAmt) AS sumOfTopupAmt FROM MBMCISTopupHistory WHERE (MBMNumber = '" & txtMBMNumber & "') AND (MBMCovID = '00') GROUP BY MBMNumber ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            With TopupRec
                If TopupRec.EOF = False Then
                    tmpTopupAmt = IIf(IsNumeric(TopupRec!sumOfTopupAmt), TopupRec!sumOfTopupAmt, "")
                    lblTopupAmt = Format(tmpTopupAmt, "#,##0.00")
                End If
            End With
            TopupRec.Close
            Set TopupRec = Nothing
        ElseIf tmpLimitStatus = "E" Then
            TopupRec.Open " SELECT Top 1 SUM(TopupAmt) AS sumOfTopupAmt FROM MBMCISTopupHistory WHERE (MBMNumber = '" & txtMBMNumber & "') AND (MBMCovID = '00') GROUP BY MBMNumber ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            With TopupRec
                If TopupRec.EOF = False Then
                    tmpTopupAmt = IIf(IsNumeric(TopupRec!sumOfTopupAmt), TopupRec!sumOfTopupAmt, "")
                    lblTopupAmt = Format(tmpTopupAmt, "#,##0.00")
                End If
            End With
            TopupRec.Close
            Set TopupRec = Nothing
        ElseIf tmpLimitStatus = "D" Then
            tmpTopupAmt = 0
            lblTopupAmt = "N/A"
        ElseIf tmpLimitStatus = "F" Then
            TopupRec.Open " SELECT Top 1 SUM(TopupAmt) AS sumOfTopupAmt FROM MBMCISTopupHistory WHERE (MBMNumber = '" & txtMBMNumber & "') AND (MBMCovID = '00') GROUP BY MBMNumber ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            With TopupRec
                If TopupRec.EOF = False Then
                    tmpTopupAmt = IIf(IsNumeric(TopupRec!sumOfTopupAmt), TopupRec!sumOfTopupAmt, "")
                    lblTopupAmt = Format(tmpTopupAmt, "#,##0.00")
                End If
            End With
            TopupRec.Close
            Set TopupRec = Nothing
        ElseIf tmpLimitStatus = "" Then
            tmpTopupAmt = 0
            lblTopupAmt = "N/A"
        Else
            tmpTopupAmt = 0
            lblTopupAmt = "N/A"
        End If
        
        txtMBMPayorEffDate = !MBMPolEffDate
        txtMBMPayorExpDate = !MBMPolExpDate
        lblAgeBand = MemberCallAge(txtMBMPayorEffDate, txtDOB)
        CboCardType.Text = IIf(Len(!MBMCardType), !MBMCardType, "")
        txtMBMExclusion = IIf(Len(!MBMCLNWarranty), !MBMCLNWarranty, "")
        txtCISRemarks = IIf(Len(!MBMCLNRemarks), !MBMCLNRemarks, "")
        txtInvoiceDate = IIf(Len(!MBMCInvoiceDate), !MBMCInvoiceDate, "__/__/____")
        txtInvoiceNo = IIf(Len(!MBMCInvoiceNo), !MBMCInvoiceNo, "")
        
        lblMBMPremium = Format(!MBMCLNPremium, "#,##0.00")
        lblMCOFees = Format(!MBMCLNMCOFee, "#,##0.00")
        isPrinDepGPShared = !isPrinDepGPShared
        
        Dim SqlQry As String
        SqlQry = "select AnnualLimit.AnnMaternityLmt from AnnualLimit inner join PLN on PLN.PLNCode = AnnualLimit.PLNCode where PLN.PLNMaternityInd ='Y' and  PLN.PLNCode='" + Trim(!PLNType) + "'"
        SpecLmtRec.Open SqlQry, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        If SpecLmtRec.EOF = False Then
            lblMeternityLmt = IIf(Len(SpecLmtRec!AnnMaternityLmt), SpecLmtRec!AnnMaternityLmt, "")
        End If
        SpecLmtRec.Close
        Set SpecLmtRec = Nothing

        Call CallLimit("P", tmpAnnLmt, tmpVisitLmt, tmpSpecLmt, tmpLimitStatus)
        
        lblSpecAnnLmt = tmpSpecLmt
        
        tmpSpecLimitStatus = ""
        SpecLmtRec.Open "SELECT SuppLimitStatus FROM SpecialistLimit WHERE PLNCode='" & Trim(!PLNType) & _
        "' AND INSCode='" & Trim(!INSType) & "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        If SpecLmtRec.EOF = False Then
            tmpSpecLimitStatus = IIf(Len(SpecLmtRec!SuppLimitStatus), SpecLmtRec!SuppLimitStatus, "A")
        End If
        SpecLmtRec.Close
        Set SpecLmtRec = Nothing
        
        If Trim(lblSpecAnnLmt) = "N/A" Then
            If Len(txtMBMNumber) Then
                If Len(txtMBMNumber) Then
                   If InStr(1, txtMBMNumber, "-") Then
                      sArr = Split(txtMBMNumber, "-")
                      sMBMNo = sArr(0)
                      SCovId = sArr(1)
                   Else
                      sMBMNo = txtMBMNumber
                      SCovId = "00"
                   End If
                End If
                strAppend = ""
                
                
                    Call tmpMaternitylimit(" CISCas.MBMNumber = '" & Trim(sMBMNo) & "'", tmpLimitStatus)
                
                
                If isPrinDepGPShared Then
                    strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(sMBMNo) & "' "
                Else
                    If tmpLimitStatus = "A" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(sMBMNo) & "'"
                    ElseIf tmpLimitStatus = "B" Then
                        If SCovId = "00" Then
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(sMBMNo) & "' AND CoverID = '00' "
                        Else
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(sMBMNo) & "' AND CoverID <> '00' "
                        End If
                    ElseIf (tmpLimitStatus = "C" Or tmpLimitStatus = "F") Or tmpLimitStatus = "" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(sMBMNo) & "' AND CoverID ='00' "
                    ElseIf tmpLimitStatus = "D" Then
                        If SCovId = "00" Then
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(sMBMNo) & "' AND CoverID = '00' "
                        Else
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(sMBMNo) & "' AND CoverID <> '00' "
                        End If
                    ElseIf tmpLimitStatus = "E" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(sMBMNo) & "' AND CoverID ='" + SCovId + "' "
                    End If
                End If
                'If tmpLimitStatus = "A" Then
                '    strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(txtMBMNumber) & "'"
                'ElseIf tmpLimitStatus = "B" Then
                '    strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(txtMBMNumber) & "' AND CoverID ='00' "
                'ElseIf tmpLimitStatus = "C" Or tmpLimitStatus = "" Then
                '    strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(txtMBMNumber) & "' AND CoverID ='00' "
                'ElseIf tmpLimitStatus = "D" Then
                '    strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(txtMBMNumber) & "' AND CoverID ='00' "
                'ElseIf tmpLimitStatus = "E" Then
                '    strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(txtMBMNumber) & "' AND CoverID ='00' "
                'End If
            End If
            
            tmpSpecLimitStatus = ""
            SpecLmtRec.Open "SELECT Top 1 PLNMaternityInd,PDentalStatus FROM PLN WHERE PLNCode='" & Trim(!PLNType) & _
            "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            If SpecLmtRec.EOF = False Then
                tmpMaternityInd = IIf(Len(SpecLmtRec!PLNMaternityInd), SpecLmtRec!PLNMaternityInd, "")
                tmpDentalInd = IIf(Len(SpecLmtRec!PDentalStatus), SpecLmtRec!PDentalStatus, "")
            End If
            SpecLmtRec.Close
            Set SpecLmtRec = Nothing
            If Trim(strAppend) = "" Then
                tmpMBMFFS = 0
                tmpMBMCLMAmt = 0
            Else
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "CIS_FFS2013"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd.Parameters.Append Prm1
                rst2.CursorLocation = adUseClient
                rst2.CursorType = adOpenForwardOnly
                rst2.CacheSize = 100
                Set rst2 = cmd.Execute
                'CLMDental
                
                Do While Not rst2.EOF
                    dblMeternity = 0
                    dblDental = 0
                    OtherDed = 0
                    thyroid = 0
                    If IsNumeric(rst2!CLMMaternity) Then
                        dblMeternity = CDbl(rst2!CLMMaternity)
                        AccMaternity = AccMaternity + CDbl(rst2!CLMMaternity)
                    Else
                        dblMeternity = 0
                    End If
                      
                    If IsNumeric(rst2!CLMDental) Then
                        dblDental = CDbl(rst2!CLMDental)
                    Else
                        dblDental = 0
                    End If
                     
                    If dblDental > 0 Then
                        dblDental = dblDental
                    End If
                   
                    If IsNumeric(rst2!CLMOthersNoDed) Then
                        OtherDed = CDbl(rst2!CLMOthersNoDed)
                    Else
                        OtherDed = 0
                    End If
                      
                    If IsNumeric(rst2!CLMTyphiadJabNoDed) Then
                        thyroid = CDbl(rst2!CLMTyphiadJabNoDed)
                    Else
                        thyroid = 0
                    End If
    
                    If Len(Trim(rst2!CLMBordxAmt)) Then
                        If IsNumeric(rst2!CLMBordxAmt) Then
                            BordxAmt = CDbl(rst2!CLMBordxAmt)
                        Else
                            BordxAmt = 0
                        End If
                    Else
                        BordxAmt = 0
                    End If
    
                    tmpMBMFFS = CDbl(tmpMBMFFS) + IIf(IsNumeric(rst2!FFS), rst2!FFS, 0)
                    If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                        tmpMBMCLMAmt = CDbl(tmpMBMCLMAmt) + BordxAmt - (dblMeternity + OtherDed + thyroid + dblDental)
                    ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                        tmpMBMCLMAmt = CDbl(tmpMBMCLMAmt) + BordxAmt - (dblMeternity + OtherDed + thyroid)
                    ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                      tmpMBMCLMAmt = CDbl(tmpMBMCLMAmt) + BordxAmt - (OtherDed + thyroid + dblDental)
                    ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                      tmpMBMCLMAmt = CDbl(tmpMBMCLMAmt) + BordxAmt - (OtherDed + thyroid)
                    End If
                rst2.MoveNext
                Loop
                
                rst2.Close
                Set rst2 = Nothing
                Set cmd = Nothing
                Set Prm1 = Nothing
            End If
             
            lblMBMFFS = Format(tmpMBMFFS, "#,###.00")
            lblAccumAmt = Format(tmpMBMCLMAmt, "#,###.00")
            lblSpecAccumAmt = "N/A"
            lblSpecBalLimit = "N/A"
        Else
            If Len(txtMBMNumber) Then
                    sMBMNo = ""
                    SCovId = ""
                    If Len(txtMBMNumber) Then
                       If InStr(1, txtMBMNumber, "-") Then
                          sArr = Split(txtMBMNumber, "-")
                          sMBMNo = sArr(0)
                          SCovId = sArr(1)
                       Else
                          sMBMNo = txtMBMNumber
                          SCovId = "00"
                       End If
                    End If
                strAppend = ""
                Call tmpMaternitylimit(" CISCas.MBMNumber = '" & Trim(sMBMNo) & "'", tmpLimitStatus)
                If isPrinDepGPShared Then
                    strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(sMBMNo) & "' "
                Else
                    If tmpLimitStatus = "A" Then
                        strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(sMBMNo) & "'"
                    ElseIf tmpLimitStatus = "B" Then
                        If SCovId = "00" Then
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(sMBMNo) & "' AND CoverID = '00' "
                        Else
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(sMBMNo) & "' AND CoverID <> '00' "
                        End If
                    ElseIf tmpLimitStatus = "C" Or tmpLimitStatus = "" Then
                        strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(sMBMNo) & "' AND CoverID ='00' "
                    ElseIf tmpLimitStatus = "D" Then
                        If SCovId = "00" Then
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(sMBMNo) & "' AND CoverID = '00' "
                        Else
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(sMBMNo) & "' AND CoverID <> '00' "
                        End If
                    ElseIf tmpLimitStatus = "E" Then
                        strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(sMBMNo) & "' AND CoverID ='" + SCovId + "' "
                    End If
                End If
            End If
            
            If Trim(strAppend) = "" Then
                tmpMBMFFS = 0
                tmpMBMCLMAmt = 0
            Else
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "CIS_FFS2013"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd.Parameters.Append Prm1
                rst2.CursorLocation = adUseClient
                rst2.CursorType = adOpenForwardOnly
                rst2.CacheSize = 100
                Set rst2 = cmd.Execute

                SpecLmtRec.Open "SELECT Top 1 PLNMaternityInd,PDentalStatus FROM PLN WHERE PLNCode='" & Trim(!PLNType) & _
                "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                If SpecLmtRec.EOF = False Then
                    tmpMaternityInd = IIf(Len(SpecLmtRec!PLNMaternityInd), SpecLmtRec!PLNMaternityInd, "")
                    tmpDentalInd = IIf(Len(SpecLmtRec!PDentalStatus), SpecLmtRec!PDentalStatus, "")
                End If
                SpecLmtRec.Close
                Set SpecLmtRec = Nothing
                Do While Not rst2.EOF
                
                dblMeternity = 0
                dblDental = 0
                OtherDed = 0
                thyroid = 0
                BordxAmt = 0
                If Len(Trim(rst2!CLMBordxAmt)) Then
                    If IsNumeric(rst2!CLMBordxAmt) Then
                        BordxAmt = CDbl(rst2!CLMBordxAmt)
                    Else
                        BordxAmt = 0
                    End If
                Else
                    BordxAmt = 0
                End If

                If IsNumeric(rst2!CLMMaternity) Then
                    dblMeternity = CDbl(rst2!CLMMaternity)
                    AccMaternity = AccMaternity + CDbl(rst2!CLMMaternity)
                Else
                    dblMeternity = 0
                End If
                  
                If IsNumeric(rst2!CLMDental) Then
                    dblDental = CDbl(rst2!CLMDental)
                Else
                    dblDental = 0
                End If
                                
                If IsNumeric(rst2!CLMOthersNoDed) Then
                    OtherDed = CDbl(rst2!CLMOthersNoDed)
                Else
                    OtherDed = 0
                End If
                  
                If IsNumeric(rst2!CLMTyphiadJabNoDed) Then
                    thyroid = CDbl(rst2!CLMTyphiadJabNoDed)
                Else
                    thyroid = 0
                End If
                
                If rst2!PLNGPSPShared = "Y" Then
                    If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                        tmpMBMCLMAmt = CDbl(tmpMBMCLMAmt) + BordxAmt - (dblMeternity + OtherDed + thyroid + dblDental)
                    ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                        tmpMBMCLMAmt = CDbl(tmpMBMCLMAmt) + BordxAmt - (dblMeternity + OtherDed + thyroid)
                    ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                      tmpMBMCLMAmt = CDbl(tmpMBMCLMAmt) + BordxAmt - (OtherDed + thyroid + dblDental)
                    ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                      tmpMBMCLMAmt = CDbl(tmpMBMCLMAmt) + BordxAmt - (OtherDed + thyroid)
                    End If
                Else
                    If rst2!CLNType <> "S" Then
                            If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                                tmpMBMCLMAmt = CDbl(tmpMBMCLMAmt) + BordxAmt - (dblMeternity + OtherDed + thyroid + dblDental)
                            ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                                tmpMBMCLMAmt = CDbl(tmpMBMCLMAmt) + BordxAmt - (dblMeternity + OtherDed + thyroid)
                            ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                              tmpMBMCLMAmt = CDbl(tmpMBMCLMAmt) + BordxAmt - (OtherDed + thyroid + dblDental)
                            ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                              tmpMBMCLMAmt = CDbl(tmpMBMCLMAmt) + BordxAmt - (OtherDed + thyroid)
                            End If
                    End If
                    End If
                    rst2.MoveNext
                Loop
                rst2.Close
                Set rst2 = Nothing
                Set cmd = Nothing
                Set Prm1 = Nothing
            End If
            
            
            If Len(txtMBMNumber) Then
                    sMBMNo = ""
                    SCovId = ""
                    If Len(txtMBMNumber) Then
                       If InStr(1, txtMBMNumber, "-") Then
                          sArr = Split(txtMBMNumber, "-")
                          sMBMNo = sArr(0)
                          SCovId = sArr(1)
                       Else
                          sMBMNo = txtMBMNumber
                          SCovId = "00"
                       End If
                    End If
                strAppend = ""
                If tmpSpecLimitStatus = "A" Then
                    strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(sMBMNo) & "'"
                ElseIf tmpSpecLimitStatus = "B" Then
                    strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(sMBMNo) & "' AND CoverID ='00' "
                ElseIf tmpSpecLimitStatus = "C" Or tmpSpecLimitStatus = "" Then
                    strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(sMBMNo) & "' AND CoverID ='00' "
                ElseIf tmpSpecLimitStatus = "D" Then
                    strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(sMBMNo) & "' AND CoverID ='00' "
                ElseIf tmpSpecLimitStatus = "E" Then
                    strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(sMBMNo) & "' AND CoverID ='" + SCovId + "' "
                End If
            End If
            
            If Trim(strAppend) = "" Then
                tmpSpecMBMFFS = 0
                tmpSpecAccLmt = 0
            Else
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "CIS_FFS2013"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd.Parameters.Append Prm1
                rst2.CursorLocation = adUseClient
                rst2.CursorType = adOpenForwardOnly
                rst2.CacheSize = 100
                Set rst2 = cmd.Execute
        
                Do While Not rst2.EOF
                    If rst2!CLNType = "S" Then
                    
                        BordxAmt = 0
                        If Len(Trim(rst2!CLMBordxAmt)) Then
                            If IsNumeric(rst2!CLMBordxAmt) Then
                                BordxAmt = CDbl(rst2!CLMBordxAmt)
                            Else
                                BordxAmt = 0
                            End If
                        Else
                            BordxAmt = 0
                        End If
                        
                        tmpSpecMBMFFS = CDbl(tmpSpecMBMFFS) + IIf(IsNumeric(rst2!FFS), rst2!FFS, 0)
                        tmpSpecAccLmt = CDbl(tmpSpecAccLmt) + BordxAmt - CDbl(rst2!CLMOthersNoDed) - IIf(IsNumeric(rst2!CLMMaternity), rst2!CLMMaternity, 0)
                    End If
                    rst2.MoveNext
                Loop
                       
                rst2.Close
                Set rst2 = Nothing
                Set cmd = Nothing
                Set Prm1 = Nothing
            End If
            
            lblMBMFFS = Format(CDbl(tmpMBMFFS) + CDbl(tmpSpecMBMFFS), "#,##0.00")
            lblAccumAmt = Format(tmpMBMCLMAmt, "#,###.000")
            lblSpecAccumAmt = Format(tmpSpecAccLmt, "#,###.000")
            If IsNumeric(tmpSpecLmt) Then
                lblSpecAnnLmt = Format(lblSpecAnnLmt, "#,##0.00")
            End If
            If IsNumeric(lblSpecAnnLmt) Then
                lblSpecAccumAmt = Format(tmpSpecAccLmt, "#,###.000")
                lblSpecBalLimit = Format(CDbl(tmpSpecLmt) - CDbl(tmpSpecAccLmt), "#,##0.00")
            Else
                If Trim(lblSpecAnnLmt) = "FULL" Then
                    lblSpecAccumAmt = Format(tmpSpecAccLmt, "#,###.00")
                    lblSpecBalLimit = "FULL"
                ElseIf Trim(lblSpecAnnLmt) = "N/A" Then
                    lblSpecAccumAmt = "N/A"
                    lblSpecBalLimit = "N/A"
                Else
                    lblSpecAnnLmt = "N/A"
                    lblSpecAccumAmt = "N/A"
                    lblSpecBalLimit = "N/A"
                End If
            End If
        End If
        
        If SCovId = "00" And slblStatus = "E" Then
            lblAnnLmt = "FULL"
        Else
            lblAnnLmt = tmpAnnLmt
            If IsNumeric(tmpAnnLmt) Then
                lblAnnLmt = Format(lblAnnLmt, "#,##0.00")
            End If
            
            If IsNumeric(lblAnnLmt) Then
                    If IsNumeric(lblTopupAmt) Then
                        lblBalLimit = Format(CDbl(tmpAnnLmt) + CDbl(tmpTopupAmt) - CDbl(tmpMBMCLMAmt), "#,##0.00")
                    Else
                        lblBalLimit = Format(CDbl(tmpAnnLmt) - CDbl(tmpMBMCLMAmt), "#,##0.00")
                    End If
            Else
                If lblAnnLmt = "FULL" Then
                    lblBalLimit = "FULL"
                Else
                    lblBalLimit = "N/A"
                End If
            End If
        End If
        lblMBMVisitLimit = tmpVisitLmt
        If IsNumeric(tmpVisitLmt) Then
            lblMBMVisitLimit = Format(!MBMCLNVisitLmt, "#,##0.00")
        End If
        If IsNumeric(lblMeternityLmt) Then
           If IsNumeric(AccMaternity) Then
                lblMatAcc = AccMaternity
                lblMatBal = Format(CDbl(lblMeternityLmt) - CDbl(AccMaternity), "#,##0.00")
           End If
           If tmpLimitStatus = "C" Then
                lblMatAcc = tmpAccMaternity
                lblMatBal = Format(CDbl(lblMeternityLmt) - CDbl(tmpAccMaternity), "#,##0.00")
           End If
        Else
        lblMatAcc = ""
        lblMatBal = ""
        End If
        lblMEPLimit = Format(!MBMMEPLmt, "#,##0.00")
        lblDentalLimit = Format(!MBMDentalLmt, "#,##0.00")
        lblMBMSpecLimit = Format(!MBMSpecLmt, "#,##0.00")
        lblDentalExtLimit = Format(!MBMDentalExtLmt, "#,##0.00")
        lblFirstSpecLimit = Format(!MBMFirstSpecLmt, "#,##0.00")
        txtMBMPosition = IIf(Len(!MBMPosition), !MBMPosition, "")
    End With
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
   
End Sub

Private Sub showMBMCov()
Const METHOD_NAME = "showMBMCov"
On Error GoTo debugs


Dim TopupRec As New ADODB.Recordset
Dim tmpTopupAmt As Double
Dim MBMCoVRec As ListItem
Dim strMBMRef, tmpAnnLmt As String
Dim MBMRef As New ADODB.Recordset
Dim tmpVisitLmt As String
Dim PLnRec As New ADODB.Recordset
Dim StrPlnSql As String
Dim tmpSuppLmtStatus As String
Dim tmpSpecSuppLmtStatus As String
Dim tmpSuppFFS As Double
Dim strAppend As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim tmpMBMFFS As Double
Dim tmpSpecMBMFFS As Double
Dim tempSuppCovID As String
Dim rst2 As New ADODB.Recordset
Dim tmpMBMCCLMAmt As Double
Dim strAppend2 As String
Dim cmd2 As New ADODB.Command
Dim prm2 As New ADODB.Parameter
Dim rst3 As New ADODB.Recordset
Dim tmpSpecAccLmt As Double
Dim tmpSpecLmt As String
Dim isPrinDepGPShared As Boolean

tmpMBMFFS = 0
tmpSpecMBMFFS = 0
    lstCovAdd.ListItems.Clear
    tmpSpecLmt = ""
    Do While Not ADORSTMBM.EOF
        With ADORSTMBM
            Set MBMCoVRec = lstCovAdd.ListItems.Add(, , IIf(Trim(!MBMCName) <> "", !MBMCName, ""))
            MBMCoVRec.SubItems(1) = IIf(Len(Trim(!MBMCCoverID)), !MBMCCoverID, "")
            MBMCoVRec.SubItems(2) = IIf(Len(Trim(!MBMCICBCPPNo)), !MBMCICBCPPNo, "")
            MBMCoVRec.SubItems(3) = IIf(Len(Trim(!MBMCRELCode)), !MBMCRELCode, "")
            MBMCoVRec.SubItems(4) = IIf(Len(!MBMCCLNStatus), !MBMCCLNStatus, "")
            'MBMCoVRec.SubItems(5) = IIf(Len(Trim(!MBMCExclusion)), Trim(!MBMCExclusion), "")
            MBMCoVRec.SubItems(5) = IIf(Len(Trim(!MBMCCLNWarranty)), Trim(!MBMCCLNWarranty), "")
            MBMCoVRec.SubItems(6) = IIf(Len(Trim(!MBMCCLNRemarks)), !MBMCCLNRemarks, "")
            MBMCoVRec.SubItems(7) = IIf(Len(Trim(!MBMCDOB)), !MBMCDOB, "")
            If Trim(!MBMCCLNPLNCode) <> "" Then
                MBMCoVRec.SubItems(8) = IIf(Len(!MBMCCLNPLNCode), !MBMCCLNPLNCode, "")
            Else
                MBMCoVRec.SubItems(8) = IIf(Len(!PLNType), !PLNType, "")
            End If
            
            isPrinDepGPShared = !isPrinDepGPShared
            
            StrPlnSql = "Select SuppLimitStatus from AnnualLimit Where INSCode ='" & Trim(!INSType) & "' And PLNCode ='" & Trim(MBMCoVRec.SubItems(8)) & _
                 "' AND AnnEffDate <= '" & Format(!MBMPolEffDate, "mm/dd/yyyy") & "' order by AnnEffDate  desc"
            PLnRec.Open StrPlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            tmpSuppLmtStatus = ""
            Do While Not PLnRec.EOF
                tmpSuppLmtStatus = IIf(Len(Trim(PLnRec!SuppLimitStatus)), Trim(PLnRec!SuppLimitStatus), "")
                PLnRec.MoveNext
            Loop
            PLnRec.Close
            Set PLnRec = Nothing
            
            StrPlnSql = "Select SuppLimitStatus from SpecialistLimit Where INSCode ='" & Trim(!INSType) & "' And PLNCode ='" & Trim(MBMCoVRec.SubItems(8)) & _
                 "' AND EffDate <= '" & Format(!MBMPolEffDate, "mm/dd/yyyy") & "' order by EffDate  desc"
            PLnRec.Open StrPlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            tmpSpecSuppLmtStatus = ""
            Do While Not PLnRec.EOF
                tmpSpecSuppLmtStatus = IIf(Len(Trim(PLnRec!SuppLimitStatus)), Trim(PLnRec!SuppLimitStatus), "A")
                PLnRec.MoveNext
            Loop
            PLnRec.Close
            Set PLnRec = Nothing
            MBMCoVRec.SubItems(9) = IIf(Len(!MBMCCardType), !MBMCCardType, "")
            ' (9) IS FOR THE Amendment status
            MBMCoVRec.SubItems(10) = IIf(IsNumeric(!MBMCCLNPrem), !MBMCCLNPrem, 0)
            ' :- (11) - PremRefund, (12) - EffPrem
            MBMCoVRec.SubItems(13) = IIf(IsNumeric(!MBMCCLNMCO), !MBMCCLNMCO, 0)
            ' :- (14) - MCORefund, (15) - EffMCO
            
            If !ASOInd = "Y" Then
                If !MBMCCoverID <> "00" Then
                    DepID = IIf(Len(Trim(!MBMCCoverID)), !MBMCCoverID, "")
                    Call CallLimitDept("S", tmpAnnLmt, tmpVisitLmt, tmpSpecLmt, tmpLimitStatus)
                Else
                    Call CallLimit("S", tmpAnnLmt, tmpVisitLmt, tmpSpecLmt, tmpLimitStatus)
                End If
            ElseIf (tmpSuppLmtStatus = "F") Then
                
                Dim PLnRec1 As New ADODB.Recordset
                Dim PlnSql As String
                PlnSql = "SELECT ANNLmtIndSupp, ChildLimitInd, SuppLimit, ChildDepLimit,SSpecLimit,VISLmtInd from cisdb..PLN inner join cisdb..AnnualLimit on AnnualLimit.PLNCode=pln.PLNCode WHERE pln.PLNCode='" & Trim(!PLNType) & "'"
                
                PLnRec1.Open PlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                
                
                
                If PLnRec1.EOF = False Then
                    If Trim(!MBMCRELCode) = "W" Or Trim(!MBMCRELCode) = "H" Then
                        If Trim(PLnRec1!ANNLmtIndSupp) = "Y" Then
                           tmpAnnLmt = PLnRec1!SuppLimit
                        Else
                           tmpAnnLmt = "FULL"
                        End If
                    ElseIf Trim(!MBMCRELCode) = "S" Or Trim(!MBMCRELCode) = "D" Then
                        If Trim(PLnRec1!ANNLmtIndSupp) = "Y" Then
                           tmpAnnLmt = PLnRec1!ChildDepLimit
                        Else
                           tmpAnnLmt = "FULL"
                        End If
                    End If
                    tmpSpecLmt = PLnRec1!SSpecLimit
                    tmpVisitLmt = PLnRec1!VISLmtInd
                End If
                PLnRec1.Close
            Else
                Call CallLimit("S", tmpAnnLmt, tmpVisitLmt, tmpSpecLmt, tmpLimitStatus)
            End If
            
            
            If Trim(tmpSpecLmt) = "N/A" Then
                If Len(txtMBMNumber) Then
                    strAppend2 = ""
                    If tmpSuppLmtStatus = "B" And MBMCoVRec.SubItems(1) = "01" Then
                        strAppend2 = strAppend2 & "  CISCas.MBMnumber= '" & Trim(txtMBMNumber) & "' AND CoverID <> '00' "
                    ElseIf (tmpSuppLmtStatus = "C" Or tmpSuppLmtStatus = "F") Or tmpSuppLmtStatus = "" Then
                        strAppend2 = strAppend2 & "  CISCas.MBMnumber= '" & Trim(txtMBMNumber) & "' AND CoverID = '" & MBMCoVRec.SubItems(1) & "'"
                    ElseIf tmpSuppLmtStatus = "D" And MBMCoVRec.SubItems(1) = "01" Then
                        strAppend2 = strAppend2 & "  CISCas.MBMnumber= '" & Trim(txtMBMNumber) & "' And CoverID  <> '00' "
                    ElseIf tmpSuppLmtStatus = "E" Then
                        strAppend2 = strAppend2 & "  CISCas.MBMnumber= '" & Trim(txtMBMNumber) & "' And CoverID  = '" & MBMCoVRec.SubItems(1) & "'"
                    
                    End If
                End If
                
                If tmpSuppLmtStatus <> "A" Then
                    tmpMBMCCLMAmt = 0
                    tmpMBMFFS = 0
                    tmpSpecAccLmt = 0
                    tmpSpecMBMFFS = 0
                    If Trim(strAppend2) = "" Then
                        tmpMBMCCLMAmt = 0
                        tmpMBMFFS = 0
                    Else
                        Set cmd2.ActiveConnection = MediConnCISDB
                        cmd2.CommandText = "CIS_FFS2013"
                        cmd2.CommandType = adCmdStoredProc
                        cmd2.CommandTimeout = 300
                        Set prm2 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend2)
                        cmd2.Parameters.Append prm2
                        rst3.CursorLocation = adUseClient
                        rst3.CursorType = adOpenForwardOnly
                        rst3.CacheSize = 100
                        Set rst3 = cmd2.Execute
                        
                        Do While Not rst3.EOF
                            If tmpSuppLmtStatus = "B" And MBMCoVRec.SubItems(1) = "01" Then
                                tmpMBMFFS = CDbl(tmpMBMFFS) + IIf(IsNumeric(rst3!FFS), rst3!FFS, 0)
                                tmpMBMCCLMAmt = CDbl(tmpMBMCCLMAmt) + IIf(IsNumeric(rst3!CLMBordxAmt), rst3!CLMBordxAmt, 0) - IIf(IsNumeric(rst3!CLMOthersNoDed), rst3!CLMOthersNoDed, 0) - IIf(IsNumeric(rst3!CLMMaternity), rst3!CLMMaternity, 0)
                            'murali
                            ElseIf tmpSuppLmtStatus = "C" Or tmpSuppLmtStatus = "" Then
                                tmpMBMFFS = CDbl(tmpMBMFFS) + IIf(IsNumeric(rst3!FFS), rst3!FFS, 0)
                                tmpMBMCCLMAmt = CDbl(tmpMBMCCLMAmt) + IIf(IsNumeric(rst3!CLMBordxAmt), rst3!CLMBordxAmt, 0) - IIf(IsNumeric(rst3!CLMOthersNoDed), rst3!CLMOthersNoDed, 0) - IIf(IsNumeric(rst3!CLMMaternity), rst3!CLMMaternity, 0)
                            ElseIf tmpSuppLmtStatus = "D" And MBMCoVRec.SubItems(1) = "01" Then
                                tmpMBMFFS = CDbl(tmpMBMFFS) + IIf(IsNumeric(rst3!FFS), rst3!FFS, 0)
                                tmpMBMCCLMAmt = CDbl(tmpMBMCCLMAmt) + IIf(IsNumeric(rst3!CLMBordxAmt), rst3!CLMBordxAmt, 0) - IIf(IsNumeric(rst3!CLMOthersNoDed), rst3!CLMOthersNoDed, 0) - IIf(IsNumeric(rst3!CLMMaternity), rst3!CLMMaternity, 0)
                            ElseIf tmpSuppLmtStatus = "E" Then
                                tmpMBMFFS = CDbl(tmpMBMFFS) + IIf(IsNumeric(rst3!FFS), rst3!FFS, 0)
                                tmpMBMCCLMAmt = CDbl(tmpMBMCCLMAmt) + IIf(IsNumeric(rst3!CLMBordxAmt), rst3!CLMBordxAmt, 0) - IIf(IsNumeric(rst3!CLMOthersNoDed), rst3!CLMOthersNoDed, 0) - IIf(IsNumeric(rst3!CLMMaternity), rst3!CLMMaternity, 0)
                            End If
                            
                            
                            rst3.MoveNext
                        Loop
                        rst3.Close
                        Set rst3 = Nothing
                        Set cmd2 = Nothing
                        Set prm2 = Nothing
                    End If
                    tmpSpecAccLmt = 0
                    tmpSpecMBMFFS = 0
                Else
                    tmpMBMCCLMAmt = 0
                    tmpMBMFFS = 0
                    tmpSpecAccLmt = 0
                    tmpSpecMBMFFS = 0
                End If
                MBMCoVRec.SubItems(27) = "N/A"
                MBMCoVRec.SubItems(28) = "N/A"
                MBMCoVRec.SubItems(29) = "N/A"
            Else
            
                If Len(txtMBMNumber) Then
                    strAppend2 = ""
                    If tmpSuppLmtStatus = "B" And MBMCoVRec.SubItems(1) = "01" Then
                        strAppend2 = strAppend2 & "  CISCas.MBMnumber= '" & Trim(txtMBMNumber) & "' AND CoverID <> '00' "
                    ElseIf (tmpSuppLmtStatus = "C" Or tmpSuppLmtStatus = "F") Or tmpSuppLmtStatus = "" Then
                        strAppend2 = strAppend2 & "  CISCas.MBMnumber= '" & Trim(txtMBMNumber) & "' AND CoverID = '" & MBMCoVRec.SubItems(1) & "'"
                    ElseIf tmpSuppLmtStatus = "D" And MBMCoVRec.SubItems(1) = "01" Then
                        strAppend2 = strAppend2 & "  CISCas.MBMnumber= '" & Trim(txtMBMNumber) & "' And CoverID  <> '00' "
                    ElseIf tmpSuppLmtStatus = "E" Then
                        strAppend2 = strAppend2 & "  CISCas.MBMnumber= '" & Trim(txtMBMNumber) & "' AND CoverID = '" & MBMCoVRec.SubItems(1) & "'"
                    End If
                End If
            
                If tmpSuppLmtStatus <> "A" Then
                    tmpMBMCCLMAmt = 0
                    tmpMBMFFS = 0
                    tmpSpecAccLmt = 0
                    tmpSpecMBMFFS = 0
                    
                    If Trim(strAppend2) = "" Then
                        tmpMBMCCLMAmt = 0
                        tmpMBMFFS = 0
                    Else
                        Set cmd2.ActiveConnection = MediConnCISDB
                        cmd2.CommandText = "CIS_FFS2013"
                        cmd2.CommandType = adCmdStoredProc
                        cmd2.CommandTimeout = 300
                        Set prm2 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend2)
                        cmd2.Parameters.Append prm2
                        rst3.CursorLocation = adUseClient
                        rst3.CursorType = adOpenForwardOnly
                        rst3.CacheSize = 100
                        Set rst3 = cmd2.Execute
              
                        Do While Not rst3.EOF
                            If tmpSuppLmtStatus = "B" And MBMCoVRec.SubItems(1) = "01" Then
                                If rst3!CLNType <> "S" Then
                                    tmpMBMFFS = CDbl(tmpMBMFFS) + IIf(IsNumeric(rst3!FFS), rst3!FFS, 0)
                                    tmpMBMCCLMAmt = CDbl(tmpMBMCCLMAmt) + IIf(IsNumeric(rst3!CLMBordxAmt), rst3!CLMBordxAmt, 0) - IIf(IsNumeric(rst3!CLMOthersNoDed), rst3!CLMOthersNoDed, 0)
                                End If
                            ElseIf tmpSuppLmtStatus = "C" Or tmpSuppLmtStatus = "F" Or tmpSuppLmtStatus = "" Then
                                If rst3!CLNType <> "S" Then
                                    tmpMBMFFS = CDbl(tmpMBMFFS) + IIf(IsNumeric(rst3!FFS), rst3!FFS, 0)
                                    tmpMBMCCLMAmt = CDbl(tmpMBMCCLMAmt) + IIf(IsNumeric(rst3!CLMBordxAmt), rst3!CLMBordxAmt, 0) - IIf(IsNumeric(rst3!CLMOthersNoDed), rst3!CLMOthersNoDed, 0)
                                    
                                End If
                            ElseIf tmpSuppLmtStatus = "D" And MBMCoVRec.SubItems(1) = "01" Then
                                If rst3!CLNType <> "S" Then
                                    tmpMBMFFS = CDbl(tmpMBMFFS) + IIf(IsNumeric(rst3!FFS), rst3!FFS, 0)
                                    tmpMBMCCLMAmt = CDbl(tmpMBMCCLMAmt) + IIf(IsNumeric(rst3!CLMBordxAmt), rst3!CLMBordxAmt, 0) - IIf(IsNumeric(rst3!CLMOthersNoDed), rst3!CLMOthersNoDed, 0)
                                End If
                            ElseIf tmpSuppLmtStatus = "E" Then
                                If rst3!CLNType <> "S" Then
                                    tmpMBMFFS = CDbl(tmpMBMFFS) + IIf(IsNumeric(rst3!FFS), rst3!FFS, 0)
                                    tmpMBMCCLMAmt = CDbl(tmpMBMCCLMAmt) + IIf(IsNumeric(rst3!CLMBordxAmt), rst3!CLMBordxAmt, 0) - IIf(IsNumeric(rst3!CLMOthersNoDed), rst3!CLMOthersNoDed, 0)
                                End If
                            End If
                            rst3.MoveNext
                        Loop
                    
                        rst3.Close
                        Set rst3 = Nothing
                        Set cmd2 = Nothing
                        Set prm2 = Nothing
                    End If
                
                Else
                    tmpMBMCCLMAmt = 0
                    tmpMBMFFS = 0
                End If
                
                
                If Len(txtMBMNumber) Then
                    strAppend2 = ""
                    If tmpSpecSuppLmtStatus = "B" And MBMCoVRec.SubItems(1) = "01" Then
                        strAppend2 = strAppend2 & "  CISCas.MBMnumber= '" & Trim(txtMBMNumber) & "' AND CoverID <> '00' "
                    ElseIf (tmpSpecSuppLmtStatus = "C" Or tmpSpecSuppLmtStatus = "F") Or tmpSpecSuppLmtStatus = "" Then
                        strAppend2 = strAppend2 & "  CISCas.MBMnumber= '" & Trim(txtMBMNumber) & "' AND CoverID = '" & MBMCoVRec.SubItems(1) & "'"
                    ElseIf tmpSpecSuppLmtStatus = "D" And MBMCoVRec.SubItems(1) = "01" Then
                        strAppend2 = strAppend2 & "  CISCas.MBMnumber= '" & Trim(txtMBMNumber) & "' And CoverID  <> '00' "
                    ElseIf tmpSpecSuppLmtStatus = "E" Then
                        strAppend2 = strAppend2 & "  CISCas.MBMnumber= '" & Trim(txtMBMNumber) & "' AND CoverID = '" & MBMCoVRec.SubItems(1) & "'"
                    End If
                End If
                
                If tmpSpecSuppLmtStatus <> "A" Then
                    tmpSpecMBMFFS = 0
                    tmpSpecAccLmt = 0
                    
                    If Trim(strAppend2) = "" Then
                        tmpSpecMBMFFS = 0
                        tmpSpecAccLmt = 0
                    Else
                        Set cmd2.ActiveConnection = MediConnCISDB
                        cmd2.CommandText = "CIS_FFS2013"
                        cmd2.CommandType = adCmdStoredProc
                        cmd2.CommandTimeout = 300
                        Set prm2 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend2)
                        cmd2.Parameters.Append prm2
                        rst3.CursorLocation = adUseClient
                        rst3.CursorType = adOpenForwardOnly
                        rst3.CacheSize = 100
                        Set rst3 = cmd2.Execute
        
             
                        Do While Not rst3.EOF
                            If tmpSpecSuppLmtStatus = "B" And MBMCoVRec.SubItems(1) = "01" Then
                                If rst3!CLNType = "S" Then
                                    tmpSpecMBMFFS = CDbl(tmpSpecMBMFFS) + IIf(IsNumeric(rst3!FFS), rst3!FFS, 0)
                                    tmpSpecAccLmt = CDbl(tmpSpecAccLmt) + IIf(IsNumeric(rst3!CLMBordxAmt), rst3!CLMBordxAmt, 0) - IIf(IsNumeric(rst3!CLMOthersNoDed), rst3!CLMOthersNoDed, 0) - IIf(IsNumeric(rst3!CLMMaternity), rst3!CLMMaternity, 0)
                                End If
                            ElseIf tmpSpecSuppLmtStatus = "C" Or tmpSuppLmtStatus = "" Then
                                If rst3!CLNType = "S" Then
                                    tmpSpecMBMFFS = CDbl(tmpSpecMBMFFS) + IIf(IsNumeric(rst3!FFS), rst3!FFS, 0)
                                    tmpSpecAccLmt = CDbl(tmpSpecAccLmt) + IIf(IsNumeric(rst3!CLMBordxAmt), rst3!CLMBordxAmt, 0) - IIf(IsNumeric(rst3!CLMOthersNoDed), rst3!CLMOthersNoDed, 0) - IIf(IsNumeric(rst3!CLMMaternity), rst3!CLMMaternity, 0)
                                End If
                            ElseIf tmpSpecSuppLmtStatus = "D" And MBMCoVRec.SubItems(1) = "01" Then
                                If rst3!CLNType = "S" Then
                                    tmpSpecMBMFFS = CDbl(tmpSpecMBMFFS) + IIf(IsNumeric(rst3!FFS), rst3!FFS, 0)
                                    tmpSpecAccLmt = CDbl(tmpSpecAccLmt) + IIf(IsNumeric(rst3!CLMBordxAmt), rst3!CLMBordxAmt, 0) - IIf(IsNumeric(rst3!CLMOthersNoDed), rst3!CLMOthersNoDed, 0) - IIf(IsNumeric(rst3!CLMMaternity), rst3!CLMMaternity, 0)
                                End If
                            End If
                            rst3.MoveNext
                        Loop

                        rst3.Close
                        Set rst3 = Nothing
                        Set cmd2 = Nothing
                        Set prm2 = Nothing
                    End If
                
                Else
                    tmpSpecAccLmt = 0
                    tmpSpecMBMFFS = 0
                End If
                
                If IsNumeric(tmpSpecLmt) Then
                    If tmpSpecSuppLmtStatus <> "A" Then
                        If tmpSpecSuppLmtStatus = "B" And MBMCoVRec.SubItems(1) = "01" Then
                            MBMCoVRec.SubItems(27) = tmpSpecLmt
                            MBMCoVRec.SubItems(28) = IIf(IsNumeric(tmpSpecAccLmt), tmpSpecAccLmt, 0)
                            MBMCoVRec.SubItems(28) = Format(tmpSpecAccLmt, "#,##0.00")
                            MBMCoVRec.SubItems(29) = Format(CDbl(tmpSpecLmt) - CDbl(tmpSpecAccLmt), "#,##0.00")
                        ElseIf tmpSpecSuppLmtStatus = "C" Or tmpSpecSuppLmtStatus = "" Then
                            MBMCoVRec.SubItems(27) = tmpSpecLmt
                            MBMCoVRec.SubItems(28) = IIf(IsNumeric(tmpSpecAccLmt), tmpSpecAccLmt, 0)
                            MBMCoVRec.SubItems(28) = Format(tmpSpecAccLmt, "#,##0.00")
                            MBMCoVRec.SubItems(29) = Format(CDbl(tmpSpecLmt) - CDbl(tmpSpecAccLmt), "#,##0.00")
                        ElseIf tmpSpecSuppLmtStatus = "D" And MBMCoVRec.SubItems(1) = "01" Then
                            MBMCoVRec.SubItems(27) = tmpSpecLmt
                            MBMCoVRec.SubItems(28) = IIf(IsNumeric(tmpSpecAccLmt), tmpSpecAccLmt, 0)
                            MBMCoVRec.SubItems(28) = Format(tmpSpecAccLmt, "#,##0.00")
                            MBMCoVRec.SubItems(29) = Format(CDbl(tmpSpecLmt) - CDbl(tmpSpecAccLmt), "#,##0.00")
                        Else
                            MBMCoVRec.SubItems(27) = 0
                            MBMCoVRec.SubItems(28) = 0
                            MBMCoVRec.SubItems(29) = 0
                        End If
                    Else
                        MBMCoVRec.SubItems(27) = 0
                        MBMCoVRec.SubItems(28) = 0
                        MBMCoVRec.SubItems(29) = 0
                    End If
                Else
                    If Trim(tmpSpecLmt) = "FULL" Then
                        MBMCoVRec.SubItems(27) = tmpSpecLmt
                        MBMCoVRec.SubItems(28) = IIf(IsNumeric(tmpSpecAccLmt), tmpSpecAccLmt, 0)
                        MBMCoVRec.SubItems(28) = Format(tmpSpecAccLmt, "#,##0.00")
                        MBMCoVRec.SubItems(29) = "FULL"
                    Else
                        MBMCoVRec.SubItems(27) = 0
                        MBMCoVRec.SubItems(28) = 0
                        MBMCoVRec.SubItems(29) = 0
                    End If
                End If
            End If
            
            
            If tmpSuppLmtStatus <> "A" And isPrinDepGPShared = False Then
                If tmpSuppLmtStatus = "B" And MBMCoVRec.SubItems(1) = "01" Then
                    MBMCoVRec.SubItems(17) = IIf(IsNumeric(tmpMBMCCLMAmt), tmpMBMCCLMAmt, 0)
                    MBMCoVRec.SubItems(17) = Format(tmpMBMCCLMAmt, "#,##0.00")
                    'MBMCoVRec.SubItems(28) = IIf(IsNumeric(tmpSpecAccLmt), tmpSpecAccLmt, 0)
                    'MBMCoVRec.SubItems(28) = Format(tmpSpecAccLmt, "#,##0.00")
                ElseIf (tmpSuppLmtStatus = "C" Or tmpSuppLmtStatus = "F") Or tmpSuppLmtStatus = "" Then
                    MBMCoVRec.SubItems(17) = IIf(IsNumeric(tmpMBMCCLMAmt), tmpMBMCCLMAmt, 0)
                    MBMCoVRec.SubItems(17) = Format(tmpMBMCCLMAmt, "#,##0.00")
                    'MBMCoVRec.SubItems(28) = IIf(IsNumeric(tmpSpecAccLmt), tmpSpecAccLmt, 0)
                    'MBMCoVRec.SubItems(28) = Format(tmpSpecAccLmt, "#,##0.00")
                ElseIf tmpSuppLmtStatus = "D" And MBMCoVRec.SubItems(1) = "01" Then
                    MBMCoVRec.SubItems(17) = IIf(IsNumeric(tmpMBMCCLMAmt), tmpMBMCCLMAmt, 0)
                    MBMCoVRec.SubItems(17) = Format(tmpMBMCCLMAmt, "#,##0.00")
                    'MBMCoVRec.SubItems(28) = IIf(IsNumeric(tmpSpecAccLmt), tmpSpecAccLmt, 0)
                    'MBMCoVRec.SubItems(28) = Format(tmpSpecAccLmt, "#,##0.00")
                ElseIf tmpSuppLmtStatus = "E" Then
                    MBMCoVRec.SubItems(17) = IIf(IsNumeric(tmpMBMCCLMAmt), tmpMBMCCLMAmt, 0)
                    MBMCoVRec.SubItems(17) = Format(tmpMBMCCLMAmt, "#,##0.00")
                    'MBMCoVRec.SubItems(28) = IIf(IsNumeric(tmpSpecAccLmt), tmpSpecAccLmt, 0)
                    'MBMCoVRec.SubItems(28) = Format(tmpSpecAccLmt, "#,##0.00")
                Else
                    MBMCoVRec.SubItems(17) = 0
                    'MBMCoVRec.SubItems(28) = 0
                End If
            Else
                MBMCoVRec.SubItems(17) = 0
            End If
            
            
            If IsNumeric(tmpAnnLmt) Then
                'MBMCoVRec.SubItems(18) = Format(!MBMCLNBalLmt, "#,##0.00")
                TopupRec.Open " SELECT SUM(TopupAmt) AS sumOfTopupAmt FROM MBMCISTopupHistory WHERE (MBMNumber = '" & txtMBMNumber & "') AND (MBMCovID = '01') GROUP BY MBMNumber ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                With TopupRec
                    If TopupRec.EOF = False Then
                        tmpTopupAmt = IIf(IsNumeric(TopupRec!sumOfTopupAmt), TopupRec!sumOfTopupAmt, "")
                        lblCovTopUp = Format(tmpTopupAmt, "#,##0.00")
                    Else
                        tmpTopupAmt = 0
                        lblCovTopUp = Format("0", "#,##0.00")
                    End If
                End With
                TopupRec.Close
                Set TopupRec = Nothing
                If tmpSuppLmtStatus <> "A" And isPrinDepGPShared = False Then
                    If tmpSuppLmtStatus = "B" And MBMCoVRec.SubItems(1) = "01" Then
                        MBMCoVRec.SubItems(16) = tmpAnnLmt
                        'MBMCoVRec.SubItems(18) = tmpAnnLmt
                        MBMCoVRec.SubItems(18) = Format(CDbl(tmpAnnLmt) - CDbl(tmpMBMCCLMAmt) + CDbl(tmpTopupAmt), "#,##0.00")
                        sharelmtstatus = True
                    ElseIf (tmpSuppLmtStatus = "C" Or tmpSuppLmtStatus = "F") Or tmpSuppLmtStatus = "" Then
                        MBMCoVRec.SubItems(16) = tmpAnnLmt
                        'MBMCoVRec.SubItems(18) = tmpAnnLmt
                        MBMCoVRec.SubItems(18) = Format(CDbl(tmpAnnLmt) - CDbl(tmpMBMCCLMAmt) + CDbl(tmpTopupAmt), "#,##0.00")
                    ElseIf tmpSuppLmtStatus = "D" And MBMCoVRec.SubItems(1) = "01" Then
                    
                        MBMCoVRec.SubItems(16) = tmpAnnLmt
                        'MBMCoVRec.SubItems(18) = tmpAnnLmt
                        MBMCoVRec.SubItems(18) = Format(CDbl(tmpAnnLmt) - CDbl(tmpMBMCCLMAmt) + CDbl(tmpTopupAmt), "#,##0.00")
                    ElseIf tmpSuppLmtStatus = "E" Then
                        MBMCoVRec.SubItems(16) = tmpAnnLmt
                        'MBMCoVRec.SubItems(18) = tmpAnnLmt
                        MBMCoVRec.SubItems(18) = Format(CDbl(tmpAnnLmt) - CDbl(tmpMBMCCLMAmt) + CDbl(tmpTopupAmt), "#,##0.00")
                    ElseIf tmpSuppLmtStatus = "B" And MBMCoVRec.SubItems(1) = "02" And sharelmtstatus = False Then
                        MBMCoVRec.SubItems(16) = tmpAnnLmt
                        MBMCoVRec.SubItems(18) = Format(CDbl(tmpAnnLmt) - CDbl(tmpMBMCCLMAmt) + CDbl(tmpTopupAmt), "#,##0.00")
                        sharelmtstatus = True
                    Else
                        MBMCoVRec.SubItems(16) = 0
                        'MBMCoVRec.SubItems(18) = 0
                        MBMCoVRec.SubItems(18) = 0
                    End If
                Else
                    MBMCoVRec.SubItems(16) = 0
                    'MBMCoVRec.SubItems(18) = 0
                    MBMCoVRec.SubItems(18) = 0
                End If
            Else
                If Trim(tmpAnnLmt) = "FULL" Then
                    MBMCoVRec.SubItems(16) = "FULL"
                    MBMCoVRec.SubItems(18) = "FULL"
                Else
                    MBMCoVRec.SubItems(16) = 0
                    'MBMCoVRec.SubItems(18) = 0
                    MBMCoVRec.SubItems(18) = 0
                End If
            End If
            
             
            MBMCoVRec.SubItems(19) = tmpVisitLmt
            
            If IsNumeric(tmpVisitLmt) Then
                MBMCoVRec.SubItems(19) = Format(tmpVisitLmt, "#,##0.00")
            End If
            
            MBMCoVRec.SubItems(20) = IIf(IsNumeric(!MBMCCLNMEPLmt), !MBMCCLNMEPLmt, 0)
            MBMCoVRec.SubItems(21) = IIf(IsNumeric(!MBMCCLNDentalLmt), !MBMCCLNDentalLmt, 0)
            MBMCoVRec.SubItems(22) = IIf(IsNumeric(!MBMCCLNDentalExtLmt), !MBMCCLNDentalExtLmt, 0)
            MBMCoVRec.SubItems(23) = IIf(IsNumeric(!MBMCCLNSpecLmt), !MBMCCLNSpecLmt, 0)
            MBMCoVRec.SubItems(24) = IIf(IsNumeric(!MBMCCLNFirstSpecLmt), !MBMCCLNFirstSpecLmt, 0)
            
            If ADORSTMBM!MBMCLNStatus = "C" Then
                strMBMRef = "Select MBMPremReduction,MBMMCORefund,MBMMCOEffective, MBMPREMEffective " & _
                               " from MBMCISRefund where MBMNumber =  '" & Trim(ADORSTMBM!MBMNumber) & _
                               "' AND MBMCoveredID = '" & Trim(!MBMCCoverID) & "'"
                MBMRef.Open strMBMRef, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
                 Do While Not MBMRef.EOF
                    With MBMRef
                        MBMCoVRec.SubItems(11) = Format(!MBMPremReduction, "#,##0.00")
                        MBMCoVRec.SubItems(12) = Format(!MBMPREMEffective, "#,##0.00")
                        
                        MBMCoVRec.SubItems(14) = Format(!MBMMCORefund, "#,##0.00")
                        MBMCoVRec.SubItems(15) = Format(!MBMMCOEffective, "#,##0.00")
                    End With
                    MBMRef.MoveNext
                Loop
                MBMRef.Close
                Set MBMRef = Nothing
            End If
            MBMCoVRec.SubItems(26) = "A"
            MBMCoVRec.SubItems(25) = MemberCallAge(txtMBMPayorEffDate, MBMCoVRec.SubItems(7))
            
            If Trim(MBMCoVRec.SubItems(25)) = "01" Then
                MBMCoVRec.SubItems(26) = "C"
            End If

            
            If !MBMCEffDate <> "" Then
                MBMCoVRec.SubItems(30) = Format(!MBMCEffDate, "dd/mm/yyyy")
            Else
                MBMCoVRec.SubItems(30) = Format(txtMBMPayorEffDate, "dd/mm/yyyy")
            End If
            If !MBMCEffDate <> "" Then
                MBMCoVRec.SubItems(31) = Format(!MBMCExpDate, "dd/mm/yyyy")
            Else
                MBMCoVRec.SubItems(31) = Format(txtMBMPayorExpDate, "dd/mm/yyyy")
            End If
            
            If isPrinDepGPShared Then
                MBMCoVRec.SubItems(16) = "N/A"
                MBMCoVRec.SubItems(17) = "N/A"
                MBMCoVRec.SubItems(18) = "N/A"
            End If
        
            ' (25) - Ageband, (26) -Adult/Child
        End With
        ADORSTMBM.MoveNext
    Loop
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
   
End Sub

Private Function MemberCallAge(TmpEffDate, tmpDOB As String) As String
Dim SelectAge As String
Dim tmpAgeband As String

    If Trim(TmpEffDate) <> "__/__/____" And Trim(tmpDOB) <> "__/__/____" Then
        If IsDate(tmpDOB) Then
            SelectAge = GetAge(CDate(TmpEffDate), CDate(tmpDOB))
            tmpAgeband = GetAgeBand(SelectAge)
        End If
    End If
    MemberCallAge = Format(IIf(Len(tmpAgeband), Mid(tmpAgeband, 1, 2), 0), "00")
End Function

Private Sub MemberCallFees(tmpMBMTYpe As String, tmpAgeCode As String, tmpAdultChild As String, tempPremium, tempMCOFee As Single)
    tempPremium = Format(GetPremium(tmpMBMTYpe, Mid(cboINSCode, 1, 1), Trim(txtPAYCode), Trim(Mid(tmpAgeCode, 1, 2)), Trim(Mid(cboPLNCode, 1, 4)), "N", txtMBMPayorEffDate), "#,##0.00")
    
    tempMCOFee = Format(GetMCO(tmpMBMTYpe, Mid(cboINSCode, 1, 1), "N", tmpAdultChild, txtMBMPayorEffDate, Trim(Mid(cboPLNCode, 1, 4)), Trim(txtPAYCode)), "#,##0.00")
End Sub

Private Sub tmpMaternitylimit(strAppend As String, Maternitystauts As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim dblMeternity As Double


Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "CIS_FFS2013"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd.Parameters.Append Prm1
                rst2.CursorLocation = adUseClient
                rst2.CursorType = adOpenForwardOnly
                rst2.CacheSize = 100
                Set rst2 = cmd.Execute
        
        Do While Not rst2.EOF
        dblMeternity = 0
            If IsNumeric(rst2!CLMMaternity) Then
                dblMeternity = CDbl(rst2!CLMMaternity)
                tmpAccMaternity = tmpAccMaternity + CDbl(rst2!CLMMaternity)
            Else
                dblMeternity = 0
            End If
        rst2.MoveNext
        Loop
rst2.Close
Set rst2 = Nothing
Set cmd = Nothing
Set Prm1 = Nothing
    
End Sub




Private Sub CallLimit(tmpMBMTYpe As String, tmpAnnLmt As String, tmpVisitLmt As String, tmpSpecLmt As String, tmpLimitStatus As String)
Const METHOD_NAME = "CallLimit"
On Error GoTo debugs

Dim ASORS As New ADODB.Recordset
Dim ASOqry As String
Dim IsCombined As String

Dim ASOlimit As Double
Dim IPLimit As Double
Dim OPLimit As Double
Dim IPUtilization As Double
Dim OPUtilization As Double

Dim sMBMNumber() As String
Dim MBMNumber As String
Dim coverid As String

    ASOlimit = 0
    IPLimit = 0
    OPLimit = 0
    IsCombined = ""
    
    If ASOInd = "Y" Then
    
        If (InStr(1, txtMBMNumber, "-")) = 1 Then
            sMBMNumber = Split(txtMBMNumber, "-")
            MBMNumber = sMBMNumber(0)
            coverid = sMBMNumber(1)
        Else
            MBMNumber = txtMBMNumber
            coverid = "00"
        End If
        ASOqry = ""
        ASOqry = "Select MBMASOLimit,MBMIPLimit,MBMOPLimit,MBMIPUtilized,MBMOPUtilized ,IsIPOPCombined  from Tbl_ASOMemberDetails  where MBMNumber= '" + MBMNumber + "' and MBMCoverID='" + coverid + "'"
        ASORS.Open ASOqry, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do Until ASORS.EOF
             ASOlimit = IIf(Len(ASORS!MBMASOLimit), ASORS!MBMASOLimit, 0)
             IPLimit = IIf(Len(ASORS!MBMIPLimit), ASORS!MBMIPLimit, 0)
             OPLimit = IIf(Len(ASORS!MBMOPLimit), ASORS!MBMOPLimit, 0)
             IPUtilization = IIf(Len(ASORS!MBMIPUtilized), ASORS!MBMIPUtilized, 0)
             OPUtilization = IIf(Len(ASORS!MBMOPUtilized), ASORS!MBMOPUtilized, 0)
             IsCombined = IIf(Len(ASORS!IsIPOPCombined), ASORS!IsIPOPCombined, "N")
        ASORS.MoveNext
        Loop
        ASORS.Close
        Set ASORS = Nothing
        
        If IsCombined = "Y" Then
            tmpAnnLmt = ASOlimit '- (IPUtilization + OPUtilization)
        Else
            tmpAnnLmt = OPLimit '- OPUtilization
        End If
    Else
        tmpAnnLmt = ""
        tmpAnnLmt = Format(GetAnnualLimit(cboINSCode, cboPLNCode, "N", txtMBMPayorEffDate, tmpMBMTYpe, Trim(txtPAYCode), tmpLimitStatus), "#,##0.00")
    End If
    
    tmpVisitLmt = ""
    tmpVisitLmt = Format(GetVisitLimit(cboINSCode, cboPLNCode, "N", txtMBMPayorEffDate, tmpMBMTYpe, Trim(txtPAYCode)), "#,##0.00")
   
    tmpSpecLmt = ""
    tmpSpecLmt = Format(GetSpecLimit(cboINSCode, cboPLNCode, "N", txtMBMPayorEffDate, tmpMBMTYpe, Trim(txtPAYCode)), "#,##0.00")
    
 
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
   
End Sub
Private Sub CallLimitDept(tmpMBMTYpe As String, tmpAnnLmt As String, tmpVisitLmt As String, tmpSpecLmt As String, tmpLimitStatus As String)

Const METHOD_NAME = "CallLimitDept"
On Error GoTo debugs

Dim ASORS As New ADODB.Recordset
Dim ASOqry As String
Dim IsCombined As String

Dim ASOlimit As Double
Dim IPLimit As Double
Dim OPLimit As Double

Dim IPUtilization As Double
Dim OPUtilization As Double

Dim sMBMNumber() As String
Dim MBMNumber As String
Dim coverid As String


    ASOlimit = 0
    IPLimit = 0
    OPLimit = 0
    IsCombined = ""
    
    'ASO
    If ASOInd = "Y" Then
    
        MBMNumber = txtMBMNumber
        coverid = DepID
        ASOqry = ""
        ASOqry = "Select MBMASOLimit,MBMIPLimit,MBMOPLimit,MBMIPUtilized,MBMOPUtilized ,IsIPOPCombined  from Tbl_ASOMemberDetails  where MBMNumber= '" + MBMNumber + "' and MBMCoverID='" + coverid + "'"
        ASORS.Open ASOqry, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do Until ASORS.EOF
             ASOlimit = IIf(Len(ASORS!MBMASOLimit), ASORS!MBMASOLimit, 0)
             IPLimit = IIf(Len(ASORS!MBMIPLimit), ASORS!MBMIPLimit, 0)
             OPLimit = IIf(Len(ASORS!MBMOPLimit), ASORS!MBMOPLimit, 0)
             
             IPUtilization = IIf(Len(ASORS!MBMIPUtilized), ASORS!MBMIPUtilized, 0)
             OPUtilization = IIf(Len(ASORS!MBMOPUtilized), ASORS!MBMOPUtilized, 0)
             
             IsCombined = IIf(Len(ASORS!IsIPOPCombined), ASORS!IsIPOPCombined, "N")
        ASORS.MoveNext
        Loop
        ASORS.Close
        Set ASORS = Nothing
        
        If IsCombined = "Y" Then
            tmpAnnLmt = ASOlimit '- (IPUtilization + OPUtilization)
        Else
            tmpAnnLmt = OPLimit '- OPUtilization
        End If
    Else
        tmpAnnLmt = ""
        tmpAnnLmt = Format(GetAnnualLimit(cboINSCode, cboPLNCode, "N", txtMBMPayorEffDate, tmpMBMTYpe, Trim(txtPAYCode), tmpLimitStatus), "#,##0.00")
    End If
    
    tmpVisitLmt = ""
    tmpVisitLmt = Format(GetVisitLimit(cboINSCode, cboPLNCode, "N", txtMBMPayorEffDate, tmpMBMTYpe, Trim(txtPAYCode)), "#,##0.00")
   
    tmpSpecLmt = ""
    tmpSpecLmt = Format(GetSpecLimit(cboINSCode, cboPLNCode, "N", txtMBMPayorEffDate, tmpMBMTYpe, Trim(txtPAYCode)), "#,##0.00")
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub
Private Sub ADJListing()
Const METHOD_NAME = "ADJListing"
On Error GoTo debugs

Dim AdjRec As New ADODB.Recordset
Dim AdjSql As String
Dim ADJList As ListItem

    AdjSql = "SELECT MBMNumber, MBMPolicyNo, MBMEffDate, MBMExpiryDate, MBMAmendEdtType," & _
             " MBMAEndtBordxDate, MBMAEndtEffDate, MBMAmendUser, MBMAmendDate " & _
             " from MBMCISAmendHistory WHERE MBMNumber='" & Trim(txtMBMNumber) & "'"
    AdjRec.Open AdjSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do Until AdjRec.EOF
        If Len(AdjRec!MBMAEndtBordxDate) Then
            Set ADJList = lstAdjustmentCis.ListItems.Add(, , AdjRec!MBMAEndtBordxDate)
        Else
            Set ADJList = lstAdjustmentCis.ListItems.Add(, , " ")
        End If
        If Len(Trim(AdjRec!MBMAmendDate)) Then
            ADJList.SubItems(1) = AdjRec!MBMAmendDate
        End If
        If Len(AdjRec!MBMAmendEdtType) Then
            ADJList.SubItems(2) = AdjRec!MBMAmendEdtType
        End If
        If Len(AdjRec!MBMAEndtEffDate) Then
            ADJList.SubItems(3) = AdjRec!MBMAEndtEffDate
        End If
        If Len(Trim(AdjRec!MBMAmendUser)) Then
            ADJList.SubItems(4) = AdjRec!MBMAmendUser
        End If
            
        AdjRec.MoveNext
    Loop
    AdjRec.Close
    Set AdjRec = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub MBMHISListing()
Const METHOD_NAME = "MBMHISListing"
On Error GoTo debugs

Dim MBMHIS As New ADODB.Recordset
Dim HISList As ListItem
Dim sql As String
    If Len(Trim(txtMBMIcBcPP)) <> 0 Then
        If lstMBMHis.ListItems.Count = 0 Then
            
            sql = "Exec SearchCISMBMHistory '" & Trim(txtMBMIcBcPP) & "'"
            MBMHIS.Open sql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        
            Do While Not MBMHIS.EOF
                Set HISList = lstMBMHis.ListItems.Add(, , MBMHIS!MBMNumber)
                    HISList.SubItems(1) = IIf(Len(MBMHIS!MBMPolNumber), MBMHIS!MBMPolNumber, "")
                    HISList.SubItems(2) = IIf(MBMHIS!MBMPolEffDate <> "", MBMHIS!MBMPolEffDate, "")
                    HISList.SubItems(3) = IIf(MBMHIS!MBMPolExpDate <> "", MBMHIS!MBMPolExpDate, "")
                    If MBMHIS!MBMCLNStatus = "C" Then
                        HISList.SubItems(4) = IIf(Len(MBMHIS!MBMCLNEndtEffDate), MBMHIS!MBMCLNEndtEffDate, "")
                    End If
                    HISList.SubItems(5) = IIf(Len(MBMHIS!MBMCLNDataStatus), MBMHIS!MBMCLNDataStatus, "")
                    HISList.SubItems(6) = IIf(Len(MBMHIS!MBMCLNStatus), MBMHIS!MBMCLNStatus, "")
                    HISList.SubItems(7) = IIf(Len(MBMHIS!MBMBordxDate), MBMHIS!MBMBordxDate, "")
                    HISList.SubItems(8) = IIf(Len(MBMHIS!MBMCLNEndtBordxDate), MBMHIS!MBMCLNEndtBordxDate, "")
                    HISList.SubItems(9) = IIf(Len(MBMHIS!MBMValueDate), MBMHIS!MBMValueDate, "")
                MBMHIS.MoveNext
            Loop
            MBMHIS.Close
            Set MBMHIS = Nothing
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub











