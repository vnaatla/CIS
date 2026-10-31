VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMBMCLNAdj 
   Caption         =   "Clinical Membership Adjustment"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11796
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   11796
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameEndt 
      Height          =   855
      Left            =   120
      TabIndex        =   32
      Top             =   1600
      Width           =   11610
      Begin VB.TextBox txtMBMIcBcPP 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         MaxLength       =   30
         TabIndex        =   14
         Top             =   160
         Width           =   1800
      End
      Begin VB.TextBox txtMBMName 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1440
         MaxLength       =   60
         TabIndex        =   12
         Top             =   165
         Width           =   5895
      End
      Begin VB.TextBox txtMBMPolNo 
         Height          =   330
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   13
         Top             =   460
         Width           =   5895
      End
      Begin VB.ComboBox txtPAYCode 
         Enabled         =   0   'False
         Height          =   288
         Left            =   10680
         Style           =   1  'Simple Combo
         TabIndex        =   11
         Top             =   160
         Width           =   840
      End
      Begin VB.Label lblLimitStatus 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   7560
         TabIndex        =   146
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label80 
         Caption         =   "Member Name"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   210
         Width           =   1140
      End
      Begin VB.Label Label73 
         Caption         =   "Policy Number"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   495
         Width           =   1140
      End
      Begin VB.Label Label64 
         Caption         =   "IC. Num"
         Height          =   255
         Left            =   7560
         TabIndex        =   34
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label74 
         Caption         =   "Payor"
         Height          =   255
         Left            =   10200
         TabIndex        =   33
         Top             =   225
         Width           =   915
      End
   End
   Begin VB.Frame Frame10 
      Height          =   850
      Left            =   120
      TabIndex        =   37
      Top             =   760
      Width           =   11610
      Begin VB.ComboBox cboEndorseStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmMBMCLNAdj.frx":0000
         Left            =   1440
         List            =   "FrmMBMCLNAdj.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   2055
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "FrmMBMCLNAdj.frx":0004
         Left            =   5640
         List            =   "FrmMBMCLNAdj.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   180
         Width           =   1680
      End
      Begin MSMask.MaskEdBox txtEndtEffDate 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   460
         Width           =   2055
         _ExtentX        =   3620
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtEndtBordxDate 
         Height          =   285
         Left            =   5640
         TabIndex        =   9
         Top             =   465
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   508
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboDataStatus 
         Height          =   315
         ItemData        =   "FrmMBMCLNAdj.frx":0008
         Left            =   9240
         List            =   "FrmMBMCLNAdj.frx":000A
         TabIndex        =   7
         Text            =   "A - Active"
         Top             =   160
         Width           =   1680
      End
      Begin VB.ComboBox cboExpStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmMBMCLNAdj.frx":000C
         Left            =   9240
         List            =   "FrmMBMCLNAdj.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   420
         Width           =   1680
      End
      Begin VB.Label Label89 
         Caption         =   "Exp. Status"
         Height          =   255
         Left            =   8160
         TabIndex        =   43
         Top             =   465
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Endt Bordx Date"
         Height          =   255
         Left            =   4200
         TabIndex        =   42
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblAdj 
         Caption         =   "Endt Eff Date"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label62 
         Caption         =   "Data Status"
         Height          =   255
         Left            =   8160
         TabIndex        =   40
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label55 
         Caption         =   "Pol. Status"
         Height          =   255
         Left            =   4200
         TabIndex        =   38
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label24 
         Caption         =   "Endt. Status"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Frame Frame7 
      Height          =   525
      Left            =   0
      TabIndex        =   53
      Top             =   7560
      Width           =   11655
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   10500
         TabIndex        =   56
         Top             =   150
         Width           =   960
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9555
         TabIndex        =   55
         Top             =   150
         Width           =   960
      End
      Begin VB.CommandButton cmdEditMBM 
         Caption         =   "E&dit Status"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8640
         TabIndex        =   54
         Top             =   150
         Width           =   960
      End
   End
   Begin VB.Frame Frame16 
      Height          =   540
      Left            =   100
      TabIndex        =   30
      Top             =   240
      Width           =   11730
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Pre"
         Height          =   285
         Left            =   5040
         TabIndex        =   2
         Top             =   160
         Width           =   810
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
         Height          =   285
         Left            =   5880
         TabIndex        =   3
         Top             =   160
         Width           =   810
      End
      Begin VB.CommandButton cmdSearchMBM 
         Caption         =   "&Search"
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
         Left            =   6720
         TabIndex        =   4
         Top             =   160
         Width           =   810
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Go"
         Height          =   285
         Left            =   4200
         TabIndex        =   1
         Top             =   160
         Width           =   810
      End
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   0
         Top             =   160
         Width           =   2055
      End
      Begin VB.Label Label76 
         Caption         =   "Member Number"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   200
         Width           =   1275
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4995
      Left            =   120
      TabIndex        =   44
      Top             =   2520
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   8805
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Principal &Details"
      TabPicture(0)   =   "FrmMBMCLNAdj.frx":0010
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePrincDet"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Supplementary"
      TabPicture(1)   =   "FrmMBMCLNAdj.frx":002C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameSupp"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstCovAdd"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Ajustment History"
      TabPicture(2)   =   "FrmMBMCLNAdj.frx":0048
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstAdjustmentCis"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "MBM History"
      TabPicture(3)   =   "FrmMBMCLNAdj.frx":0064
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstMBMHis"
      Tab(3).ControlCount=   1
      Begin VB.Frame FramePrincDet 
         Height          =   4530
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   11325
         Begin VB.TextBox txtCISRemark 
            Height          =   915
            Left            =   5640
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   157
            Top             =   3240
            Width           =   5295
         End
         Begin VB.TextBox txtMBMbankaccno 
            Height          =   285
            Left            =   7080
            TabIndex        =   156
            Top             =   1200
            Width           =   4155
         End
         Begin VB.TextBox txtbankname 
            Height          =   285
            Left            =   1200
            TabIndex        =   155
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtphonenumber 
            Height          =   285
            Left            =   4200
            TabIndex        =   152
            Top             =   840
            Width           =   1515
         End
         Begin VB.TextBox txtEmail 
            Height          =   285
            Left            =   7080
            TabIndex        =   150
            Top             =   840
            Width           =   4155
         End
         Begin VB.Frame Frame2 
            Height          =   1605
            Left            =   120
            TabIndex        =   57
            Top             =   2760
            Width           =   5265
            Begin VB.Line Line1 
               Index           =   1
               X1              =   120
               X2              =   5160
               Y1              =   1200
               Y2              =   1200
            End
            Begin VB.Label Label26 
               Caption         =   "Balance Limit"
               Height          =   255
               Left            =   5400
               TabIndex        =   140
               Top             =   920
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label lblAccumAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6720
               TabIndex        =   139
               Top             =   540
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.Label Label14 
               Caption         =   "Annual Limit"
               Height          =   255
               Left            =   5400
               TabIndex        =   138
               Top             =   165
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label lblAnnLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6720
               TabIndex        =   137
               Top             =   180
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.Label Label105 
               Caption         =   "MEP Limit"
               Height          =   255
               Left            =   2715
               TabIndex        =   83
               Top             =   1290
               Width           =   1095
            End
            Begin VB.Label Label102 
               Caption         =   "First Spec. Limit"
               Height          =   255
               Left            =   8355
               TabIndex        =   82
               Top             =   1300
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.Label lblFirstSpecLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9650
               TabIndex        =   81
               Top             =   1280
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.Label Label94 
               Caption         =   "Dental Ext Limit"
               Height          =   255
               Left            =   8355
               TabIndex        =   80
               Top             =   550
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.Label lblDentalExtLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9650
               TabIndex        =   79
               Top             =   525
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   120
               X2              =   5160
               Y1              =   840
               Y2              =   840
            End
            Begin VB.Label lblMEPLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   78
               Top             =   1280
               Width           =   1200
            End
            Begin VB.Label lblMBMSpecLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9650
               TabIndex        =   77
               Top             =   885
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.Label lblDentalLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9650
               TabIndex        =   76
               Top             =   195
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.Label lblMCOFees 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   75
               Top             =   180
               Width           =   1185
            End
            Begin VB.Label lblMBMPremium 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   74
               Top             =   180
               Width           =   1020
            End
            Begin VB.Label lblMBMPremiumEff 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   73
               Top             =   915
               Width           =   1035
            End
            Begin VB.Label lblMBMPremiumRed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   72
               Top             =   525
               Width           =   1035
            End
            Begin VB.Label lblMCOFee 
               Caption         =   "MCO Fees"
               Height          =   255
               Left            =   2715
               TabIndex        =   71
               Top             =   230
               Width           =   870
            End
            Begin VB.Label Label85 
               Caption         =   "Effective MCO"
               Height          =   255
               Left            =   2715
               TabIndex        =   70
               Top             =   930
               Width           =   1095
            End
            Begin VB.Label Label83 
               Caption         =   "Effective Prem "
               Height          =   255
               Left            =   120
               TabIndex        =   69
               Top             =   930
               Width           =   1140
            End
            Begin VB.Label lblMCOFeesEff 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   68
               Top             =   915
               Width           =   1185
            End
            Begin VB.Label lblMBMVisitLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6720
               TabIndex        =   67
               Top             =   1275
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.Label lblPremium 
               Caption         =   "Premium"
               Height          =   255
               Left            =   120
               TabIndex        =   66
               Top             =   230
               Width           =   960
            End
            Begin VB.Label Label58 
               Caption         =   "Prem Refund"
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   550
               Width           =   1020
            End
            Begin VB.Label lblMBMMCORed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   64
               Top             =   540
               Width           =   1185
            End
            Begin VB.Label Label61 
               Caption         =   "MCO Refund"
               Height          =   255
               Left            =   2715
               TabIndex        =   63
               Top             =   550
               Width           =   1095
            End
            Begin VB.Label Label79 
               Caption         =   "Dental Limit"
               Height          =   255
               Left            =   8355
               TabIndex        =   62
               Top             =   230
               Visible         =   0   'False
               Width           =   1500
            End
            Begin VB.Label Label77 
               Caption         =   "Spec. Limit"
               Height          =   255
               Left            =   8355
               TabIndex        =   61
               Top             =   930
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.Label lblBalLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6720
               TabIndex        =   60
               Top             =   900
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.Line Line6 
               X1              =   120
               X2              =   5160
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label Label5 
               Caption         =   "Accum Amt"
               Height          =   255
               Left            =   5400
               TabIndex        =   59
               Top             =   540
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label11 
               Caption         =   "Visit Limit"
               Height          =   255
               Left            =   5400
               TabIndex        =   58
               Top             =   1275
               Visible         =   0   'False
               Width           =   1140
            End
         End
         Begin VB.TextBox txtMBMExclusion 
            Height          =   800
            Left            =   120
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   1920
            Width           =   5220
         End
         Begin VB.TextBox txtRemarks 
            Enabled         =   0   'False
            Height          =   915
            Left            =   5640
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   1920
            Width           =   5295
         End
         Begin VB.TextBox txtMBMGroupCompany 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7080
            MaxLength       =   150
            TabIndex        =   20
            Top             =   240
            Width           =   4155
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
            Left            =   4200
            TabIndex        =   18
            Top             =   270
            Width           =   1515
            _ExtentX        =   2667
            _ExtentY        =   487
            _Version        =   393216
            MaxLength       =   10
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
            Left            =   4200
            TabIndex        =   19
            Top             =   525
            Width           =   1515
            _ExtentX        =   2667
            _ExtentY        =   508
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboPLNCode 
            Height          =   315
            Left            =   1200
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   200
            Width           =   1815
         End
         Begin VB.ComboBox cboINSCode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1200
            Sorted          =   -1  'True
            TabIndex        =   16
            Text            =   "I"
            Top             =   480
            Width           =   1815
         End
         Begin VB.ComboBox CboCardType 
            Height          =   315
            ItemData        =   "FrmMBMCLNAdj.frx":0080
            Left            =   1200
            List            =   "FrmMBMCLNAdj.frx":0082
            TabIndex        =   17
            Top             =   770
            Width           =   1815
         End
         Begin VB.TextBox txtDept 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7080
            MaxLength       =   150
            TabIndex        =   147
            Top             =   480
            Width           =   4155
         End
         Begin VB.Label lblcisremark 
            Caption         =   "Clinical Remarks"
            Height          =   255
            Left            =   5640
            TabIndex        =   158
            Top             =   2880
            Width           =   2415
         End
         Begin VB.Label Label42 
            Caption         =   "Bank Acc No"
            Height          =   255
            Left            =   5880
            TabIndex        =   154
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label36 
            Caption         =   "Bank Name"
            Height          =   255
            Left            =   120
            TabIndex        =   153
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label32 
            Caption         =   "Phone No"
            Height          =   255
            Left            =   3240
            TabIndex        =   151
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label29 
            Caption         =   "Email"
            Height          =   255
            Left            =   5880
            TabIndex        =   149
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label19 
            Caption         =   "Adult/Child"
            Height          =   255
            Left            =   3120
            TabIndex        =   136
            Top             =   1200
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblDOB 
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   1200
            TabIndex        =   135
            Top             =   1410
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label lblAdultChild 
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   4320
            TabIndex        =   131
            Top             =   1140
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label lblAgeBand 
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   7080
            TabIndex        =   130
            Top             =   525
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label22 
            Caption         =   "DOB"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   1440
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Card Type"
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   810
            Width           =   915
         End
         Begin VB.Label Label40 
            Caption         =   "Principal Warranty "
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1680
            Width           =   1590
         End
         Begin VB.Label Label84 
            Caption         =   "HIS Remarks"
            Height          =   255
            Left            =   5640
            TabIndex        =   51
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label20 
            Caption         =   "Ins Type"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   520
            Width           =   840
         End
         Begin VB.Label Label21 
            Caption         =   "Plan Code"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   260
            Width           =   825
         End
         Begin VB.Label Label18 
            Caption         =   "Grp. Company"
            Height          =   255
            Left            =   5880
            TabIndex        =   48
            Top             =   260
            Width           =   1185
         End
         Begin VB.Label Label15 
            Caption         =   "Eff Date"
            Height          =   255
            Left            =   3240
            TabIndex        =   47
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label16 
            Caption         =   "Exp Date"
            Height          =   255
            Left            =   3240
            TabIndex        =   46
            Top             =   540
            Width           =   915
         End
         Begin VB.Label Label28 
            Caption         =   "Department"
            Height          =   255
            Left            =   5880
            TabIndex        =   148
            Top             =   525
            Width           =   1185
         End
         Begin VB.Label Label10 
            Caption         =   "AgeBand"
            Height          =   255
            Left            =   5880
            TabIndex        =   129
            Top             =   555
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Frame FrameSupp 
         Height          =   3100
         Left            =   -74880
         TabIndex        =   84
         Top             =   360
         Width           =   11460
         Begin VB.Frame Frame3 
            Height          =   1575
            Left            =   120
            TabIndex        =   100
            Top             =   1440
            Width           =   9825
            Begin VB.Label lblCovAnnLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6165
               TabIndex        =   144
               Top             =   180
               Width           =   1035
            End
            Begin VB.Label Label27 
               Caption         =   "Annual Limit"
               Height          =   255
               Left            =   5040
               TabIndex        =   143
               Top             =   195
               Width           =   975
            End
            Begin VB.Label lblCovBalLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6165
               TabIndex        =   142
               Top             =   900
               Width           =   1035
            End
            Begin VB.Label Label13 
               Caption         =   "Balance Limit"
               Height          =   255
               Left            =   5040
               TabIndex        =   141
               Top             =   930
               Width           =   975
            End
            Begin VB.Label lblCovPremEff 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   128
               Top             =   900
               Width           =   1035
            End
            Begin VB.Label Label12 
               Caption         =   "Effective Prem"
               Height          =   255
               Left            =   200
               TabIndex        =   127
               Top             =   930
               Width           =   1140
            End
            Begin VB.Label lblCovMCOEff 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   126
               Top             =   900
               Width           =   1035
            End
            Begin VB.Label Label8 
               Caption         =   "Effective MCO"
               Height          =   255
               Left            =   2640
               TabIndex        =   125
               Top             =   930
               Width           =   1095
            End
            Begin VB.Line Line3 
               X1              =   120
               X2              =   9720
               Y1              =   1200
               Y2              =   1200
            End
            Begin VB.Label Label39 
               Caption         =   "Visit Limit"
               Height          =   255
               Left            =   5040
               TabIndex        =   122
               Top             =   1275
               Width           =   900
            End
            Begin VB.Label Label37 
               Caption         =   "Accum Amt"
               Height          =   255
               Left            =   5040
               TabIndex        =   121
               Top             =   555
               Width           =   975
            End
            Begin VB.Line Line5 
               X1              =   120
               X2              =   9720
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label lblCovAccumAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6165
               TabIndex        =   120
               Top             =   540
               Width           =   1035
            End
            Begin VB.Label Label35 
               Caption         =   "Spec. Limit"
               Height          =   255
               Left            =   7440
               TabIndex        =   119
               Top             =   930
               Width           =   780
            End
            Begin VB.Label Label34 
               Caption         =   "Dental Limit"
               Height          =   255
               Left            =   7440
               TabIndex        =   118
               Top             =   225
               Width           =   1020
            End
            Begin VB.Label Label33 
               Caption         =   "MCO Refund"
               Height          =   255
               Left            =   2640
               TabIndex        =   117
               Top             =   555
               Width           =   975
            End
            Begin VB.Label lblCovMCORed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   116
               Top             =   540
               Width           =   1035
            End
            Begin VB.Label Label31 
               Caption         =   "Prem Refund"
               Height          =   255
               Left            =   200
               TabIndex        =   115
               Top             =   525
               Width           =   1020
            End
            Begin VB.Label Label30 
               Caption         =   "Premium"
               Height          =   255
               Left            =   200
               TabIndex        =   114
               Top             =   180
               Width           =   960
            End
            Begin VB.Label lblCovVisitLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6165
               TabIndex        =   113
               Top             =   1260
               Width           =   1035
            End
            Begin VB.Label Label25 
               Caption         =   "MCO Fees"
               Height          =   255
               Left            =   2640
               TabIndex        =   112
               Top             =   195
               Width           =   870
            End
            Begin VB.Label lblCovPremRed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   111
               Top             =   540
               Width           =   1035
            End
            Begin VB.Label lblCovPrem 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   110
               Top             =   180
               Width           =   1020
            End
            Begin VB.Label lblCovMCO 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   109
               Top             =   180
               Width           =   1035
            End
            Begin VB.Label lblCovDentLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8565
               TabIndex        =   108
               Top             =   195
               Width           =   1035
            End
            Begin VB.Label lblCovSpecLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8565
               TabIndex        =   107
               Top             =   900
               Width           =   1035
            End
            Begin VB.Label lblCovMepLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   106
               Top             =   1260
               Width           =   1035
            End
            Begin VB.Line Line4 
               X1              =   120
               X2              =   9720
               Y1              =   840
               Y2              =   840
            End
            Begin VB.Label lblCovDentExtLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8565
               TabIndex        =   105
               Top             =   540
               Width           =   1035
            End
            Begin VB.Label Label9 
               Caption         =   "Dental Ext Lmt"
               Height          =   255
               Left            =   7440
               TabIndex        =   104
               Top             =   540
               Width           =   1125
            End
            Begin VB.Label lblCov1stSpecLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8565
               TabIndex        =   103
               Top             =   1260
               Width           =   1035
            End
            Begin VB.Label Label6 
               Caption         =   "1st Spec. Lmt"
               Height          =   255
               Left            =   7440
               TabIndex        =   102
               Top             =   1290
               Width           =   1125
            End
            Begin VB.Label Label1 
               Caption         =   "MEP Limit"
               Height          =   255
               Left            =   2640
               TabIndex        =   101
               Top             =   1290
               Width           =   1095
            End
         End
         Begin VB.TextBox txtCovName 
            Enabled         =   0   'False
            Height          =   330
            Left            =   3720
            MaxLength       =   60
            TabIndex        =   26
            Top             =   240
            Width           =   3495
         End
         Begin VB.Frame Frame9 
            Height          =   1185
            Left            =   10320
            TabIndex        =   85
            Top             =   1440
            Width           =   1065
            Begin VB.CommandButton CmdCancelSupp 
               Caption         =   "Cancel"
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   88
               Top             =   800
               Width           =   840
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "Update"
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   87
               Top             =   160
               Width           =   840
            End
            Begin VB.CommandButton cmdClear 
               Caption         =   "Clear"
               Height          =   315
               Left            =   120
               TabIndex        =   86
               Top             =   480
               Width           =   840
            End
         End
         Begin VB.TextBox txtCovIcBcPp 
            Enabled         =   0   'False
            Height          =   330
            Left            =   8400
            MaxLength       =   30
            TabIndex        =   29
            Top             =   240
            Width           =   2895
         End
         Begin VB.ComboBox cboCovRelationship 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1140
            TabIndex        =   23
            Text            =   "W -Wife"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtCOVCLNRemarks 
            Height          =   795
            Left            =   8400
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   540
            Width           =   2895
         End
         Begin VB.TextBox txtSuppPLNCode 
            Height          =   285
            Left            =   1140
            MaxLength       =   4
            TabIndex        =   24
            Top             =   530
            Width           =   1455
         End
         Begin VB.TextBox txtCovExclusion 
            Height          =   795
            Left            =   3720
            MaxLength       =   2500
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   540
            Width           =   3495
         End
         Begin VB.ComboBox CboCovCardType 
            Height          =   315
            ItemData        =   "FrmMBMCLNAdj.frx":0084
            Left            =   1140
            List            =   "FrmMBMCLNAdj.frx":0086
            TabIndex        =   25
            Top             =   780
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "Card Type"
            Height          =   255
            Left            =   120
            TabIndex        =   145
            Top             =   840
            Width           =   915
         End
         Begin VB.Label lblCovDOB 
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   1140
            TabIndex        =   134
            Top             =   1080
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label17 
            Caption         =   "AgeBand"
            Height          =   255
            Left            =   120
            TabIndex        =   133
            Top             =   1260
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblCovAgeBand 
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   1140
            TabIndex        =   132
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Date of Birth"
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   1095
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label66 
            Caption         =   "Plan Code"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   580
            Width           =   1125
         End
         Begin VB.Label Label92 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   7440
            TabIndex        =   93
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label38 
            Caption         =   "Name"
            Height          =   255
            Left            =   2880
            TabIndex        =   92
            Top             =   280
            Width           =   660
         End
         Begin VB.Label Label41 
            Caption         =   "IC Num"
            Height          =   255
            Left            =   7440
            TabIndex        =   91
            Top             =   280
            Width           =   645
         End
         Begin VB.Label Label52 
            Caption         =   "Relationship"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   285
            Width           =   1005
         End
         Begin VB.Label Label47 
            Caption         =   "Supp. Warranty"
            Height          =   495
            Left            =   2880
            TabIndex        =   89
            Top             =   600
            Width           =   900
         End
      End
      Begin MSComctlLib.ListView lstCovAdd 
         Height          =   1395
         Left            =   -74880
         TabIndex        =   95
         Top             =   3540
         Width           =   11415
         _ExtentX        =   20130
         _ExtentY        =   2455
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
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   28
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
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Warranty/Exclusion"
            Object.Width           =   2646
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Mep Lmt"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   21
            Text            =   "Dental Limit"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   22
            Text            =   "Dent Ext Lmt"
            Object.Width           =   1764
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
            Text            =   "AmendStatus"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lstAdjustmentCis 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   96
         Top             =   360
         Width           =   11355
         _ExtentX        =   20024
         _ExtentY        =   7853
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
            Size            =   8.4
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
         Height          =   4455
         Left            =   -74880
         TabIndex        =   97
         Top             =   360
         Width           =   11415
         _ExtentX        =   20130
         _ExtentY        =   7853
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
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Member Adjustment - Clinical"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   0
      TabIndex        =   98
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "FrmMBMCLNAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADORSTMBM As New ADODB.Recordset
Dim tmpLimitStatus As String

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
            cboCovRelationship.AddItem ChkStr(REL!RELCODE) '& " - " & REL!RELdescription
            REL.MoveNext
        Loop
        REL.Close
        Set REL = Nothing
        
        cboEndorseStatus.AddItem "A" & " - " & "Amendment"
        cboEndorseStatus.AddItem "C" & " - " & "Cancellation"
        cboEndorseStatus.AddItem "K" & " - " & "Cancel Supplementary"
        cboEndorseStatus.AddItem "X" & " - " & "Suspend Membership"
        cboEndorseStatus.AddItem "Y" & " - " & "Re-Activate Membership"
        cboEndorseStatus.AddItem "D" & " - " & "Delete Policy"
        cboEndorseStatus.ListIndex = 0
        
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
            txtEmail = ""
        End If
    Next ctl
    Call ClearCoveredEntry
End Sub

Private Sub CboCardType_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboCardType_KeyPress"
On Error GoTo debugs
KeyAscii = KeyCharacter(KeyAscii)
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(CboCardType.Text, CboCardType.SelStart) + Chr(KeyAscii) + Mid(CboCardType.Text, CboCardType.SelStart + CboCardType.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To CboCardType.ListCount - 1
            
            If NewText = LCase(Left(CboCardType.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = CboCardType.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                If ValidCount = 1 Then
                  CboCardType.Text = ValidValue
                  CboCardType.SelStart = 0
                  CboCardType.SelLength = Len(ValidValue)
                End If
            
            End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CboCovCardType_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboCovCardType_KeyPress"
On Error GoTo debugs

    KeyAscii = KeyCharacter(KeyAscii)
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(CboCovCardType.Text, CboCovCardType.SelStart) + Chr(KeyAscii) + Mid(CboCovCardType.Text, CboCovCardType.SelStart + CboCovCardType.SelLength + 1))
        ValidCount = 0
        ValidValue = ""
        For i = 0 To CboCovCardType.ListCount - 1
            If NewText = LCase(Left(CboCovCardType.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = CboCovCardType.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
                If ValidCount = 1 Then
                  CboCovCardType.Text = ValidValue
                  CboCovCardType.SelStart = 0
                  CboCovCardType.SelLength = Len(ValidValue)
                End If
            End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cboDataStatus_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "cboDataStatus_KeyPress"
On Error GoTo debugs
    KeyAscii = KeyCharacter(KeyAscii)
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboDataStatus.Text, cboDataStatus.SelStart) + Chr(KeyAscii) + Mid(cboDataStatus.Text, cboDataStatus.SelStart + cboDataStatus.SelLength + 1))
        ValidCount = 0
        ValidValue = ""
        For i = 0 To cboDataStatus.ListCount - 1
            If NewText = LCase(Left(cboDataStatus.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboDataStatus.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
                If ValidCount = 1 Then
                  cboDataStatus.Text = ValidValue
                  cboDataStatus.SelStart = 0
                  cboDataStatus.SelLength = Len(ValidValue)
                End If
            End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cboEndorseStatus_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "cboEndorseStatus_KeyPress"
On Error GoTo debugs

    KeyAscii = KeyCharacter(KeyAscii)
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboEndorseStatus.Text, cboEndorseStatus.SelStart) + Chr(KeyAscii) + Mid(cboEndorseStatus.Text, cboEndorseStatus.SelStart + cboEndorseStatus.SelLength + 1))
        ValidCount = 0
        ValidValue = ""
        For i = 0 To cboEndorseStatus.ListCount - 1
            If NewText = LCase(Left(cboEndorseStatus.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboEndorseStatus.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
                If ValidCount = 1 Then
                  cboEndorseStatus.Text = ValidValue
                  cboEndorseStatus.SelStart = 0
                  cboEndorseStatus.SelLength = Len(ValidValue)
                End If
            End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cboEndorseStatus_LostFocus()
Const METHOD_NAME = "cboEndorseStatus_LostFocus"
On Error GoTo debugs

Dim tmpEndtStatus As String

    tmpEndtStatus = Mid(cboEndorseStatus, 1, 1)
    Call EnableEntries(False)
    Call EnableSupplementary(False)
    cboStatus.Enabled = False
    Call EnableSupplementary(False)
    If tmpEndtStatus = "Y" And Mid(cboStatus, 1, 1) <> "C" Then
        MsgBox "No Re-Activate allowed ! ", vbCritical + vbOKOnly
        Exit Sub
    End If
    If tmpEndtStatus <> "Y" And Mid(cboStatus, 1, 1) = "C" Then
        MsgBox "Only Re-Activate allow ! ", vbCritical + vbOKOnly
        Exit Sub
    End If

    Frame10.Enabled = True
    txtEndtBordxDate.Enabled = True
    CmdCancelSupp.Enabled = False
    If tmpEndtStatus = "A" Then
        Call EnableEntries(True)
        Call EnableSupplementary(True)
        txtMBMPayorEffDate.Enabled = True
        txtMBMPayorExpDate.Enabled = True
        cmdUpdate.Enabled = True
    End If
    If tmpEndtStatus = "K" Then
        Call EnableEntries(False)
        Call EnableSupplementary(True)
        CmdCancelSupp.Enabled = True
        txtMBMPayorEffDate.Enabled = False
        txtMBMPayorExpDate.Enabled = False
        cmdUpdate.Enabled = False
        Call EnableSupplementary(True)
        
    End If
    cboEndorseStatus.Enabled = False
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cboExpStatus_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "cboExpStatus_KeyPress"
On Error GoTo debugs

    KeyAscii = KeyCharacter(KeyAscii)
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboExpStatus.Text, cboExpStatus.SelStart) + Chr(KeyAscii) + Mid(cboExpStatus.Text, cboExpStatus.SelStart + cboExpStatus.SelLength + 1))
        ValidCount = 0
        ValidValue = ""
        For i = 0 To cboExpStatus.ListCount - 1
            If NewText = LCase(Left(cboExpStatus.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboExpStatus.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
                If ValidCount = 1 Then
                  cboExpStatus.Text = ValidValue
                  cboExpStatus.SelStart = 0
                  cboExpStatus.SelLength = Len(ValidValue)
                End If
            End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "cboStatus_KeyPress"
On Error GoTo debugs

    KeyAscii = KeyCharacter(KeyAscii)
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboStatus.Text, cboStatus.SelStart) + Chr(KeyAscii) + Mid(cboStatus.Text, cboStatus.SelStart + cboStatus.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboStatus.ListCount - 1
            
            If NewText = LCase(Left(cboStatus.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboStatus.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                If ValidCount = 1 Then
                  cboStatus.Text = ValidValue
                  cboStatus.SelStart = 0
                  cboStatus.SelLength = Len(ValidValue)
                End If
            
            End If
            
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdCancelSupp_Click()
Const METHOD_NAME = "CmdCancelSupp_Click"
On Error GoTo debugs


Dim ForLoop As Integer
Dim ListCount As Integer
Dim RecordSaved
    ListCount = 0
    If txtCovName = "" Then
        MsgBox "Please Choose a Record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            ListCount = lstCovAdd.ListItems.Count
            For ForLoop = 1 To ListCount
                If lstCovAdd.ListItems(ForLoop).Selected = True Then
                    lstCovAdd.SelectedItem.SubItems(27) = "C"
                End If
            Next ForLoop
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ClearCoveredEntry()
    txtCovName = ""
    txtCovIcBcPp = ""
    txtSuppPLNCode = ""
    lblCovDOB = ""
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

Private Sub cmdClear_Click()
    Call ClearCoveredEntry
End Sub

Private Sub cmdEditMBM_Click()
    Call EnableEntries(True)
    Call EnableSupplementary(True)
    cmdUpdate.Enabled = False
    cboEndorseStatus.Enabled = True
    cboEndorseStatus.SetFocus
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrev_Click()
Const METHOD_NAME = "CmdPrev_Click"
On Error GoTo debugs

Dim MBM As New ADODB.Recordset
Dim sql As String
        sql = " select MBMNumber from MBM WHERE HSCLNInd <> 'H' AND MBMNumber < '" & RemStr(txtMBMNumber) & "' " & _
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

Private Sub CmdSave_Click()
Const METHOD_NAME = "CmdSave_Click"
On Error GoTo debugs

On Error GoTo ErrorSave
Dim RecordSaved, confirmCancel
Dim StrSql As String
Dim CAS As New ADODB.Recordset
Dim tmpEndtStatus As String
Dim tmpPremEff As Double
Dim tmpPremRed As Double
Dim tmpMCOEff As Double
Dim tmpMCORed As Double
Dim AA As Integer
Dim bSucess As Boolean
    bSucess = False
    tmpEndtStatus = Mid(cboEndorseStatus, 1, 1)
    If Trim(txtMBMNumber) = "" Then
        MsgBox "Please select a record to edit!", vbCritical
        txtMBMNumber.SetFocus
    ElseIf txtEndtBordxDate = "__/__/____" Then
        MsgBox "Please Enter Endt. Bordx Date!", vbCritical
        txtEndtBordxDate.SetFocus
    ElseIf txtEndtEffDate = "__/__/____" Then
        MsgBox "Please Enter Endt. Effective Date!", vbCritical
        txtEndtEffDate.SetFocus
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            Screen.MousePointer = vbHourglass
                MediConnCISDB.beginTrans
                medixmasterconn.beginTrans
                bSucess = True
                If tmpEndtStatus = "C" Then ' Cancellation
                    StrSql = "Select CASNumber from CISCAS Where MBMNumber ='" & Trim(txtMBMNumber) & "' And CASStatus = 'O'"
                    CAS.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                    If CAS.RecordCount > 0 Then
                       MsgBox "There is at least an outstanding case under this member!", vbOKOnly + vbCritical, DialogBoxTitle
                       CAS.Close
                       Set CAS = Nothing
                       Exit Sub
                    Else
                       CAS.Close
                       Set CAS = Nothing
                       Call Write2MBMnHistory("C", 1)
                       Call WriteAdjStatus2Cov("C", "C", "01")
                       tmpPremEff = 0
                       tmpPremRed = 0
                       tmpMCOEff = 0
                       tmpMCORed = 0
                       Call MemberMCORefund("P", lblAgeBand, Trim(cboPLNCode), lblDOB, lblAdultChild, "00", tmpPremEff, tmpPremRed, tmpMCOEff, tmpMCORed)
                       lblMBMPremiumEff = Format(Round(tmpPremEff, 0), "#,##0.00")
                       lblMCOFeesEff = Format(Round(tmpMCOEff, 0), "#,##0.00")
                       lblMBMPremiumRed = Format(Round(tmpPremRed, 0), "#,##0.00")
                       lblMBMMCORed = Format(Round(tmpMCORed, 0), "#,##0.00")
                       If lstCovAdd.ListItems.Count > 0 Then
                            For AA = 1 To lstCovAdd.ListItems.Count
                                tmpPremEff = 0
                                tmpPremRed = 0
                                tmpMCOEff = 0
                                tmpMCORed = 0
                                Call MemberMCORefund("S", Trim(lstCovAdd.ListItems(AA).SubItems(25)), Trim(lstCovAdd.ListItems(AA).SubItems(8)), lstCovAdd.ListItems(AA).SubItems(7), lstCovAdd.ListItems(AA).SubItems(26), lstCovAdd.ListItems(AA).SubItems(1), tmpPremEff, tmpPremRed, tmpMCOEff, tmpMCORed)
                                lstCovAdd.ListItems(AA).SubItems(11) = tmpPremRed
                                lstCovAdd.ListItems(AA).SubItems(12) = tmpPremEff
                                lstCovAdd.ListItems(AA).SubItems(14) = tmpMCOEff
                                lstCovAdd.ListItems(AA).SubItems(15) = tmpMCOEff
                            Next AA
                       End If
                    End If
                ElseIf tmpEndtStatus = "K" Then 'Cancel Supplementary
                    If lstCovAdd.ListItems.Count > 0 Then
                        For AA = 1 To lstCovAdd.ListItems.Count
                            If Trim(lstCovAdd.ListItems(AA).SubItems(27)) = "C" Then
                                tmpPremEff = 0
                                tmpPremRed = 0
                                tmpMCOEff = 0
                                tmpMCORed = 0
                                Call MemberMCORefund("S", Trim(lstCovAdd.ListItems(AA).SubItems(25)), Trim(lstCovAdd.ListItems(AA).SubItems(8)), lstCovAdd.ListItems(AA).SubItems(7), lstCovAdd.ListItems(AA).SubItems(26), lstCovAdd.ListItems(AA).SubItems(1), tmpPremEff, tmpPremRed, tmpMCOEff, tmpMCORed)
                                lstCovAdd.ListItems(AA).SubItems(11) = tmpPremRed
                                lstCovAdd.ListItems(AA).SubItems(12) = tmpPremEff
                                lstCovAdd.ListItems(AA).SubItems(14) = tmpMCOEff
                                lstCovAdd.ListItems(AA).SubItems(15) = tmpMCOEff
                                Call WriteAdjStatus2Cov("C", tmpEndtStatus, lstCovAdd.ListItems(AA).SubItems(1))
                            End If
                        Next AA
                    End If
                ElseIf tmpEndtStatus = "X" Then   ' Suspend MEMbership
                    Call Write2MBMnHistory("S", 2)
                    Call WriteAdjStatus2Cov("S", "S", "01")
                ElseIf tmpEndtStatus = "Y" Then   ' Re-Activate
                    Call Write2MBMnHistory("A", 0)
                    Call WriteAdjStatus2Cov("A", "A", "01")
                    If Mid(cboStatus, 1, 1) = "C" Then
'                        Call RecalculateMCO
                    End If
                ElseIf tmpEndtStatus = "A" Then  ' Amendment
                    Call Write2MBMnHistory(tmpEndtStatus, 0)
                    'Commented By murali'' Call WriteAdjStatus2Cov(tmpEndtStatus, tmpEndtStatus, "01")
                    Call SaveCLNOthers
                    Call SaveMBMTwo
                    Call SaveMBMOthersTwo
                    'Call SaveMBMCIS
                    Call SaveCisCovPersons
                
                ElseIf tmpEndtStatus = "D" Then  ' Delete Policy
                    Call Write2MBMnHistory("D", 2)
                    Call WriteAdjStatus2Cov("D", tmpEndtStatus, "01")
                End If
                MediConnCISDB.commitTrans
                medixmasterconn.commitTrans
                
                If tmpEndtStatus = "X" Then
                    MsgBox "Record is Suspended", vbOKOnly + vbInformation, DialogBoxTitle
                ElseIf tmpEndtStatus = "D" Then
                    MsgBox "Policy Deleted Successfully", vbOKOnly + vbInformation, DialogBoxTitle
                ElseIf tmpEndtStatus = "Y" Then
                    MsgBox "Record is Activated", vbOKOnly + vbInformation, DialogBoxTitle
                ElseIf tmpEndtStatus = "A" Or tmpEndtStatus = "P" Then
                    MsgBox "Updated Successfully", vbOKOnly + vbInformation, DialogBoxTitle
                ElseIf tmpEndtStatus = "C" Then
                    MsgBox "This Membership has been cancelled successfully", vbOKOnly + vbInformation, DialogBoxTitle
                ElseIf tmpEndtStatus = "K" Then
                    MsgBox "Supplementary has been cancelled successfully", vbOKOnly + vbInformation, DialogBoxTitle
                End If
            End If
            Screen.MousePointer = vbDefault
    End If
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If bSucess = True Then
        medixmasterconn.RollbackTrans
        MediConnCISDB.RollbackTrans
    End If
    
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub cmdSearchMBM_Click()
    Call ClearEntries
    frmSearchMBM.Source = 5
    frmSearchMBM.Show
End Sub

Private Sub CmdUpdate_Click()
Const METHOD_NAME = "CmdUpdate_Click"
On Error GoTo debugs
Dim ForLoop As Integer
Dim RecordSaved
    If Trim(txtCovName) = "" Then
       MsgBox "Please select covered persons to update !", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(txtCovIcBcPp) = "" Then
       MsgBox "Please Enter Covered Person Identification", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            For ForLoop = 1 To lstCovAdd.ListItems.Count
                If lstCovAdd.ListItems(ForLoop).Selected = True Then
                    lstCovAdd.SelectedItem.SubItems(5) = Trim(txtCovExclusion)
                    lstCovAdd.SelectedItem.SubItems(6) = Trim(txtCOVCLNRemarks)
                    lstCovAdd.SelectedItem.SubItems(8) = Trim(txtSuppPLNCode)
                    lstCovAdd.SelectedItem.SubItems(9) = Trim(CboCovCardType)
                End If
            Next ForLoop
        End If
        Call ClearCoveredEntry
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "FrmMBMCLNAdj")
    cboEndorseStatus.Enabled = False
    cboStatus.Enabled = False
    cboExpStatus.Enabled = False
    Call ComboListing
    cmdSave.Enabled = False
    Call EnableEntries(False)
    Call EnableSupplementary(False)
    SSTab1.Tab = 0
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
Private Sub lstAdjustmentCis_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstAdjustmentCis.SortKey = ColumnHeader.Index - 1
    lstAdjustmentCis.Sorted = True
    If lstAdjustmentCis.SortOrder = lvwAscending Then
        lstAdjustmentCis.SortOrder = lvwDescending
    Else
        lstAdjustmentCis.SortOrder = lvwAscending
    End If
End Sub

Private Sub lstCovAdd_Click()
Const METHOD_NAME = "lstCovAdd_Click"
On Error GoTo debugs
Dim MBMRef As New ADODB.Recordset
Dim strMBMRef As String
    If lstCovAdd.ListItems.Count > 0 Then
        txtCovName = lstCovAdd.SelectedItem.Text
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        cboCovRelationship.Text = lstCovAdd.SelectedItem.SubItems(3)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(5)
        txtCOVCLNRemarks = lstCovAdd.SelectedItem.SubItems(6)
        lblCovDOB = lstCovAdd.SelectedItem.SubItems(7)
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
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
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
    End If
    Screen.MousePointer = Default
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub txtCOVCLNRemarks_LostFocus()
    txtCOVCLNRemarks = ChkStr(txtCOVCLNRemarks)
End Sub

Private Sub txtEndtBordxDate_LostFocus()
    If IsDate(txtEndtBordxDate) Then
    Else
        If txtEndtBordxDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtEndtBordxDate.SetFocus
        End If
    End If
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

Private Sub txtEndtEffDate_LostFocus()
    If IsDate(txtEndtEffDate) Then
    Else
        If txtEndtEffDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtEndtEffDate.SetFocus
        End If
    End If
End Sub

Private Sub txtEndtBordxDate_GotFocus()
    If txtEndtEffDate <> "__/__/____" Then
        cmdSave.Enabled = True
    End If
End Sub

Private Sub CmdNext_Click()
Const METHOD_NAME = "CmdNext_Click"
On Error GoTo debugs

    Dim MBM As New ADODB.Recordset
    Dim sql As String
        sql = " Select MBMNumber from MBM Where HSCLNInd <> 'H' AND MBMNumber > '" & RemStr(txtMBMNumber) & "' " & _
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
On Error GoTo ErrorShow
Dim MBMNumber As String
Dim tmpRecCnt As String

        cmdSave.Enabled = False
        txtEndtEffDate.mask = ""
        txtEndtEffDate.mask = "__/__/____"
        txtEndtEffDate.mask = "##/##/####"
        txtEndtBordxDate.mask = ""
        txtEndtBordxDate.mask = "__/__/____"
        txtEndtBordxDate.mask = "##/##/####"
        
        MBMNumber = txtMBMNumber
        Call ClearEntries
        txtMBMNumber = MBMNumber
        If Trim(txtMBMNumber) = "" Then
            MsgBox "Please Enter Membership No", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMNumber.SetFocus
        Else
            tmpRecCnt = 1
            If Len(Trim(txtMBMNumber)) = 8 Or Len(Trim(txtMBMNumber)) = 9 Then
                frmSearchMBM05.Source = 4
                frmSearchMBM05.Show
                tmpRecCnt = frmSearchMBM05.SearchRecCnt
                If frmSearchMBM05.SearchRecCnt < 1 Then
                    Unload frmSearchMBM05
                ElseIf frmSearchMBM05.SearchRecCnt = 1 Then
                    txtMBMNumber = frmSearchMBM05.tmpMBMNo
                    Unload frmSearchMBM05
                End If
            End If
            
            If tmpRecCnt < 2 Then
                Screen.MousePointer = vbHourglass
                Call FindMBM
                Screen.MousePointer = Default
            End If
        End If
    
    Exit Sub
ErrorShow:
    Screen.MousePointer = vbDefault
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
    cmdEditMBM.Enabled = True
    If Mid(cboStatus, 1, 1) = "D" Then
        MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdEditMBM.Enabled = False
        txtMBMNumber.SetFocus
    ElseIf Mid(cboStatus, 1, 1) = "C" Then
        MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdEditMBM.Enabled = True
        txtMBMNumber.SetFocus
    ElseIf Mid(cboStatus, 1, 1) = "S" Then
        MsgBox "Membership has been Suspended. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdEditMBM.Enabled = False
        txtMBMNumber.SetFocus
    End If
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
Dim MBMO As New ADODB.Recordset
Dim tmpAnnLmt As String
Dim tmpVisitLmt As String
Dim AnnLmtRec As New ADODB.Recordset

    With ADORSTMBM
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
        lblDOB = !MBMDOB
        lblLimitStatus = "N"
        lblAdultChild = !MBMAdultChild
        MBMO.Open "SELECT MBMGroupCompany, MBMDepartment from MBMOthers where MBMNumber='" & txtMBMNumber & "'", medixmasterconn, adOpenForwardOnly, adLockReadOnly
        If MBMO.EOF = False Then
            txtDept = IIf(Len(MBMO!MBMDepartment), MBMO!MBMDepartment, "")
            txtMBMGroupCompany = IIf(Len(MBMO!MBMGroupCompany), MBMO!MBMGroupCompany, "")
        End If
        MBMO.Close
        Set MBMO = Nothing
        lblDOB = !MBMDOB
        lblLimitStatus = "N"
        lblAdultChild = !MBMAdultChild
        AnnLmtRec.Open "SELECT SuppLimitStatus FROM AnnualLimit WHERE PLNCode='" & Trim(!PLNType) & _
        "' AND INSCode='" & Trim(!INSType) & "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        If AnnLmtRec.EOF = False Then
            lblLimitStatus = IIf(Len(AnnLmtRec!SuppLimitStatus), AnnLmtRec!SuppLimitStatus, "A")
            tmpLimitStatus = IIf(Len(AnnLmtRec!SuppLimitStatus), AnnLmtRec!SuppLimitStatus, "A")
        End If
        AnnLmtRec.Close
        Set AnnLmtRec = Nothing
        If Trim(lblLimitStatus) = "N" Then
            lblLimitStatus = "Unlimited Limit"
        ElseIf Trim(lblLimitStatus) = "A" Then
            lblLimitStatus = "Share Limit For Family"
        ElseIf Trim(lblLimitStatus) = "B" Then
            lblLimitStatus = "Share Limit For Supplementary"
        ElseIf Trim(lblLimitStatus) = "C" Then
            lblLimitStatus = "Separate Limit For Supplementary"
        ElseIf Trim(lblLimitStatus) = "D" Then
            lblLimitStatus = "Share For Supplementary and Full For Principle"
        ElseIf Trim(lblLimitStatus) = "E" Then
            lblLimitStatus = "Separate For Supplementary and Full For Principle"
        End If
        txtMBMPayorEffDate = !MBMPolEffDate
        txtMBMPayorExpDate = !MBMPolExpDate
        lblAgeBand = MemberCallAge(txtMBMPayorEffDate, lblDOB)
        CboCardType.Text = IIf(Len(!MBMCardType), !MBMCardType, "")
        txtMBMExclusion = IIf(Len(!MBMCLNWarranty), !MBMCLNWarranty, "")
        txtRemarks = IIf(Len(!MBMRemarks), !MBMRemarks, "")
        txtCISRemark = IIf(Len(!MBMCLNRemarks), !MBMCLNRemarks, "")
        lblMBMPremium = Format(!MBMCLNPremium, "#,##0.00")
        lblMCOFees = Format(!MBMCLNMCOFee, "#,##0.00")
        lblAccumAmt = Format(!MBMAccumAmt, "#,##0.00")
        
        txtEmail = IIf(Len(!MBMEmail), !MBMEmail, "")
        txtphonenumber = IIf(Len(!MBMTelNoM), !MBMTelNoM, "")
        txtbankname = IIf(Len(!MBMBankName), !MBMBankName, "")
        txtMBMbankaccno = IIf(Len(!MBMBankACNo), !MBMBankACNo, "")
        
        Call CallLimit("P", tmpAnnLmt, tmpVisitLmt, tmpLimitStatus)
        lblAnnLmt = tmpAnnLmt
        lblMBMVisitLimit = tmpVisitLmt
        If IsNumeric(tmpVisitLmt) Then
            lblMBMVisitLimit = Format(!MBMCLNVisitLmt, "#,##0.00")
        End If
        lblBalLimit = tmpAnnLmt
        If IsNumeric(tmpAnnLmt) Then
            lblBalLimit = Format(!MBMCLNBalLmt, "#,##0.00")
        End If
        
        lblMEPLimit = Format(!MBMMEPLmt, "#,##0.00")
        lblDentalLimit = Format(!MBMDentalLmt, "#,##0.00")
        lblMBMSpecLimit = Format(!MBMSpecLmt, "#,##0.00")
        lblDentalExtLimit = Format(!MBMDentalExtLmt, "#,##0.00")
        lblFirstSpecLimit = Format(!MBMFirstSpecLmt, "#,##0.00")
    End With
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub showMBMCov()
Const METHOD_NAME = "showMBMCov"
On Error GoTo debugs

Dim MBMCoVRec As ListItem
Dim strMBMRef, tmpAnnLmt As String
Dim MBMRef As New ADODB.Recordset
Dim tmpVisitLmt As String
Dim PLnRec As New ADODB.Recordset
Dim StrPlnSql As String
Dim tmpSuppLmtStatus As String

    lstCovAdd.ListItems.Clear
    
    Do While Not ADORSTMBM.EOF
        With ADORSTMBM
            Set MBMCoVRec = lstCovAdd.ListItems.Add(, , IIf(Trim(!MBMCName) <> "", !MBMCName, ""))
            
            MBMCoVRec.SubItems(1) = !MBMCCoverID
            MBMCoVRec.SubItems(2) = IIf(Len(Trim(!MBMCICBCPPNo)), !MBMCICBCPPNo, "")
            MBMCoVRec.SubItems(3) = IIf(Len(Trim(!MBMCRELCode)), !MBMCRELCode, "")
            MBMCoVRec.SubItems(4) = IIf(Len(!MBMCCLNStatus), !MBMCCLNStatus, "")
            MBMCoVRec.SubItems(5) = IIf(Len(Trim(!MBMCCLNWarranty)), Trim(!MBMCCLNWarranty), "")
            MBMCoVRec.SubItems(6) = IIf(Len(Trim(!MBMCCLNRemarks)), !MBMCCLNRemarks, "")
            MBMCoVRec.SubItems(7) = !MBMCDOB
            If Trim(!MBMCCLNPLNCode) <> "" Then
                MBMCoVRec.SubItems(8) = IIf(Len(!MBMCCLNPLNCode), !MBMCCLNPLNCode, "")
            Else
                MBMCoVRec.SubItems(8) = IIf(Len(!PLNType), !PLNType, "")
            End If
            MBMCoVRec.SubItems(9) = IIf(Len(!MBMCCardType), !MBMCCardType, "")
            ' (9) IS FOR THE Amendment status
            MBMCoVRec.SubItems(10) = IIf(IsNumeric(!MBMCCLNPrem), !MBMCCLNPrem, 0)
            ' :- (11) - PremRefund, (12) - EffPrem
            MBMCoVRec.SubItems(13) = IIf(IsNumeric(!MBMCCLNMCO), !MBMCCLNMCO, 0)
            ' :- (14) - MCORefund, (15) - EffMCO
            
            'Call CallLimit("S", tmpAnnLmt, tmpVisitLmt) ' MBMCAccumAmt
            StrPlnSql = "Select SuppLimitStatus from AnnualLimit Where INSCode ='" & Trim(!INSType) & "' And PLNCode ='" & Trim(MBMCoVRec.SubItems(8)) & _
                 "' AND AnnEffDate <= '" & Format(!MBMPolEffDate, "mm/dd/yyyy") & "' order by AnnEffDate  desc"
            PLnRec.Open StrPlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            tmpSuppLmtStatus = "A"
            Do While Not PLnRec.EOF
                tmpSuppLmtStatus = IIf(Len(Trim(PLnRec!SuppLimitStatus)), Trim(PLnRec!SuppLimitStatus), "A")
                PLnRec.MoveNext
            Loop
            PLnRec.Close
            Set PLnRec = Nothing
            MBMCoVRec.SubItems(17) = Format(!MBMAccumAmt, "#,##0.00")
            If tmpSuppLmtStatus <> "A" Then
                MBMCoVRec.SubItems(17) = IIf(IsNumeric(!MBMCAccumAmt), !MBMCAccumAmt, 0)
            End If
            Call CallLimit("S", tmpAnnLmt, tmpVisitLmt, tmpLimitStatus)
            MBMCoVRec.SubItems(16) = tmpAnnLmt
            MBMCoVRec.SubItems(18) = tmpAnnLmt
            If IsNumeric(tmpAnnLmt) Then
                MBMCoVRec.SubItems(18) = Format(!MBMCLNBalLmt, "#,##0.00")
                If tmpSuppLmtStatus <> "A" Then
                    MBMCoVRec.SubItems(18) = Format(!MBMCCLNBalLmt, "#,##0.00")
                End If
            End If
            MBMCoVRec.SubItems(19) = tmpVisitLmt
            If IsNumeric(tmpVisitLmt) Then
                MBMCoVRec.SubItems(19) = Format(!MBMCLNVisitLmt, "#,##0.00")
                If tmpSuppLmtStatus <> "A" Then
                    MBMCoVRec.SubItems(19) = Format(!MBMCCLNVisitLmt, "#,##0.00")
                End If
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

            ' (25) - Ageband, (26) -Adult/Child
        End With
        ADORSTMBM.MoveNext
    Loop
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub EnableEntries(EStatus)
    FrameEndt.Enabled = EStatus
    FramePrincDet.Enabled = EStatus
End Sub

Private Sub EnableSupplementary(status As Boolean)
    FrameSupp.Enabled = status
    txtCovExclusion.Enabled = status
    txtCOVCLNRemarks.Enabled = status
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
    tempMCOFee = InsertMCOFees(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(tempMCOFee, "#,##0.00"))
    tempMCOFee = Format(tempMCOFee, "#,##0.00")

End Sub

Private Sub CallLimit(tmpMBMTYpe As String, tmpAnnLmt As String, tmpVisitLmt As String, tmpLimitStatus As String)

    tmpAnnLmt = ""
    tmpAnnLmt = Format(GetAnnualLimit(cboINSCode, cboPLNCode, "N", txtMBMPayorEffDate, tmpMBMTYpe, Trim(txtPAYCode), tmpLimitStatus), "#,##0.00")
    tmpVisitLmt = ""
    tmpVisitLmt = Format(GetVisitLimit(cboINSCode, cboPLNCode, "N", txtMBMPayorEffDate, tmpMBMTYpe, Trim(txtPAYCode)), "#,##0.00")
    
End Sub

Private Sub txtMBMPayorEffDate_LostFocus()
    Call EffdateExpDateChg
End Sub

Private Sub txtMBMPayorExpDate_LostFocus()
    Call EffdateExpDateChg
End Sub

Private Sub EffdateExpDateChg()
Const METHOD_NAME = "EffdateExpDateChg"
On Error GoTo debugs

Dim tempPremium, tempMCOFee As Single
Dim Confirm2
Dim AA As Integer
Dim tmpMBMTYpe As String

    lblAdultChild = "A"
    lblAgeBand = MemberCallAge(txtMBMPayorEffDate, lblDOB)
    If Trim(lblAgeBand) = "01" Then
        lblAdultChild = "C"
    End If
    tmpMBMTYpe = "P"
    Call MemberCallFees(tmpMBMTYpe, lblAgeBand, lblAdultChild, tempPremium, tempMCOFee)
    lblMBMPremium = tempPremium
    Confirm2 = vbNo
    If Format(lblMCOFees, "#,##0.00") <> tempMCOFee Then
        Confirm2 = MsgBox(" Do you want to change member's MCO Fee?", vbCritical + vbYesNo)
        If Confirm2 = vbYes Then
            lblMCOFees = tempMCOFee
        End If
    End If
    If lstCovAdd.ListItems.Count > 0 Then
        For AA = 1 To lstCovAdd.ListItems.Count
            tmpMBMTYpe = "S"
            lstCovAdd.ListItems(AA).SubItems(26) = "A"
            lstCovAdd.ListItems(AA).SubItems(25) = MemberCallAge(txtMBMPayorEffDate, lstCovAdd.ListItems(AA).SubItems(7))
            
            If Trim(lstCovAdd.ListItems(AA).SubItems(25)) = "01" Then
                lstCovAdd.ListItems(AA).SubItems(26) = "C"
            End If
        
            Call MemberCallFees(tmpMBMTYpe, lstCovAdd.ListItems(AA).SubItems(25), lstCovAdd.ListItems(AA).SubItems(26), tempPremium, tempMCOFee)
            lstCovAdd.ListItems(AA).SubItems(10) = tempPremium
            If Format(lstCovAdd.ListItems(AA).SubItems(13), "#,##0.00") <> tempMCOFee Then
                If Confirm2 = vbYes Then
                    lstCovAdd.ListItems(AA).SubItems(13) = tempMCOFee
                End If
            End If

        Next AA
    End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub Write2MBMnHistory(WriteStatus As String, cboIndex As Integer)
Const METHOD_NAME = "Write2MBMnHistory"
On Error GoTo debugs

Dim MBMClnOth As New ADODB.Recordset
Dim StrSql As String
Dim StrAmendSql As String
    
    StrSql = "Select MBMNumber, MBMPolNumber, MBMCLNLastEndtStatus, MBMCLNStatus, MBMPolEffDate, MBMPolExpDate, " & _
             " MBMCLNEndtEffDate, MBMCLNEndtBordxDate,MBMCancelDate from MBMCLNOthers Where MBMNumber = '" & txtMBMNumber & "'"
    MBMClnOth.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    StrAmendSql = "Insert into MBMCISAmendHistory(MBMNumber, MBMPolicyNo, " & _
                  "MBMEffDate, MBMExpiryDate, MBMAmendEdtType, " & _
                  "MBMAEndtBordxDate, MBMAEndtEffDate, " & _
                  "MBMAmendUser, MBMAmendDate) " & _
                  "Values ('" & MBMClnOth!MBMNumber & "','" & MBMClnOth!MBMPolNumber & "','" & Format(MBMClnOth!MBMPolEffDate, "YYYY/MM/DD") & "'," & _
                  "'" & Format(MBMClnOth!MBMPolExpDate, "YYYY/MM/DD") & "','" & Mid(cboEndorseStatus, 1, 1) & "'," & _
                  "'" & Format(txtEndtBordxDate, "YYYY/MM/DD") & "','" & Format(txtEndtEffDate, "YYYY/MM/DD") & "'," & _
                  "'" & Logon & "','" & Format(Date, "YYYY/MM/DD") & "')"
    MediConnCISDB.Execute StrAmendSql
    
    With MBMClnOth
        !MBMCLNLastEndtStatus = Mid(cboEndorseStatus, 1, 1)
        !MBMCLNStatus = WriteStatus
        !MBMCLNEndtEffDate = txtEndtEffDate
        !MBMCancelDate = txtEndtEffDate
        !MBMCLNEndtBordxDate = txtEndtBordxDate
        MBMClnOth.update
    End With
    
    MBMClnOth.Close
    Set MBMClnOth = Nothing
    
    cboStatus.ListIndex = cboIndex
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub WriteAdjStatus2Cov(AdjStatus As String, AmendStatus As String, tmpCovID As String)
Const METHOD_NAME = "WriteAdjStatus2Cov"
On Error GoTo debugs

Dim strsqlMBMCov As String
Dim MBMCov As New ADODB.Recordset
    
    strsqlMBMCov = "SELECT MBMCNumber, MBMCCLNStatus,MBMCCancelDate From MBMCoveredPersonsCLN  Where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
    If Trim(AmendStatus) = "K" Then
        strsqlMBMCov = strsqlMBMCov & " AND MBMCCoverID='" & Trim(tmpCovID) & "'"
    End If
    MBMCov.Open strsqlMBMCov, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MBMCov.RecordCount Then
        Do Until MBMCov.EOF
            With MBMCov
                !MBMCCLNStatus = AdjStatus
                If Trim(AmendStatus) = "K" Then
                    !MBMCCancelDate = Format(Date, "YYYY/MM/DD")
                End If
            End With
            MBMCov.update
            MBMCov.MoveNext
        Loop
    End If
    MBMCov.Close
    Set MBMCov = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveCLNOthers()
Const METHOD_NAME = "SaveCLNOthers"
On Error GoTo debugs


Dim MBMCLNOthers As New ADODB.Recordset
Dim strsqlCLNOthers As String

    strsqlCLNOthers = " Select MBMNumber, MBMPolNumber, MBMPolEffDate, " & _
                      " MBMPolExpDate, MBMRemarks, MBMCLNPremium, MBMCardType, " & _
                      " MBMCLNMCOFee, MBMCLNStatus, MBMCLNDataStatus, MBMCLNWarranty,PLNType "
    strsqlCLNOthers = strsqlCLNOthers & " From MBMCLNOthers Where MBMNumber ='" & Trim(txtMBMNumber) & "'"
    MBMCLNOthers.Open strsqlCLNOthers, medixmasterconn, adOpenKeyset, adLockOptimistic
    With MBMCLNOthers
         !MBMPolNumber = txtMBMPolNo
         If IsDate(txtMBMPayorEffDate) Then
            !MBMPolEffDate = txtMBMPayorEffDate
         End If
         If IsDate(txtMBMPayorExpDate) Then
            !MBMPolExpDate = txtMBMPayorExpDate
         End If
        '!MBMRemarks = IIf(Len(Trim(txtRemarks)), txtRemarks, "")
        !MBMRemarks = IIf(Len(Trim(txtCISRemark)), txtCISRemark, "")
        !MBMCardType = IIf(Len(CboCardType.Text), CboCardType, "")
        !MBMCLNMCOFee = IIf(IsNumeric(lblMCOFees), lblMCOFees, 0)
        !MBMCLNPremium = IIf(IsNumeric(lblMBMPremium), lblMBMPremium, 0)
        !MBMCLNDataStatus = Mid(cboDataStatus, 1, 1)
        !MBMCLNStatus = Mid(cboStatus, 1, 1)
        !MBMCLNWarranty = IIf(Len(Trim(txtMBMExclusion)), txtMBMExclusion, "")
        !PLNType = IIf(Len(Trim(cboPLNCode)), cboPLNCode, "")
    End With
    MBMCLNOthers.update
    MBMCLNOthers.Close
    Set MBMCLNOthers = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub
Private Sub SaveMBMTwo()
Const METHOD_NAME = "SaveMBMTwo"
On Error GoTo debugs


Dim MBMTwo As New ADODB.Recordset
Dim strsqlMBMTwo As String

    strsqlMBMTwo = "Select MBMTwo.MBMEmail,MBMTwo.MBMTelNoM from HISDB..MBMTwo Where MBMTwo.MBMNumber ='" & Trim(txtMBMNumber) & "'"
    
    MBMTwo.Open strsqlMBMTwo, medixmasterconn, adOpenKeyset, adLockOptimistic

    With MBMTwo
       !MBMEmail = IIf(Len(Trim(txtEmail)), txtEmail, "")
       !MBMTelNoM = IIf(Len(Trim(txtphonenumber)), txtphonenumber, "")
    End With
    MBMTwo.update
    MBMTwo.Close
    Set MBMTwo = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub
Private Sub SaveMBMOthersTwo()
Const METHOD_NAME = "SaveMBMOthersTwo"
On Error GoTo debugs

Dim MBMothersTwo As New ADODB.Recordset
Dim strsqlMBMothersTwo As String
    strsqlMBMothersTwo = "Select MBMOthersTwo.MBMBankName,MBMOthersTwo.MBMBankACNo from HISDB..MBMOthersTwo Where MBMOthersTwo.MBMNumber ='" & Trim(txtMBMNumber) & "'"
    MBMothersTwo.Open strsqlMBMothersTwo, medixmasterconn, adOpenKeyset, adLockOptimistic
    With MBMothersTwo
       !MBMBankName = IIf(Len(Trim(txtbankname)), txtbankname, "")
       !MBMBankACNo = IIf(Len(Trim(txtMBMbankaccno)), txtMBMbankaccno, "")
    End With
    MBMothersTwo.update
    MBMothersTwo.Close
    Set MBMothersTwo = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub txtMBMPolNo_Change()
    txtMBMPolNo = ChkStr(txtMBMPolNo)
End Sub

Private Sub SaveCisCovPersons()
Const METHOD_NAME = "SaveCisCovPersons"
On Error GoTo debugs

Dim CISCov As New ADODB.Recordset
Dim strCISCov As String
Dim AA As Integer
    
    For AA = 1 To lstCovAdd.ListItems.Count
        strCISCov = " Select MBMCCLNPLNCode, MBMCCLNRemarks, MBMCCLNPrem, MBMCCLNMCO, MBMCCLNStatus, MBMCCardType, " & _
                    " MBMCCLNWarranty, MBMCCLNPLNCode, MBMCCLNEffDate, MBMCCLNExpDate From MBMCoveredPersonsCLN " & _
                    " Where MBMCNumber ='" & Trim(txtMBMNumber) & "' and MBMCCoverID='" & Trim(lstCovAdd.ListItems(AA).SubItems(1)) & "'"
        CISCov.Open strCISCov, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        With CISCov
            !MBMCCLNRemarks = IIf(Len(Trim(lstCovAdd.ListItems(AA).SubItems(6))), lstCovAdd.ListItems(AA).SubItems(6), "")
            !MBMCCLNMCO = IIf(IsNumeric(lstCovAdd.ListItems(AA).SubItems(13)), lstCovAdd.ListItems(AA).SubItems(13), 0)
            !MBMCCLNPrem = IIf(IsNumeric(lstCovAdd.ListItems(AA).SubItems(16)), lstCovAdd.ListItems(AA).SubItems(16), 0)
            '!MBMCCLNStatus = Mid(cboStatus, 1, 1)
            !MBMCCLNWarranty = IIf(Len(Trim(lstCovAdd.ListItems(AA).SubItems(5))), lstCovAdd.ListItems(AA).SubItems(5), "")
            !MBMCCardType = IIf(Len(lstCovAdd.ListItems(AA).SubItems(9)), lstCovAdd.ListItems(AA).SubItems(9), "")
            !MBMCCLNPLNCode = IIf(Len(lstCovAdd.ListItems(AA).SubItems(8)), lstCovAdd.ListItems(AA).SubItems(8), "")
            !MBMCCLNEffDate = txtMBMPayorEffDate
            !MBMCCLNExpDate = txtMBMPayorExpDate
        End With
        CISCov.update
        CISCov.Close
        Set CISCov = Nothing
    Next AA
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub MemberMCORefund(tmpMBMTYpe As String, tmpAgeCode As String, tmpPLNCode As String, tmpDOB As String, tmpAdultChild As String, CoverdID As String, tmpPremEff, tmpPremRed, tmpMCOEff, tmpMCORed As Double)
Const METHOD_NAME = "MemberMCORefund"
On Error GoTo debugs

Dim StrMCORed As String
Dim BasicMCOFees As Double
Dim BasicPremium As Double
Dim InsuredCode As String
Dim HLTCode As String
Dim Payor As String
Dim PayEffDate As String
Dim PayExpDate As String
Dim tempMCOFee As Double
Dim tmpPremium As Double
Dim StrSql, sqlREA As String
Dim DeleteRecord
Dim rsDeleteREA As New ADODB.Recordset
   
    HLTCode = "N"
    InsuredCode = cboINSCode
    Payor = Trim(txtPAYCode)
    PayEffDate = txtMBMPayorEffDate
    PayExpDate = txtMBMPayorExpDate
    BasicMCOFees = GetMCO(tmpMBMTYpe, InsuredCode, HLTCode, tmpAdultChild, CDate(PayEffDate), Trim(tmpPLNCode), Payor)
    tempMCOFee = InsertMCOFees(CDate(PayExpDate), CDate(PayEffDate), Format(BasicMCOFees, "#,##0.00"))
    tmpMCORed = Format(MCORefund(CDate(PayExpDate), CDate(txtEndtEffDate), CDate(PayEffDate), tempMCOFee), "#,##0.00")
    BasicPremium = Format(GetPremium(tmpMBMTYpe, InsuredCode, Payor, Trim(tmpAgeCode), tmpPLNCode, "N", CDate(PayEffDate)), "#,##0.00")
    tmpPremium = InsertPremium(CDate(PayExpDate), CDate(PayEffDate), Format(BasicPremium, "#,##0.00"))
    tmpPremRed = Format(MCOPremiumReduction(CDate(PayExpDate), CDate(txtEndtEffDate), CDate(PayEffDate), tmpPremium), "#,##0.00")
    tmpPremEff = Format(CDbl(tmpPremium) - CDbl(tmpPremRed), "#,##0.00")
    tmpMCOEff = Format(CDbl(tempMCOFee) - CDbl(tmpMCORed), "#,##0.00")
    
    StrSql = "Delete From MBMCISRefund Where MBMNumber = '" & Trim(txtMBMNumber) & "' AND MBMCoveredID = '" & Trim(CoverdID) & "' "
    rsDeleteREA.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    StrMCORed = "Insert into MBMCISRefund(MBMNumber, MBMCoveredID, " & _
                  "MBMPREMEffective, MBMPremReduction, " & _
                  "MBMMCOEffective, MBMMCORefund) " & _
                  "Values ('" & txtMBMNumber & "','" & CoverdID & "'," & _
                  "" & tmpPremEff & "," & tmpPremRed & "," & _
                  "" & tmpMCOEff & "," & tmpMCORed & ")"
    MediConnCISDB.Execute StrMCORed
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub MBMCoveredPersonsRefund(MBMCNUmber As String, MBMCCoveredID As String, MBMCCLNPLNCode As String, MBMCDOB As Date, MBMCCLNPrem As Double)

End Sub

Private Sub ADJListing()
Const METHOD_NAME = "MemberMCORefund"
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

Private Sub RecalculateMCO()
Const METHOD_NAME = "RecalculateMCO"
On Error GoTo debugs

    Dim MBMRefund As New ADODB.Recordset
    Dim StrSqlRefund As String
    Dim MBMOthers As New ADODB.Recordset
    Dim StrSqlOhers As String
    Dim PremiumRefundValue As Double
    Dim MCORefundValue As Double
    
    StrSqlRefund = "Select MBMNumber,MBMPremReduction,MBMMCORefund from MBMRefund Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    MBMRefund.Open StrSqlRefund, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
        If MBMRefund.RecordCount Then
            PremiumRefundValue = MBMRefund!MBMPremReduction
            MCORefundValue = MBMRefund!MBMMCORefund
        Else
            PremiumRefundValue = 0
            MCORefundValue = 0
        End If
    MBMRefund.Close
    Set MBMRefund = Nothing

    StrSqlOhers = "Select MBMNumber, MBMCCLNMCO, MBMCCLNPrem from MBMCLNOthers Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    MBMOthers.Open StrSqlOhers, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
        If MBMOthers.RecordCount Then
            If MCORefundValue <> "0" Then
                MBMOthers!MBMCCLNMCO = MCORefundValue
            End If
            If PremiumRefundValue <> "0" Then
                MBMOthers!MBMCCLNPrem = PremiumRefundValue
            End If
        MBMOthers.update
            
        lblMBMPremium = Format(Round(MBMOthers!MBMPremium, 0), "#,##0.00")
        lblMCOFees = Format(Round(MBMOthers!MBMMCOFee, 0), "#,##0.00")
        End If
    MBMOthers.Close
    Set MBMOthers = Nothing


    lblMBMPremiumEff = Format(0, "#,##0.00")
    lblMCOFeesEff = Format(0, "#,##0.00")
    lblMBMPremiumRed = Format(0, "#,##0.00")
    lblMBMMCORed = Format(0, "#,##0.00")
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


