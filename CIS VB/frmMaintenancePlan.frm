VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaintenancePlan 
   Caption         =   "frmMaintenancePlan"
   ClientHeight    =   9972
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   14520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9972
   ScaleWidth      =   14520
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      TabIndex        =   107
      Top             =   0
      Width           =   14325
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Plan Maintenance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   120
         TabIndex        =   108
         Top             =   120
         Width           =   3285
      End
   End
   Begin VB.Frame Frame10 
      Height          =   705
      Left            =   240
      TabIndex        =   104
      Top             =   8640
      Width           =   13455
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   12360
         TabIndex        =   11
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   11520
         TabIndex        =   10
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   3360
         TabIndex        =   6
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   4200
         TabIndex        =   7
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   10680
         TabIndex        =   9
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   5040
         TabIndex        =   8
         Top             =   285
         Width           =   840
      End
      Begin VB.TextBox temptxtEffDate 
         Height          =   285
         Left            =   360
         TabIndex        =   105
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtPrmCode 
         Height          =   285
         Left            =   120
         TabIndex        =   106
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9045
      Left            =   120
      TabIndex        =   113
      Top             =   480
      Width           =   14085
      _ExtentX        =   24850
      _ExtentY        =   15960
      _Version        =   393216
      Tabs            =   8
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "HealthType"
      TabPicture(0)   =   "frmMaintenancePlan.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(2)"
      Tab(0).Control(1)=   "lstHLTList"
      Tab(0).Control(2)=   "Label25"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Insured Type"
      TabPicture(1)   =   "frmMaintenancePlan.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(3)"
      Tab(1).Control(1)=   "lstINSList"
      Tab(1).Control(2)=   "Label26"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Plan Code"
      TabPicture(2)   =   "frmMaintenancePlan.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label63"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cboPLNDiag"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lstPlanList"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame1(1)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Premium"
      TabPicture(3)   =   "frmMaintenancePlan.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameAGEBand"
      Tab(3).Control(1)=   "Frame1(0)"
      Tab(3).Control(2)=   "Frame6"
      Tab(3).Control(3)=   "lstPmPRMList"
      Tab(3).Control(4)=   "Label24"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Annual Limit"
      TabPicture(4)   =   "frmMaintenancePlan.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).Control(1)=   "lstAnnualList"
      Tab(4).Control(2)=   "Frame1(6)"
      Tab(4).Control(3)=   "Label33"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Visit Limit"
      TabPicture(5)   =   "frmMaintenancePlan.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame1(4)"
      Tab(5).Control(1)=   "Frame7"
      Tab(5).Control(2)=   "lstVisitLimit"
      Tab(5).Control(3)=   "Label58"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "MCO Fee"
      TabPicture(6)   =   "frmMaintenancePlan.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame5"
      Tab(6).Control(1)=   "lstMCOList"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Co-Payment (MBM)"
      TabPicture(7)   =   "frmMaintenancePlan.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame1(5)"
      Tab(7).Control(1)=   "lstCoPaymentMBMList"
      Tab(7).Control(2)=   "Label66"
      Tab(7).ControlCount=   3
      Begin VB.Frame Frame1 
         Height          =   735
         Index           =   5
         Left            =   -74760
         TabIndex        =   208
         Top             =   840
         Width           =   10695
         Begin VB.ComboBox cboCOPPLNCode 
            Height          =   315
            ItemData        =   "frmMaintenancePlan.frx":00E0
            Left            =   5640
            List            =   "frmMaintenancePlan.frx":00E2
            Sorted          =   -1  'True
            TabIndex        =   131
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cboCOPHLTType 
            Height          =   315
            ItemData        =   "frmMaintenancePlan.frx":00E4
            Left            =   2160
            List            =   "frmMaintenancePlan.frx":00E6
            Sorted          =   -1  'True
            TabIndex        =   130
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtCOPAmt 
            Height          =   285
            Left            =   9000
            MaxLength       =   5
            TabIndex        =   132
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label14 
            Caption         =   "Plan Code  *"
            Height          =   255
            Index           =   3
            Left            =   4440
            TabIndex        =   211
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label65 
            Caption         =   "Health Type   *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   210
            Top             =   285
            Width           =   1740
         End
         Begin VB.Label Label64 
            Caption         =   "Amount *"
            Height          =   255
            Left            =   8160
            TabIndex        =   209
            Top             =   285
            Width           =   810
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5115
         Index           =   1
         Left            =   120
         TabIndex        =   166
         Top             =   720
         Width           =   13455
         Begin VB.TextBox txtSpecialNote 
            Height          =   735
            Left            =   1680
            MaxLength       =   1500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   274
            Top             =   4320
            Width           =   11655
         End
         Begin VB.Frame Frame8 
            Caption         =   "Supplementary"
            Height          =   960
            Index           =   2
            Left            =   120
            TabIndex        =   216
            Top             =   2520
            Width           =   13215
            Begin VB.TextBox txtSGPLimit 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   5640
               MaxLength       =   4
               TabIndex        =   271
               Top             =   600
               Width           =   1290
            End
            Begin VB.TextBox txtsSPLimit 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   5640
               MaxLength       =   4
               TabIndex        =   270
               Top             =   240
               Width           =   1290
            End
            Begin VB.ComboBox CboDCStatus 
               Height          =   315
               Left            =   3600
               Sorted          =   -1  'True
               TabIndex        =   38
               Top             =   240
               Width           =   1215
            End
            Begin VB.ComboBox CboSMEPStatus 
               Height          =   315
               Left            =   9960
               Sorted          =   -1  'True
               TabIndex        =   37
               Top             =   480
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.ComboBox CboSAnnLimit 
               Height          =   315
               Left            =   1080
               Sorted          =   -1  'True
               TabIndex        =   31
               Top             =   255
               Width           =   975
            End
            Begin VB.ComboBox CboSVisitLimit 
               Height          =   315
               Left            =   1080
               Sorted          =   -1  'True
               TabIndex        =   32
               Top             =   495
               Width           =   975
            End
            Begin VB.ComboBox CboSDentalStatus 
               Height          =   315
               Left            =   3600
               Sorted          =   -1  'True
               TabIndex        =   33
               Top             =   255
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.ComboBox CboSSpecStatus 
               Height          =   315
               Left            =   8400
               Sorted          =   -1  'True
               TabIndex        =   35
               Top             =   255
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.ComboBox CboSDentalLimit 
               Height          =   315
               Left            =   3600
               Sorted          =   -1  'True
               TabIndex        =   34
               Top             =   495
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.ComboBox CboSSpecLimit 
               Height          =   315
               Left            =   8400
               Sorted          =   -1  'True
               TabIndex        =   36
               Top             =   495
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Label92 
               Caption         =   "SP Limit"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4920
               TabIndex        =   269
               Top             =   270
               Width           =   1455
            End
            Begin VB.Label Label91 
               Caption         =   "GP Limit"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4920
               TabIndex        =   268
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label6 
               Caption         =   "DC Status *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2160
               TabIndex        =   240
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label79 
               Caption         =   "Spec. Limit *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7080
               TabIndex        =   228
               Top             =   585
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.Label Label78 
               Caption         =   "Dental Limit *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2160
               TabIndex        =   227
               Top             =   580
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Label77 
               Caption         =   "MEP Status *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   9960
               TabIndex        =   226
               Top             =   240
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Label71 
               Caption         =   "Vst Lmt *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   220
               Top             =   580
               Width           =   1215
            End
            Begin VB.Label Label70 
               Caption         =   "Spec. Status  *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7080
               TabIndex        =   219
               Top             =   255
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.Label Label69 
               Caption         =   "Dental Status  *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2160
               TabIndex        =   218
               Top             =   260
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Label68 
               Caption         =   "Ann Lmt*"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   217
               Top             =   260
               Width           =   1215
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Principal"
            Height          =   1080
            Index           =   1
            Left            =   120
            TabIndex        =   199
            Top             =   1440
            Width           =   13215
            Begin VB.TextBox txtPGPLimit 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   4680
               MaxLength       =   4
               TabIndex        =   267
               Top             =   600
               Width           =   1290
            End
            Begin VB.TextBox txtPSPLimit 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   4680
               MaxLength       =   4
               TabIndex        =   266
               Top             =   240
               Width           =   1290
            End
            Begin VB.ComboBox cboPLNAnnLimit 
               Height          =   315
               Left            =   1080
               Sorted          =   -1  'True
               TabIndex        =   22
               Top             =   255
               Width           =   975
            End
            Begin VB.ComboBox cboPLNVisitLimit 
               Height          =   315
               Left            =   1080
               Sorted          =   -1  'True
               TabIndex        =   23
               Top             =   495
               Width           =   975
            End
            Begin VB.ComboBox cboPLNPremium 
               Height          =   315
               Left            =   2880
               Sorted          =   -1  'True
               TabIndex        =   24
               Top             =   255
               Width           =   975
            End
            Begin VB.ComboBox cboPLNMCO 
               Height          =   315
               Left            =   2880
               Sorted          =   -1  'True
               TabIndex        =   25
               Top             =   495
               Width           =   975
            End
            Begin VB.ComboBox CboPDentalStatus 
               Height          =   315
               Left            =   7560
               Sorted          =   -1  'True
               TabIndex        =   26
               Top             =   255
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.ComboBox CboPSpectStatus 
               Height          =   315
               Left            =   10560
               Sorted          =   -1  'True
               TabIndex        =   28
               Top             =   135
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.ComboBox CboPMEPStatus 
               Height          =   315
               Left            =   10560
               Sorted          =   -1  'True
               TabIndex        =   30
               Top             =   720
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.ComboBox CboPDentalLimit 
               Height          =   315
               Left            =   7560
               Sorted          =   -1  'True
               TabIndex        =   27
               Top             =   495
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.ComboBox CboPSpectLimit 
               Height          =   315
               Left            =   10560
               Sorted          =   -1  'True
               TabIndex        =   29
               Top             =   375
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label lblSPLimit 
               Caption         =   "SP Limit"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3960
               TabIndex        =   265
               Top             =   270
               Width           =   1455
            End
            Begin VB.Label Label90 
               Caption         =   "GP Limit"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3960
               TabIndex        =   264
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label76 
               Caption         =   "Spect Lmt *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9240
               TabIndex        =   225
               Top             =   435
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label75 
               Caption         =   "Dental Limit *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6120
               TabIndex        =   224
               Top             =   555
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Label74 
               Caption         =   "MEP Status *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9240
               TabIndex        =   223
               Top             =   720
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Label73 
               Caption         =   "Spect Status *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9240
               TabIndex        =   222
               Top             =   165
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label72 
               Caption         =   "Dental Status *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6120
               TabIndex        =   221
               Top             =   285
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Label59 
               Caption         =   "Ann Lmt *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   203
               Top             =   280
               Width           =   1215
            End
            Begin VB.Label Label17 
               Caption         =   "Prem *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2160
               TabIndex        =   202
               Top             =   280
               Width           =   1215
            End
            Begin VB.Label Label29 
               Caption         =   "M C O *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2160
               TabIndex        =   201
               Top             =   550
               Width           =   1215
            End
            Begin VB.Label Label62 
               Caption         =   "Vst Lmt *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   200
               Top             =   550
               Width           =   1215
            End
         End
         Begin VB.Frame Frame8 
            Height          =   880
            Index           =   3
            Left            =   120
            TabIndex        =   236
            Top             =   3360
            Width           =   13215
            Begin VB.TextBox txtDesc6 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   9045
               MaxLength       =   200
               TabIndex        =   237
               Top             =   480
               Width           =   1410
            End
            Begin VB.TextBox txtDesc5 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   7485
               MaxLength       =   200
               TabIndex        =   48
               Top             =   480
               Width           =   1410
            End
            Begin VB.TextBox txtDesc4 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5880
               MaxLength       =   200
               TabIndex        =   46
               Top             =   480
               Width           =   1410
            End
            Begin VB.TextBox txtDesc3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4320
               MaxLength       =   200
               TabIndex        =   44
               Top             =   480
               Width           =   1410
            End
            Begin VB.TextBox txtDesc2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2925
               MaxLength       =   200
               TabIndex        =   42
               Top             =   480
               Width           =   1290
            End
            Begin VB.TextBox txtDesc1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1485
               MaxLength       =   200
               TabIndex        =   40
               Top             =   480
               Width           =   1290
            End
            Begin VB.TextBox txtPLN1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   1485
               MaxLength       =   4
               TabIndex        =   39
               Top             =   120
               Width           =   1290
            End
            Begin VB.TextBox txtPLN2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   2925
               MaxLength       =   4
               TabIndex        =   41
               Top             =   120
               Width           =   1290
            End
            Begin VB.TextBox txtPLN3 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   4365
               MaxLength       =   4
               TabIndex        =   43
               Top             =   120
               Width           =   1410
            End
            Begin VB.TextBox txtPLN4 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   5880
               MaxLength       =   4
               TabIndex        =   45
               Top             =   120
               Width           =   1410
            End
            Begin VB.TextBox txtPLN5 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   7485
               MaxLength       =   4
               TabIndex        =   47
               Top             =   120
               Width           =   1410
            End
            Begin VB.TextBox txtPLN6 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   9045
               MaxLength       =   4
               TabIndex        =   49
               Top             =   120
               Width           =   1410
            End
            Begin VB.Label Label36 
               Caption         =   "Plan Code"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   239
               Top             =   280
               Width           =   1350
            End
            Begin VB.Label Label19 
               Caption         =   "Plan Description"
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
               Left            =   240
               TabIndex        =   238
               Top             =   600
               Width           =   1350
            End
         End
         Begin VB.Frame Frame8 
            Height          =   1200
            Index           =   0
            Left            =   120
            TabIndex        =   204
            Top             =   240
            Width           =   13215
            Begin VB.ComboBox cboOPDependent 
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
               ItemData        =   "frmMaintenancePlan.frx":00E8
               Left            =   10560
               List            =   "frmMaintenancePlan.frx":00EA
               TabIndex        =   277
               Text            =   "Y - Yes"
               Top             =   840
               Width           =   975
            End
            Begin VB.ComboBox cmbOPStatus 
               Height          =   315
               ItemData        =   "frmMaintenancePlan.frx":00EC
               Left            =   7080
               List            =   "frmMaintenancePlan.frx":00F9
               Sorted          =   -1  'True
               TabIndex        =   272
               Text            =   "B-Both"
               Top             =   840
               Width           =   2055
            End
            Begin VB.ComboBox cmbHRApproval 
               Height          =   315
               ItemData        =   "frmMaintenancePlan.frx":0128
               Left            =   3840
               List            =   "frmMaintenancePlan.frx":0132
               Sorted          =   -1  'True
               TabIndex        =   262
               Text            =   "N-No"
               Top             =   840
               Width           =   1335
            End
            Begin VB.ComboBox cmbInsured 
               Height          =   315
               ItemData        =   "frmMaintenancePlan.frx":0143
               Left            =   1200
               List            =   "frmMaintenancePlan.frx":014D
               Sorted          =   -1  'True
               TabIndex        =   260
               Text            =   "N-No"
               Top             =   840
               Width           =   1215
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Check2"
               Height          =   255
               Left            =   8640
               TabIndex        =   21
               Top             =   550
               Width           =   255
            End
            Begin VB.ComboBox cboPLNPayor2 
               Height          =   315
               ItemData        =   "frmMaintenancePlan.frx":015E
               Left            =   7080
               List            =   "frmMaintenancePlan.frx":0160
               Sorted          =   -1  'True
               TabIndex        =   18
               Top             =   230
               Width           =   1815
            End
            Begin VB.ComboBox cboPLNHLTCode 
               Height          =   315
               Left            =   1200
               Sorted          =   -1  'True
               TabIndex        =   15
               Top             =   240
               Width           =   1215
            End
            Begin VB.ComboBox cboPLNPayor 
               Height          =   315
               ItemData        =   "frmMaintenancePlan.frx":0162
               Left            =   3840
               List            =   "frmMaintenancePlan.frx":0164
               Sorted          =   -1  'True
               TabIndex        =   17
               Top             =   230
               Width           =   1815
            End
            Begin VB.ComboBox cboPLNCLNContract 
               Height          =   315
               Left            =   1200
               Sorted          =   -1  'True
               TabIndex        =   16
               Top             =   480
               Width           =   1215
            End
            Begin VB.ComboBox CboNoofPlan 
               Height          =   315
               ItemData        =   "frmMaintenancePlan.frx":0166
               Left            =   10200
               List            =   "frmMaintenancePlan.frx":0168
               Sorted          =   -1  'True
               TabIndex        =   19
               Top             =   240
               Width           =   975
            End
            Begin VB.ComboBox txtGrpCom 
               Height          =   315
               Left            =   3840
               Sorted          =   -1  'True
               TabIndex        =   20
               Top             =   480
               Width           =   4695
            End
            Begin VB.ComboBox cboGLStatus 
               Height          =   315
               Left            =   10200
               Sorted          =   -1  'True
               TabIndex        =   258
               Top             =   480
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lblisOPDependent 
               Caption         =   "OP Dependent"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9240
               TabIndex        =   276
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label16 
               Caption         =   "OP Family Status"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5280
               TabIndex        =   273
               Top             =   840
               Width           =   2175
            End
            Begin VB.Label Label89 
               Caption         =   "HR approval"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2520
               TabIndex        =   263
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label88 
               Caption         =   "Insured"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   261
               Top             =   870
               Width           =   1215
            End
            Begin VB.Label Label77 
               Caption         =   "GL Status"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   9000
               TabIndex        =   259
               Top             =   480
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Label86 
               Caption         =   "HIS Pay Code *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5760
               TabIndex        =   241
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label94 
               Caption         =   "No of Plan Code"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9000
               TabIndex        =   235
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label Label30 
               Caption         =   "Cln Cont.  *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   215
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label Label60 
               Caption         =   "Grp Comp.*"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2520
               TabIndex        =   207
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label61 
               Caption         =   "CIS Pay Code *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2520
               TabIndex        =   206
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label30 
               Caption         =   "HLT Type  *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   205
               Top             =   180
               Width           =   1215
            End
         End
         Begin VB.Label Label30 
            Caption         =   "Special Remark"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   275
            Top             =   4320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Index           =   4
         Left            =   -74640
         TabIndex        =   189
         Top             =   720
         Width           =   10875
         Begin VB.ComboBox CboVSuppStatus 
            Height          =   315
            Left            =   7440
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   720
            Width           =   3180
         End
         Begin VB.TextBox TxtVLimitVer 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10200
            MaxLength       =   8
            TabIndex        =   100
            Top             =   1440
            Width           =   420
         End
         Begin VB.ComboBox CboVINSCode 
            Height          =   315
            ItemData        =   "frmMaintenancePlan.frx":016A
            Left            =   1800
            List            =   "frmMaintenancePlan.frx":016C
            Sorted          =   -1  'True
            TabIndex        =   95
            Top             =   1080
            Width           =   3255
         End
         Begin VB.ComboBox CboVHltCode 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   93
            Top             =   360
            Width           =   3255
         End
         Begin VB.ComboBox CboVPLNCode 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   94
            Top             =   720
            Width           =   3255
         End
         Begin VB.TextBox TxtVisitLimit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7440
            MaxLength       =   8
            TabIndex        =   96
            Top             =   360
            Width           =   3180
         End
         Begin VB.TextBox TxtVSuppVisit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7440
            MaxLength       =   8
            TabIndex        =   98
            Top             =   1080
            Width           =   3180
         End
         Begin MSMask.MaskEdBox TxtVLimitEffDate 
            Height          =   315
            Left            =   7440
            TabIndex        =   99
            Top             =   1440
            Width           =   1065
            _ExtentX        =   1884
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label57 
            Caption         =   "Supp. Visit Lmt Status"
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
            Left            =   5520
            TabIndex        =   197
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label56 
            Caption         =   "Visit Limit Version *"
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
            Left            =   8640
            TabIndex        =   196
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label55 
            Caption         =   "Insured Type  *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   195
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label54 
            Caption         =   "Health Type    *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   194
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label53 
            Caption         =   "Plan Code        *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   193
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label52 
            Caption         =   "Visit Limit                   *"
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
            Left            =   5520
            TabIndex        =   192
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label51 
            Caption         =   "Visit Limit Eff. Date  *"
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
            Left            =   5520
            TabIndex        =   191
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label50 
            Caption         =   "Supp. Visit Lmt"
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
            Left            =   5520
            TabIndex        =   190
            Top             =   1110
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         Height          =   615
         Left            =   -74640
         TabIndex        =   186
         Top             =   2520
         Width           =   10935
         Begin VB.ComboBox CboVGrpCom 
            Height          =   315
            ItemData        =   "frmMaintenancePlan.frx":016E
            Left            =   7440
            List            =   "frmMaintenancePlan.frx":0170
            Sorted          =   -1  'True
            TabIndex        =   102
            Top             =   240
            Width           =   3255
         End
         Begin VB.ComboBox CboVPAYCode 
            Height          =   315
            ItemData        =   "frmMaintenancePlan.frx":0172
            Left            =   1800
            List            =   "frmMaintenancePlan.frx":0174
            Sorted          =   -1  'True
            TabIndex        =   101
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label37 
            Caption         =   "Group Company  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   188
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label22 
            Caption         =   "Payor Code    "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   187
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2535
         Left            =   -74640
         TabIndex        =   182
         Top             =   720
         Width           =   10935
         Begin VB.Frame Frame13 
            Caption         =   "Principal Insured Type *"
            Height          =   1455
            Left            =   360
            TabIndex        =   248
            Top             =   1080
            Width           =   5415
            Begin VB.TextBox TxtICAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   4320
               MaxLength       =   8
               TabIndex        =   257
               Top             =   360
               Width           =   945
            End
            Begin VB.TextBox TxtIAAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   1680
               MaxLength       =   8
               TabIndex        =   118
               Top             =   360
               Width           =   945
            End
            Begin VB.TextBox TxtFCAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   4320
               MaxLength       =   8
               TabIndex        =   121
               Top             =   600
               Width           =   945
            End
            Begin VB.TextBox TxtGCAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   4320
               MaxLength       =   8
               TabIndex        =   123
               Top             =   840
               Width           =   945
            End
            Begin VB.TextBox TxtHCAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   4320
               MaxLength       =   8
               TabIndex        =   125
               Top             =   1080
               Width           =   945
            End
            Begin VB.TextBox TxtFAAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   1680
               MaxLength       =   8
               TabIndex        =   120
               Top             =   600
               Width           =   945
            End
            Begin VB.TextBox TxtGAAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   1680
               MaxLength       =   8
               TabIndex        =   122
               Top             =   840
               Width           =   945
            End
            Begin VB.TextBox TxtHAAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   1680
               MaxLength       =   8
               TabIndex        =   124
               Top             =   1080
               Width           =   945
            End
            Begin VB.Label Label14 
               Caption         =   "Grp Family (Child)"
               Height          =   255
               Index           =   10
               Left            =   2760
               TabIndex        =   256
               Top             =   1110
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Grp Family (Adult)"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   255
               Top             =   1110
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Grp Ind. (Child)"
               Height          =   255
               Index           =   8
               Left            =   2760
               TabIndex        =   254
               Top             =   855
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Grp Ind. (Adult)"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   253
               Top             =   855
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Family (Child)"
               Height          =   255
               Index           =   6
               Left            =   2760
               TabIndex        =   252
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label Label14 
               Caption         =   "Family (Adult)"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   251
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Individual (Child)"
               Height          =   255
               Index           =   4
               Left            =   2760
               TabIndex        =   250
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label LblIA 
               Caption         =   "Individual (Adult)"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   249
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Supp. Insured Type *"
            Height          =   1455
            Left            =   5760
            TabIndex        =   243
            Top             =   1080
            Width           =   5175
            Begin VB.TextBox TxtSFAAmt 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1560
               MaxLength       =   8
               TabIndex        =   126
               Top             =   360
               Width           =   825
            End
            Begin VB.TextBox TxtSHAAmt 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1560
               MaxLength       =   8
               TabIndex        =   128
               Top             =   840
               Width           =   825
            End
            Begin VB.TextBox TxtSFCAmt 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3840
               MaxLength       =   8
               TabIndex        =   127
               Top             =   360
               Width           =   825
            End
            Begin VB.TextBox TxtSHCAmt 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3840
               MaxLength       =   8
               TabIndex        =   129
               Top             =   840
               Width           =   825
            End
            Begin VB.Label Label14 
               Caption         =   "Grp Family (Child)"
               Height          =   255
               Index           =   13
               Left            =   2520
               TabIndex        =   247
               Top             =   855
               Width           =   1455
            End
            Begin VB.Label Label14 
               Caption         =   "Grp Family (Adult)"
               Height          =   255
               Index           =   12
               Left            =   240
               TabIndex        =   246
               Top             =   855
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Family (Child)"
               Height          =   255
               Index           =   14
               Left            =   2520
               TabIndex        =   245
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label14 
               Caption         =   "Family (Adult)"
               Height          =   255
               Index           =   15
               Left            =   240
               TabIndex        =   244
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.ComboBox cboSuppMCOStatus 
            Height          =   315
            Left            =   1920
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   115
            Top             =   700
            Width           =   3330
         End
         Begin VB.ComboBox CboMCOHlt 
            Height          =   315
            ItemData        =   "frmMaintenancePlan.frx":0176
            Left            =   1920
            List            =   "frmMaintenancePlan.frx":0178
            Sorted          =   -1  'True
            TabIndex        =   114
            Top             =   360
            Width           =   3345
         End
         Begin VB.ComboBox CboMCOPln 
            Height          =   315
            ItemData        =   "frmMaintenancePlan.frx":017A
            Left            =   7560
            List            =   "frmMaintenancePlan.frx":017C
            Sorted          =   -1  'True
            TabIndex        =   116
            Top             =   360
            Width           =   1185
         End
         Begin MSMask.MaskEdBox txtEffDate 
            Height          =   315
            Left            =   7560
            TabIndex        =   117
            Top             =   700
            Width           =   1185
            _ExtentX        =   2096
            _ExtentY        =   572
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label87 
            Caption         =   "Supp. MCO Status   *"
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
            Left            =   360
            TabIndex        =   242
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label14 
            Caption         =   "Plan Code *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   6120
            TabIndex        =   185
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Health Type     *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   184
            Top             =   390
            Width           =   1575
         End
         Begin VB.Label Label34 
            Caption         =   "MCO  Eff. Date *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6120
            TabIndex        =   183
            Top             =   800
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   -74640
         TabIndex        =   175
         Top             =   3120
         Width           =   10875
         Begin VB.ComboBox CboPayorCode 
            Height          =   315
            ItemData        =   "frmMaintenancePlan.frx":017E
            Left            =   1800
            List            =   "frmMaintenancePlan.frx":0180
            Sorted          =   -1  'True
            TabIndex        =   90
            Top             =   240
            Width           =   3255
         End
         Begin VB.ComboBox CboGrpCom 
            Height          =   315
            ItemData        =   "frmMaintenancePlan.frx":0182
            Left            =   7440
            List            =   "frmMaintenancePlan.frx":0184
            Sorted          =   -1  'True
            TabIndex        =   91
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label5 
            Caption         =   "Payor Code    "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   177
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Group Company  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   176
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame FrameAGEBand 
         Caption         =   "AgeBand"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74640
         TabIndex        =   167
         Top             =   2400
         Width           =   5520
         Begin VB.TextBox txt07 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4680
            MaxLength       =   8
            TabIndex        =   65
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox txt01 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   360
            MaxLength       =   8
            TabIndex        =   59
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox txt02 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1080
            MaxLength       =   8
            TabIndex        =   60
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox txt03 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   61
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox txt04 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2520
            MaxLength       =   8
            TabIndex        =   62
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox txt05 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3240
            MaxLength       =   8
            TabIndex        =   63
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox txt06 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3960
            MaxLength       =   8
            TabIndex        =   64
            Top             =   480
            Width           =   690
         End
         Begin VB.CheckBox chkSTDPremium 
            Caption         =   "Std Premium"
            Height          =   255
            Left            =   360
            TabIndex        =   66
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lbl06 
            Caption         =   "06"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   174
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbl05 
            Caption         =   "05"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   173
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbl04 
            Caption         =   "04"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   172
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbl03 
            Caption         =   "03"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   171
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbl02 
            Caption         =   "02"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   170
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbl01 
            Caption         =   "01"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   169
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "07"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5040
            TabIndex        =   168
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Index           =   3
         Left            =   -74640
         TabIndex        =   163
         Top             =   720
         Width           =   10770
         Begin VB.TextBox txtINScode 
            Height          =   285
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   12
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtINSdesc 
            Height          =   735
            Left            =   2160
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   720
            Width           =   8295
         End
         Begin VB.Label Label15 
            Caption         =   "Insured Type Description *  "
            Height          =   255
            Left            =   210
            TabIndex        =   165
            Top             =   840
            Width           =   2205
         End
         Begin VB.Label Label14 
            Caption         =   "Insured Type           *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   164
            Top             =   435
            Width           =   1965
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   2
         Left            =   -74640
         TabIndex        =   160
         Top             =   720
         Width           =   10860
         Begin VB.TextBox txtHLTcode 
            Height          =   285
            Left            =   2040
            MaxLength       =   1
            TabIndex        =   0
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtHLTdesc 
            Height          =   1095
            Left            =   2040
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Top             =   720
            Width           =   8535
         End
         Begin VB.Label Label12 
            Caption         =   "Health Type         *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   162
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label13 
            Caption         =   "Health Description      *"
            Height          =   255
            Left            =   240
            TabIndex        =   161
            Top             =   720
            Width           =   1770
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Index           =   0
         Left            =   -74640
         TabIndex        =   151
         Top             =   720
         Width           =   10980
         Begin VB.ComboBox cboPmHLTCode 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   51
            Top             =   240
            Width           =   3690
         End
         Begin VB.ComboBox cboPmINSCode 
            Height          =   315
            ItemData        =   "frmMaintenancePlan.frx":0186
            Left            =   1800
            List            =   "frmMaintenancePlan.frx":0188
            Sorted          =   -1  'True
            TabIndex        =   52
            Top             =   600
            Width           =   3690
         End
         Begin VB.ComboBox cboPmPLNCode 
            Height          =   315
            Left            =   7320
            Sorted          =   -1  'True
            TabIndex        =   55
            Top             =   240
            Width           =   3430
         End
         Begin VB.ComboBox cboPmPAYCode 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   53
            Top             =   960
            Width           =   3690
         End
         Begin VB.ComboBox cboPmAgeCode 
            Height          =   315
            Left            =   7320
            Sorted          =   -1  'True
            TabIndex        =   56
            Top             =   600
            Width           =   3430
         End
         Begin VB.TextBox txtPREVersion 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   9840
            MaxLength       =   8
            TabIndex        =   58
            Top             =   945
            Width           =   930
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   1320
            Width           =   3690
         End
         Begin MSMask.MaskEdBox txtPRMEffdate 
            Height          =   315
            Left            =   7320
            TabIndex        =   57
            Top             =   960
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   550
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label32 
            Caption         =   "Eff Date From     *"
            Height          =   255
            Left            =   5880
            TabIndex        =   159
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label48 
            Caption         =   "Supp. Prem Status   "
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
            Left            =   360
            TabIndex        =   158
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Health Type    *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   157
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Label10 
            Caption         =   "Plan Code     *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   156
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Payor Code   "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   155
            Top             =   960
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Insured Type   *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   154
            Top             =   630
            Width           =   1380
         End
         Begin VB.Label Label9 
            Caption         =   "Age Type      "
            Height          =   375
            Left            =   5880
            TabIndex        =   153
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Prem Version  "
            Height          =   255
            Left            =   8640
            TabIndex        =   152
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "AgeBand"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -69120
         TabIndex        =   134
         Top             =   2400
         Width           =   5400
         Begin VB.CheckBox Check1 
            Caption         =   "Std Supplementary Premium"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   840
            Width           =   2775
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4560
            MaxLength       =   8
            TabIndex        =   73
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3840
            MaxLength       =   8
            TabIndex        =   72
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3120
            MaxLength       =   8
            TabIndex        =   71
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   70
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1680
            MaxLength       =   8
            TabIndex        =   69
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   960
            MaxLength       =   8
            TabIndex        =   68
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            MaxLength       =   8
            TabIndex        =   67
            Top             =   480
            Width           =   690
         End
         Begin VB.Label Label38 
            Caption         =   "01"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   141
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label39 
            Caption         =   "02"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   140
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label40 
            Caption         =   "03"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   139
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label41 
            Caption         =   "04"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   138
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label46 
            Caption         =   "05"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   137
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label47 
            Caption         =   "06"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   136
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label11 
            Caption         =   "07"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4920
            TabIndex        =   135
            Top             =   240
            Width           =   375
         End
      End
      Begin MSComctlLib.ListView lstPlanList 
         Height          =   2265
         Left            =   120
         TabIndex        =   50
         Top             =   5880
         Width           =   13440
         _ExtentX        =   23707
         _ExtentY        =   4001
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
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
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Health Type"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Plan Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Plan Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "CIS Payor Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "HIS Payor Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "GrpCom Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Group Company"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Clinic Contract"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Insured"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "HR Approval"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Princ SPLmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Princ GPLmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Dep SPLimt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Dep GPLimit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "OP Family Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Special Remarks"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "OP Dependent"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstHLTList 
         Height          =   3165
         Left            =   -74640
         TabIndex        =   2
         Top             =   3120
         Width           =   10890
         _ExtentX        =   19219
         _ExtentY        =   5588
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Health Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   6068
         EndProperty
      End
      Begin MSComctlLib.ListView lstINSList 
         Height          =   3690
         Left            =   -74640
         TabIndex        =   14
         Top             =   2640
         Width           =   10800
         _ExtentX        =   19050
         _ExtentY        =   6519
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Insured Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstPmPRMList 
         Height          =   2400
         Left            =   -74640
         TabIndex        =   75
         Top             =   3960
         Width           =   10965
         _ExtentX        =   19346
         _ExtentY        =   4233
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Premium Code"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Insured Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Payor Code"
            Object.Width           =   2535
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Health Code"
            Object.Width           =   2535
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Age Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Plan Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Premium Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Supp Premium Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Eff.Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Prem Version"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Status"
            Object.Width           =   353
         EndProperty
      End
      Begin MSComctlLib.ListView lstAnnualList 
         Height          =   2160
         Left            =   -74640
         TabIndex        =   92
         Top             =   4200
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   3810
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
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
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Health Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "index"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Plan Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Insured Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Annual Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Dental Annual Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Spectacles Annual Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "MEP Annual Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Supp. Annual Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Supp. Dental Annual Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Supp. Spectacles Annual Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Supp. MEP Annual Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "Effective Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Text            =   "Version"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   14
            Text            =   "Payor Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Group Company"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   16
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstMCOList 
         Height          =   2715
         Left            =   -74640
         TabIndex        =   119
         Top             =   3600
         Width           =   10890
         _ExtentX        =   19219
         _ExtentY        =   4784
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Health Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Insured Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Plan Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Supp MCO Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "MCO Fees (Adult)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "MCO Fees (Child)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Supp MCO Fees (Adult)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Supp MCO Fees (Child)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Eff.Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstVisitLimit 
         Height          =   2880
         Left            =   -74640
         TabIndex        =   103
         Top             =   3480
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   5080
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Health Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "index"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Plan Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Insured Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Visit Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Supply. Visit Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Visit Limit Eff. Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Payor Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Group Company"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Visit Version"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstCoPaymentMBMList 
         Height          =   4440
         Left            =   -74760
         TabIndex        =   133
         Top             =   1920
         Width           =   11115
         _ExtentX        =   19600
         _ExtentY        =   7832
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Health Type"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Plan Code"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.ComboBox cboPLNDiag 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   213
         Top             =   1680
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2475
         Index           =   6
         Left            =   -74640
         TabIndex        =   142
         Top             =   720
         Width           =   10875
         Begin VB.ComboBox CboHltCode 
            Height          =   315
            Left            =   2040
            Sorted          =   -1  'True
            TabIndex        =   76
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   7440
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   360
            Width           =   3180
         End
         Begin VB.ComboBox CboPlnCode 
            Height          =   315
            Left            =   2040
            Sorted          =   -1  'True
            TabIndex        =   77
            Top             =   650
            Width           =   2535
         End
         Begin VB.ComboBox CboInsType 
            Height          =   315
            ItemData        =   "frmMaintenancePlan.frx":018A
            Left            =   2040
            List            =   "frmMaintenancePlan.frx":018C
            Sorted          =   -1  'True
            TabIndex        =   78
            Top             =   930
            Width           =   2535
         End
         Begin VB.TextBox TxtSupplyAL 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7440
            MaxLength       =   8
            TabIndex        =   84
            Top             =   650
            Width           =   3180
         End
         Begin VB.TextBox TxtAnnualLimit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2040
            MaxLength       =   8
            TabIndex        =   79
            Top             =   1200
            Width           =   2580
         End
         Begin VB.TextBox TxtPDentalAnnLmt 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2040
            MaxLength       =   8
            TabIndex        =   80
            Top             =   1480
            Width           =   2580
         End
         Begin VB.TextBox TxtPSpecAnnLmt 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2040
            MaxLength       =   8
            TabIndex        =   81
            Top             =   1750
            Width           =   2580
         End
         Begin VB.TextBox TxtSDentalAnnLmt 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7440
            MaxLength       =   8
            TabIndex        =   85
            Top             =   930
            Width           =   3180
         End
         Begin VB.TextBox TxtSSpecAnnLmt 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7440
            MaxLength       =   8
            TabIndex        =   86
            Top             =   1200
            Width           =   3180
         End
         Begin VB.TextBox TxtPMEPAnnLmt 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2040
            MaxLength       =   8
            TabIndex        =   82
            Top             =   2040
            Width           =   2580
         End
         Begin VB.TextBox TxtSMEPAnnLmt 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7440
            MaxLength       =   8
            TabIndex        =   87
            Top             =   1480
            Width           =   3180
         End
         Begin MSMask.MaskEdBox txtALEffDate 
            Height          =   315
            Left            =   7440
            TabIndex        =   88
            Top             =   1760
            Width           =   1065
            _ExtentX        =   1884
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtANNVersion 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10200
            MaxLength       =   8
            TabIndex        =   89
            Top             =   1760
            Width           =   420
         End
         Begin VB.Label Label85 
            Caption         =   "Supp. MEP Ann. Limit"
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
            Left            =   5400
            TabIndex        =   234
            Top             =   1600
            Width           =   2175
         End
         Begin VB.Label Label84 
            Caption         =   "MEP Ann. Limit *"
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
            Left            =   360
            TabIndex        =   233
            Top             =   2200
            Width           =   1815
         End
         Begin VB.Label Label83 
            Caption         =   "Supp. Dental Ann. Limit"
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
            Left            =   5400
            TabIndex        =   232
            Top             =   1035
            Width           =   1815
         End
         Begin VB.Label Label82 
            Caption         =   "Supp. Spectacles Ann. Limit"
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
            Left            =   5400
            TabIndex        =   231
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label81 
            Caption         =   "Spectacles Ann. Limit *"
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
            Left            =   360
            TabIndex        =   230
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label80 
            Caption         =   "Dental Ann. Limit *"
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
            Left            =   360
            TabIndex        =   229
            Top             =   1640
            Width           =   1575
         End
         Begin VB.Label Label42 
            Caption         =   "Supp. Ann Lmt"
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
            Left            =   5400
            TabIndex        =   150
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label43 
            Caption         =   "Annual Limit Eff. Date  *"
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
            Left            =   5400
            TabIndex        =   149
            Top             =   1875
            Width           =   1935
         End
         Begin VB.Label Label44 
            Caption         =   "Annual Limit        *"
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
            Left            =   360
            TabIndex        =   148
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label49 
            Caption         =   "Plan Code        *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   147
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label35 
            Caption         =   "Health Type    *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   146
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label45 
            Caption         =   "Insured Type  *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   145
            Top             =   1035
            Width           =   1335
         End
         Begin VB.Label Label23 
            Caption         =   "Annual Limit Version *"
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
            Left            =   8560
            TabIndex        =   144
            Top             =   1870
            Width           =   1695
         End
         Begin VB.Label Label28 
            Caption         =   "Supp. Ann Lmt Status   "
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
            Left            =   5400
            TabIndex        =   143
            Top             =   390
            Width           =   1935
         End
      End
      Begin VB.Label Label66 
         Alignment       =   2  'Center
         Caption         =   "COPAYMENT LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74760
         TabIndex        =   212
         Top             =   1680
         Width           =   1515
      End
      Begin VB.Label Label58 
         Alignment       =   2  'Center
         Caption         =   "VISIT LIMIT LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74760
         TabIndex        =   198
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "PREMIUM LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74640
         TabIndex        =   181
         Top             =   3720
         Width           =   1395
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "INSURED LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74640
         TabIndex        =   180
         Top             =   2400
         Width           =   1275
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "HEALTH LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74640
         TabIndex        =   179
         Top             =   2880
         Width           =   1155
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Caption         =   "ANNUAL LIMIT LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74640
         TabIndex        =   178
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label63 
         Caption         =   "Diagnosis *"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   214
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Plan Code"
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
      Left            =   0
      TabIndex        =   112
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label31 
      Caption         =   "Bold - Searchable        * - Mandatory Fields"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   111
      Top             =   9600
      Width           =   4695
   End
   Begin VB.Label txtTotalRecord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6480
      TabIndex        =   110
      Top             =   9600
      Width           =   1335
   End
   Begin VB.Label Label18 
      Caption         =   "TOTAL RECORDS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5040
      TabIndex        =   109
      Top             =   9600
      Width           =   1455
   End
End
Attribute VB_Name = "frmMaintenancePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Dim bExit As Boolean
Public PRMCriteria As String
Public PLNCriteria As String
Public HLTCriteria As String
Public INSCriteria As String
Public AnnualCriteria As String
Public VisitCriteria As String
Public MCOCriteria As String
Public CoCriteria As String
Public S As Boolean
Public ListClick As Boolean
Dim Strgrpcode As String
Private m_blnDirection(1 To 17) As Boolean

Private Sub cboCOPHLTType_Change()
    If InStr(cboCOPHLTType.Text, "null") Then
       cboCOPHLTType.Enabled = False
    End If
End Sub

Private Sub cboCOPHLTType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCOPHLTType.Text, cboCOPHLTType.SelStart) + Chr(KeyAscii) + Mid(cboCOPHLTType.Text, cboCOPHLTType.SelStart + cboCOPHLTType.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboCOPHLTType.ListCount - 1
 
        If NewText = LCase(Left(cboCOPHLTType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCOPHLTType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboCOPHLTType.Text = ValidValue
               cboCOPHLTType.SelStart = 0
               cboCOPHLTType.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboCOPPLNCode_Change()
    If InStr(cboCOPPLNCode.Text, "null") Then
       cboCOPPLNCode.Enabled = False
    End If
End Sub

Private Sub cboCOPPLNCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCOPPLNCode.Text, cboCOPPLNCode.SelStart) + Chr(KeyAscii) + Mid(cboCOPPLNCode.Text, cboCOPPLNCode.SelStart + cboCOPPLNCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboCOPPLNCode.ListCount - 1
 
        If NewText = LCase(Left(cboCOPPLNCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCOPPLNCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboCOPPLNCode.Text = ValidValue
               cboCOPPLNCode.SelStart = 0
               cboCOPPLNCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboDCStatus_Change()
    If Mid(CboDCStatus, 1, 1) = "Y" Then
        cboPLNPayor.Text = "DC" & " - " & "DIRECT CLIENT"
    Else
        cboPLNPayor.Text = ""
    End If
End Sub

Private Sub CboDCStatus_Click()
    If Mid(CboDCStatus, 1, 1) = "Y" Then
        cboPLNPayor.Text = "DC" & " - " & "DIRECT CLIENT"
    Else
        cboPLNPayor.Text = ""
    End If
End Sub

Private Sub cboPLNAnnLimit_Change()
    If InStr(cboPLNAnnLimit.Text, "null") Then
       cboPLNAnnLimit.Enabled = False
    End If
End Sub

Private Sub cboPLNAnnLimit_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPLNAnnLimit.Text, cboPLNAnnLimit.SelStart) + Chr(KeyAscii) + Mid(cboPLNAnnLimit.Text, cboPLNAnnLimit.SelStart + cboPLNAnnLimit.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPLNAnnLimit.ListCount - 1
 
        If NewText = LCase(Left(cboPLNAnnLimit.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPLNAnnLimit.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboPLNAnnLimit.Text = ValidValue
               cboPLNAnnLimit.SelStart = 0
               cboPLNAnnLimit.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPLNCLNContract_Change()
    If InStr(cboPLNCLNContract.Text, "null") Then
       cboPLNCLNContract.Enabled = False
    End If
End Sub

Private Sub cboPLNCLNContract_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPLNCLNContract.Text, cboPLNCLNContract.SelStart) + Chr(KeyAscii) + Mid(cboPLNCLNContract.Text, cboPLNCLNContract.SelStart + cboPLNCLNContract.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPLNCLNContract.ListCount - 1
 
        If NewText = LCase(Left(cboPLNCLNContract.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPLNCLNContract.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboPLNCLNContract.Text = ValidValue
               cboPLNCLNContract.SelStart = 0
               cboPLNCLNContract.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPLNMCO_Change()
    If InStr(cboPLNMCO.Text, "null") Then
       cboPLNMCO.Enabled = False
    End If
End Sub

Private Sub cboPLNMCO_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPLNMCO.Text, cboPLNMCO.SelStart) + Chr(KeyAscii) + Mid(cboPLNMCO.Text, cboPLNMCO.SelStart + cboPLNMCO.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPLNMCO.ListCount - 1
 
        If NewText = LCase(Left(cboPLNMCO.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPLNMCO.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboPLNMCO.Text = ValidValue
               cboPLNMCO.SelStart = 0
               cboPLNMCO.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPLNPayor_Change()
    If Mid(cboPLNPayor, 1, 2) = "DC" Then
        CboDCStatus.Text = "Y" & " - " & "Yes"
    Else
        CboDCStatus.Text = ""
    End If
End Sub

Private Sub cboPLNPayor_Click()
    If Mid(cboPLNPayor, 1, 2) = "DC" Then
        CboDCStatus.Text = "Y" & " - " & "Yes"
    Else
        CboDCStatus.Text = ""
    End If
End Sub

Private Sub cboPLNPremium_Change()
    If InStr(cboPLNPremium.Text, "null") Then
       cboPLNPremium.Enabled = False
    End If
End Sub

Private Sub cboPLNPremium_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPLNPremium.Text, cboPLNPremium.SelStart) + Chr(KeyAscii) + Mid(cboPLNPremium.Text, cboPLNPremium.SelStart + cboPLNPremium.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPLNPremium.ListCount - 1
 
        If NewText = LCase(Left(cboPLNPremium.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPLNPremium.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboPLNPremium.Text = ValidValue
               cboPLNPremium.SelStart = 0
               cboPLNPremium.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPLNVisitLimit_Change()
    If InStr(cboPLNVisitLimit.Text, "null") Then
       cboPLNVisitLimit.Enabled = False
    End If
End Sub

Private Sub cboPLNVisitLimit_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPLNVisitLimit.Text, cboPLNVisitLimit.SelStart) + Chr(KeyAscii) + Mid(cboPLNVisitLimit.Text, cboPLNVisitLimit.SelStart + cboPLNVisitLimit.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPLNVisitLimit.ListCount - 1
 
        If NewText = LCase(Left(cboPLNVisitLimit.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPLNVisitLimit.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboPLNVisitLimit.Text = ValidValue
               cboPLNVisitLimit.SelStart = 0
               cboPLNVisitLimit.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub Check2_Click()
Dim GRPCompany As New ADODB.Recordset
            
    If Check2.Value = 1 Then
        
        'GrpCompany.Open "Select DISTINCT CoGRPCode, CoGRPName from VIEW_GRPCompany", medixmasterconn, adOpenKeyset, adLockOptimistic
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            GRPCompany.Open "SELECT GRPCoCode as CoGRPCode, GRPCoName AS CoGRPName From GrpCompany where GRPCoCode <> 'ZD01'GROUP BY GRPCoCode, GRPCoName ORDER BY GRPCoCode, GRPCoName", MediConnCISDB, adOpenKeyset, adLockOptimistic
            
        txtGrpCom = ""
        
        Do Until GRPCompany.EOF
            txtGrpCom.AddItem Trim(GRPCompany!CoGRPCode) & " - " & Trim(GRPCompany!CoGRPName)
            GRPCompany.MoveNext
        Loop
        
        GRPCompany.Close
        Set GRPCompany = Nothing
       ' MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
        
    Else
        
        txtGrpCom = ""
        
    End If
End Sub

Private Sub cboSuppMCOStatus_Change()
    If Mid(cboSuppMCOStatus, 1, 1) = "B" Or Mid(cboSuppMCOStatus, 1, 1) = "C" Then
        Frame12.Enabled = True
    Else
        Frame12.Enabled = False
    End If
End Sub

Private Sub cboSuppMCOStatus_Click()
    If Mid(cboSuppMCOStatus, 1, 1) = "B" Or Mid(cboSuppMCOStatus, 1, 1) = "C" Then
        Frame12.Enabled = True
    Else
        Frame12.Enabled = False
    End If
End Sub

'Public Function TryDate(str As String) As Boolean
'On Error GoTo errDate
'    TryDate = True
'    'str = CDate(str)
'        If Mid(str, 7, 4) <= 1900 Or Mid(str, 7, 4) <= "0000" Then
'            TryDate = False
'        ElseIf Mid(str, 4, 2) > 12 Or Mid(str, 4, 2) <= 0 Then
'            TryDate = False
'        ElseIf Mid(str, 1, 2) > 31 Or Mid(str, 1, 2) <= 0 Then
'            TryDate = False
'        End If
    

'Exit Function
'errDate:
'    TryDate = False
'End Function

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

Private Sub Form_KeyPress(KeyAscii As Integer)
Call EnterToTab(KeyAscii)
End Sub

Private Sub Form_Activate()
    SSTab1.Tab = 0
End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "frmMaintenancePlan")
    Frame3.Width = Screen.Width
    CmdRefresh.Enabled = False
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
    Frame4.Enabled = True
    Me.Width = SSTab1.Width + 500
    Me.Height = SSTab1.Height + Frame3.Height + 600

    PRMCriteria = ""
    PLNCriteria = ""
    HLTCriteria = ""
    INSCriteria = ""
    AnnualCriteria = ""
    VisitCriteria = ""
    MCOCriteria = ""
    CoCriteria = ""

End Sub





Private Sub lstCoPaymentMBMList_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If lstCoPaymentMBMList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        cboCOPHLTType = lstCoPaymentMBMList.SelectedItem.Text
        cboCOPPLNCode = lstCoPaymentMBMList.SelectedItem.SubItems(1)
        txtCOPAmt = lstCoPaymentMBMList.SelectedItem.SubItems(2)
       
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If

End Sub

Private Sub lstCoPaymentMBMList_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If lstCoPaymentMBMList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        cboCOPHLTType = lstCoPaymentMBMList.SelectedItem.Text
        cboCOPPLNCode = lstCoPaymentMBMList.SelectedItem.SubItems(1)
        txtCOPAmt = lstCoPaymentMBMList.SelectedItem.SubItems(2)
       
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If

End Sub

Private Sub lstVisitLimit_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstVisitLimit.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboVHltCode = lstVisitLimit.SelectedItem.Text
        CboVPLNCode = lstVisitLimit.SelectedItem.SubItems(2)
        CboVINSCode = lstVisitLimit.SelectedItem.SubItems(3)
        TxtVisitLimit = lstVisitLimit.SelectedItem.SubItems(4)
        TxtVSuppVisit = lstVisitLimit.SelectedItem.SubItems(5)
        TxtVLimitEffDate = lstVisitLimit.SelectedItem.SubItems(6)
        CboVPAYCode = lstVisitLimit.SelectedItem.SubItems(7)
        CboVGrpCom = lstVisitLimit.SelectedItem.SubItems(8)
        TxtVLimitVer = lstVisitLimit.SelectedItem.SubItems(9)
        
        If lstVisitLimit.SelectedItem.SubItems(10) = "A" Then
            CboVSuppStatus.ListIndex = 0
        ElseIf lstVisitLimit.SelectedItem.SubItems(10) = "B" Then
            CboVSuppStatus.ListIndex = 1
        ElseIf lstVisitLimit.SelectedItem.SubItems(10) = "C" Then
            CboVSuppStatus.ListIndex = 2
        End If
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
End Sub

Private Sub lstVisitLimit_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstVisitLimit.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboVHltCode = lstVisitLimit.SelectedItem.Text
        CboVPLNCode = lstVisitLimit.SelectedItem.SubItems(2)
        CboVINSCode = lstVisitLimit.SelectedItem.SubItems(3)
        TxtVisitLimit = lstVisitLimit.SelectedItem.SubItems(4)
        TxtVSuppVisit = lstVisitLimit.SelectedItem.SubItems(5)
        TxtVLimitEffDate = lstVisitLimit.SelectedItem.SubItems(6)
        CboVPAYCode = lstVisitLimit.SelectedItem.SubItems(7)
        CboVGrpCom = lstVisitLimit.SelectedItem.SubItems(8)
        TxtVLimitVer = lstVisitLimit.SelectedItem.SubItems(9)
        
        If lstVisitLimit.SelectedItem.SubItems(10) = "A" Then
            CboVSuppStatus.ListIndex = 0
        ElseIf lstVisitLimit.SelectedItem.SubItems(10) = "B" Then
            CboVSuppStatus.ListIndex = 1
        ElseIf lstVisitLimit.SelectedItem.SubItems(10) = "C" Then
            CboVSuppStatus.ListIndex = 2
        End If
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    Screen.MousePointer = vbHourglass
    
   'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        Select Case SSTab1.Tab
        Case 0
            Call ClearForm(Me)
            lstHLTList.ListItems.Clear
            CmdRefresh.Enabled = False
            CmdEdit.Enabled = False
            CmdSave.Enabled = True
            CmdDelete.Enabled = False
            txtTotalRecord = 0
        Case 1
            Call ClearForm(Me)
            lstINSList.ListItems.Clear
            CmdRefresh.Enabled = False
            CmdEdit.Enabled = False
            CmdSave.Enabled = True
            CmdDelete.Enabled = False
            txtTotalRecord = 0
        Case 2
            Call ClearForm(Me)
            lstPlanList.ListItems.Clear
            If cboPLNPayor.ListCount = 0 Then
                Call PLNComboListing
            End If
            CmdRefresh.Enabled = True
            CmdEdit.Enabled = False
            CmdSave.Enabled = True
            CmdDelete.Enabled = False
            txtTotalRecord = 0
        Case 3
            Call ClearForm(Me)
            CmdEdit.Enabled = False
            CmdDelete.Enabled = False
            CmdRefresh.Enabled = True
            CmdSave.Enabled = True
            Frame4.Enabled = True
            If cboPmHLTCode.ListCount = 0 Then
                Call PRMComboListing
            End If
            lstPmPRMList.ListItems.Clear
            txtTotalRecord = 0
            FrameAGEBand.Enabled = True
            txtPRMEffdate.Enabled = True
            txtPREVersion.Enabled = True
            
            If Mid(Combo3, 1, 1) = "B" Or Mid(Combo3, 1, 1) = "C" Then
                Frame6.Enabled = True
            Else
                Frame6.Enabled = False
            End If
        Case 4
            Call ClearForm(Me)
            'CboGrpCom.Clear
            CmdEdit.Enabled = False
            CmdDelete.Enabled = False
            CmdRefresh.Enabled = True
            CmdSave.Enabled = True
            If CboPlnCode.ListCount = 0 Then
                Call AnnualComboListing
            End If
            lstAnnualList.ListItems.Clear
            txtTotalRecord = 0
           
            If Mid(Combo3, 1, 1) = "B" Or Mid(Combo3, 1, 1) = "C" Then
                TxtSupplyAL.Enabled = True
            Else
                TxtSupplyAL.Enabled = False
            End If
           
        Case 5
            Call ClearForm(Me)
            'CboGrpCom.Clear
            CmdEdit.Enabled = False
            CmdDelete.Enabled = False
            CmdRefresh.Enabled = True
            CmdSave.Enabled = True
            If CboPlnCode.ListCount = 0 Then
                Call VisitComboListing
            End If
            lstVisitLimit.ListItems.Clear
            txtTotalRecord = 0
           
            If Mid(CboVSuppStatus, 1, 1) = "B" Or Mid(CboVSuppStatus, 1, 1) = "C" Then
                TxtVSuppVisit.Enabled = True
            Else
                TxtVSuppVisit.Enabled = False
            End If
        
        Case 6
            Call ClearForm(Me)
            EditFlag = False
            CmdEdit.Enabled = False
            CmdDelete.Enabled = False
            CmdRefresh.Enabled = True
            CmdSave.Enabled = True
            'If CboMCOINS.ListCount = 0 Then
                Call MCOComboListing
            'End If
            lstMCOList.ListItems.Clear
            txtTotalRecord = 0
        
        Case 7
            Call ClearForm(Me)
            lstCoPaymentMBMList.ListItems.Clear
            CmdRefresh.Enabled = False
            CmdEdit.Enabled = False
            CmdSave.Enabled = True
            CmdDelete.Enabled = False
            txtTotalRecord = 0
            Call CoMBMComboListing
        End Select
  '  MediConnCISDB.Close
  '  Set MediConnCISDB = Nothing
  '  medixmasterconn.Close
  ''  Set medixmasterconn = Nothing
  '  MediConnCISWS.Close
  '  Set MediConnCISWS = Nothing
  '  MediConnCISPayment.Close
   ' Set MediConnCISPayment = Nothing
    Screen.MousePointer = Default
End Sub


'#################### Start Add, Edit, Delete, Search, Save, Exit ####################'

Private Sub CmdAdd_Click()
'    EditFlag = False
    txtTotalRecord = 0
    Call ClearForm(Me)
    Call EntriesEnable(True)
    CmdSave.Enabled = True
'    cboPmAgeCode.Enabled = False
    
    Select Case SSTab1.Tab
        Case 0
            txtHLTcode.SetFocus
        Case 1
            txtINScode.SetFocus
        Case 2
            cboPLNHLTCode.SetFocus
            Frame1(1).Enabled = True
        Case 3
            FrameAGEBand.Enabled = True
            If Mid(Combo3, 1, 1) = "B" Or Mid(Combo3, 1, 1) = "C" Then
                Frame6.Enabled = True
            Else
                Frame6.Enabled = False
            End If
            cboPmHLTCode.SetFocus
        Case 4
            CboHltCode.SetFocus
            Frame1(6).Enabled = True
            Frame4.Enabled = True
        Case 5
            CboVHltCode.SetFocus
            Frame1(4).Enabled = True
            Frame4.Enabled = True
        Case 6
            CboMCOHlt.SetFocus
            Call setCmdButton(Me, False)
            CmdSave.Enabled = True
            CmdDelete.Enabled = False
        Case 7
            cboCOPHLTType.SetFocus
    End Select

End Sub

Private Sub CmdEdit_Click()
 
    EditFlag = True
    Call EntriesEnable(True)
    CmdSave.Enabled = True
    CmdAdd.Enabled = True
    CmdDelete.Enabled = False
 
          If txtHLTcode <> "" Then
             txtHLTcode.Enabled = False
             txtHLTdesc.SetFocus
          End If
          
          If txtINScode <> "" Then
             txtINScode.Enabled = False
             txtINSdesc.Enabled = True
             txtINSdesc.SetFocus
          End If
          
          If cboPLNHLTCode <> "" Then
             cboPLNHLTCode.Enabled = True
             cboPLNPayor.Enabled = True
             Frame1(1).Enabled = True
             Call PLNDisEnable
          End If
          
          If cboPmINSCode <> "" Then
             cboPmPAYCode.Enabled = False
             cboPmHLTCode.Enabled = False
             cboPmAgeCode.Enabled = False
             cboPmPLNCode.Enabled = False
             cboPmINSCode.Enabled = False
             txtPRMEffdate.Enabled = False
          End If
          
          If Mid(Combo1, 1, 1) = "A" Then
            TxtSupplyAL.Enabled = False
          Else
            TxtSupplyAL.Enabled = True
          End If
          
          If Mid(CboVSuppStatus, 1, 1) = "A" Then
            TxtVSuppVisit.Enabled = False
          Else
            TxtVSuppVisit.Enabled = True
          End If
          
         If cboCOPHLTType <> "" Then
            cboCOPHLTType.Enabled = False
            cboCOPPLNCode.Enabled = True
            cboCOPPLNCode.SetFocus
         End If
End Sub

Private Sub cmdDelete_Click()
Screen.MousePointer = vbHourglass

   ' MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        
        If SSTab1.Tab = 0 Then
           Call DeleteHLT
           CmdAdd.Enabled = True
        ElseIf SSTab1.Tab = 1 Then
          Call DeleteINS
        ElseIf SSTab1.Tab = 2 Then
          Call DeletePLN
        ElseIf SSTab1.Tab = 3 Then
          Call DeletePRM
        ElseIf SSTab1.Tab = 4 Then
          Call DeleteANN
        ElseIf SSTab1.Tab = 5 Then
          Call DeleteVisit
        ElseIf SSTab1.Tab = 6 Then
          Call DeleteMCO
        ElseIf SSTab1.Tab = 7 Then
          Call DeleteCoMBM
        End If
  '  MediConnCISDB.Close
  '  Set MediConnCISDB = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
  '  MediConnCISWS.Close
  '  Set MediConnCISWS = Nothing
   ' MediConnCISPayment.Close
   ' Set MediConnCISPayment = Nothing
Screen.MousePointer = Default
End Sub

Private Sub cmdSearch_Click()
Screen.MousePointer = vbHourglass


  '  MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
  '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"

        Select Case SSTab1.Tab
            Case 0
                Call SearchHLT
                Call ClearForm(Me)
                txtHLTcode.Enabled = True
                txtHLTdesc.Enabled = True
                txtHLTcode.SetFocus
            Case 1
                Call SearchINS
                txtINScode.Enabled = True
                txtINSdesc.Enabled = True
                txtINScode.SetFocus
            Case 2
                Call SearchPLN
                cboPLNHLTCode.Enabled = True
                txtGrpCom.Enabled = True
                cboPLNHLTCode.SetFocus
            Case 3
                Call SearchPRM
                cboPmINSCode.Enabled = True
                cboPmAgeCode.Enabled = True
                cboPmPAYCode.Enabled = True
                cboPmPLNCode.Enabled = True
                cboPmHLTCode.Enabled = True
                cboPmHLTCode.SetFocus
            Case 4
                Call SearchANN
                CboPlnCode.Enabled = True
                CboHltCode.Enabled = True
                CboInsType.Enabled = True
                Frame4.Enabled = True
                TxtSupplyAL.Enabled = False
                txtALEffDate.Enabled = False
                CboPlnCode.SetFocus
            Case 5
                Call SearchVisit
                CboVPLNCode.Enabled = True
                CboVHltCode.Enabled = True
                CboVINSCode.Enabled = True
                Frame4.Enabled = True
                TxtVSuppVisit.Enabled = False
                TxtVLimitEffDate.Enabled = False
                CboVPLNCode.SetFocus
            Case 6
                Call SearchMCO
                CboMCOHlt.Enabled = True
                txtEffDate.Enabled = True
                CboMCOHlt.SetFocus
            Case 7
                Call SearchCoMBM
                cboCOPHLTType.Enabled = True
                cboCOPPLNCode.Enabled = True
                txtCOPAmt.Enabled = False
           End Select
       
            CmdAdd.Enabled = True
            
       ' MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
      ''  MediConnCISWS.Close
     '   Set MediConnCISWS = Nothing
     '   MediConnCISPayment.Close
     '   Set MediConnCISPayment = Nothing
Screen.MousePointer = Default
End Sub

Private Sub CmdRefresh_Click()
Screen.MousePointer = vbHourglass
    Call ClearForm(Me)
    Call EntriesEnable(True)
    Label5.Visible = True
    
  '  MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        Select Case SSTab1.Tab
            'Case 0
                'Call HLTListing
            'Case 1
                'Call INSListing
            'Case 2
                'Call PLNListing
            Case 3
                Call PRMComboListing
            Case 4
                Call AnnualComboListing
            Case 5
                Call VisitComboListing
            Case 6
                Call MCOComboListing
            Case 7
                Call CoMBMComboListing
        End Select
  '  MediConnCISDB.Close
  '  Set MediConnCISDB = Nothing
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  MediConnCISWS.Close
  '  Set MediConnCISWS = Nothing
  '  MediConnCISPayment.Close
  '  Set MediConnCISPayment = Nothing
    
    txtTotalRecord = 0
Screen.MousePointer = Default
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    CmdAdd.Enabled = True
    lstHLTList.ListItems.Clear
    lstINSList.ListItems.Clear
    lstPlanList.ListItems.Clear
    lstPmPRMList.ListItems.Clear
    lstAnnualList.ListItems.Clear
    lstVisitLimit.ListItems.Clear
    lstMCOList.ListItems.Clear
    lstCoPaymentMBMList.ListItems.Clear
'    cboSuppMCOStatus.Text = "C" & " - " & "Separate Limit for Supplementary"
    Frame4.Enabled = True
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
   ' MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
   '' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        Select Case SSTab1.Tab
            Case 0
                Call PrintHLT
            Case 1
                Call PrintINS
            Case 2
                Call PrintPLN
            Case 3
                Call PrintPRM
            Case 4
                Call PrintANN
            Case 5
                Call PrintVisit
            Case 6
                Call PrintMCO
            Case 7
                Call PrintCoPaymentMBM
        End Select
    
   ' MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' MediConnCISWS.Close
   ' Set MediConnCISWS = Nothing
   ' MediConnCISPayment.Close
   ' Set MediConnCISPayment = Nothing
    
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    Screen.MousePointer = Default
End Sub

Private Sub CmdSave_Click()
    Screen.MousePointer = vbHourglass
    
  '  MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
  ''  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
        Select Case SSTab1.Tab
            Case 0
                Call SaveHLT
            Case 1
                Call SaveINS
            Case 2
                Call SavePLN
            Case 3
                Call SavePRM
            Case 4
                Call SaveANN
            Case 5
                Call SaveVisit
            Case 6
                Call SaveMCO
            Case 7
                Call SaveCoMBM
        End Select
    
   ' MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' MediConnCISWS.Close
   ' Set MediConnCISWS = Nothing
   ' MediConnCISPayment.Close
   ' Set MediConnCISPayment = Nothing
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    Screen.MousePointer = Default
End Sub

Private Sub CmdExit_Click()

    Unload Me
    
End Sub
'#################### End Add, Edit, Delete, Search, Save, Exit ####################'

'########################### Start Enteries Enable ###########################'
Private Sub EntriesEnable(status As Boolean)
   
    Dim ctl As Control

    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.Enabled = status
        End If
    Next ctl
   
    
End Sub
'########################### End Entries Enable ###########################'

'########################### Start Health ###########################'
Private Sub DeleteHLT()
    Dim rsDeleteHLT As New ADODB.Recordset
    Dim rsCheckHLT As New ADODB.Recordset
    Dim rsCheckAGE As New ADODB.Recordset
    Dim rsCheckBNF As New ADODB.Recordset
    Dim rsCheckMCO As New ADODB.Recordset
    Dim rsCheckPLN As New ADODB.Recordset
    Dim rsCheckPRM As New ADODB.Recordset
    Dim rsCheckPRV As New ADODB.Recordset
    Dim StrSql, sqlHLT, sqlAGE, sqlBNF, sqlMBM, sqlMCO, sqlPRM, sqlPRV, sqlPLN As String
    Dim DeleteRecord
    EditFlag = False
        
    If txtHLTcode = "" Then
       MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
       'txtHLTcode.SetFocus
    ElseIf lstHLTList.ListItems.Count = 0 Then
       MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    sqlMBM = "Select MBMNumber From MBM where HLTCode = '" & RemStr(txtHLTcode) & "'" '(ChkStr(lstHLTList.SelectedItem.Text) & "'"
    sqlPLN = "Select PLNCode From PLN where HLTCode = '" & RemStr(txtHLTcode) & "'" 'ChkStr(lstHLTList.SelectedItem.Text) & "'"
    sqlAGE = "Select AgeCode From AGE where HLTCode = '" & RemStr(txtHLTcode) & "'" 'ChkStr(lstHLTList.SelectedItem.Text) & "'"
'    sqlBNF = "Select BNFCode From BNF where HLTCode = '" & RemStr(txtHLTcode) & "'" 'ChkStr(lstHLTList.SelectedItem.Text) & "'"
    sqlMCO = "Select MCOIndex From MCO where HLTCode = '" & RemStr(txtHLTcode) & "'" 'ChkStr(lstHLTList.SelectedItem.Text) & "'"
    sqlPRM = "Select PRMCode From PRM where HLTCode = '" & RemStr(txtHLTcode) & "'" 'ChkStr(lstHLTList.SelectedItem.Text) & "'"
'    sqlPRV = "Select PRVIndex From PRV where HLTCode = '" & RemStr(txtHLTcode) & "'" 'ChkStr(lstHLTList.SelectedItem.Text) & "'"
    
    
    rsCheckHLT.MaxRecords = 1
    rsCheckHLT.Open sqlMBM, MediConnCISDB, adOpenKeyset, adLockOptimistic

    rsCheckPLN.MaxRecords = 1
    rsCheckPLN.Open sqlPLN, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    rsCheckAGE.MaxRecords = 1
    rsCheckAGE.Open sqlAGE, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
'    rsCheckBNF.MaxRecords = 1
'    rsCheckBNF.Open sqlBNF, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    rsCheckMCO.MaxRecords = 1
    rsCheckMCO.Open sqlMCO, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    rsCheckPRM.MaxRecords = 1
    rsCheckPRM.Open sqlPRM, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
'    rsCheckPRV.MaxRecords = 1
'    rsCheckPRV.Open sqlPRV, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    
    If rsCheckHLT.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in Member table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    ElseIf rsCheckPLN.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in Plan table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    ElseIf rsCheckAGE.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in Age table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
'    ElseIf rsCheckBNF.RecordCount Then
'        MsgBox "This record cannot be deleted because it is used in Benefit table!", vbCritical, "Error in deleting record..."
'        Call ClearForm(Me)
'        CmdDelete.Enabled = False
'        CmdEdit.Enabled = False
    ElseIf rsCheckMCO.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in MCO table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    ElseIf rsCheckPRM.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in Premium table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
       CmdDelete.Enabled = False
        CmdEdit.Enabled = False
'    ElseIf rsCheckPRV.RecordCount Then
'        MsgBox "This record cannot be deleted because it is used in PRV table!", vbCritical, "Error in deleting record..."
'        Call ClearForm(Me)
'        CmdDelete.Enabled = False
'        CmdEdit.Enabled = False
    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes And rsCheckHLT.RecordCount = 0 Then
             StrSql = "Delete From HLT Where HLTCode = '" & ChkStr(lstHLTList.SelectedItem.Text) & "'"
             rsDeleteHLT.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call ClearForm(Me)
             Call HLTListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             txtHLTcode.SetFocus
        Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
        End If
    End If
    End If
    Exit Sub
End Sub
    
Private Sub SearchHLT()
Dim strAppend As String
    
  Screen.MousePointer = vbHourglass
  strAppend = ""
  
  If Len(Trim(txtHLTcode)) Then
  strAppend = strAppend + " And HLTCode = '" & Trim(Mid(txtHLTcode, 1, 2)) & "'"
  'HLTCode like '%" & Mid(ChkStr(txtHLTcode), 1, 2) & "%'"
  End If
    
  HLTCriteria = strAppend
  Call HLTListing(strAppend)
  Screen.MousePointer = Default
  
End Sub

Private Sub PrintHLT()
    RptHltList.Show
End Sub

Private Sub SaveHLT()
   On Error Resume Next
   Dim RecordSaved
   Dim HLT As New ADODB.Recordset
   Dim SaveHLT As New ADODB.Recordset
   Dim StrSql As String
   
    If txtHLTcode = "" Then
        MsgBox "Please Enter Health Type", vbOKOnly + vbCritical, DialogBoxTitle
        txtHLTcode.SetFocus
    ElseIf txtHLTdesc = "" Then
        MsgBox "Please Enter Health Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtHLTdesc.SetFocus
    Else
        StrSql = "Select * From HLT where HLTCode =  '" & Trim(txtHLTcode) & "'"
        HLT.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If HLT.RecordCount > 0 And EditFlag = False Then
           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           HLT.Close
           txtHLTcode.SetFocus
           txtHLTcode.SelStart = 0
           txtHLTcode.SelLength = Len(txtHLTcode)
           Set HLT = Nothing
           Exit Sub
        Else
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           If RecordSaved = vbYes Then
              If EditFlag = False Then
                 SaveHLT.Open "HLT", MediConnCISDB, adOpenKeyset, adLockOptimistic
                 SaveHLT.AddNew
              Else
                 StrSql = "Select * From HLT Where HLTCode = '" & ChkStr(txtHLTcode) & "'"
                 SaveHLT.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
              End If
              
              With SaveHLT
                !HLTCode = ChkStr(txtHLTcode)
                !HLTDesc = ChkStr(txtHLTdesc)
                !HLTLastUpdateUser = Logon
                !HLTLastUpdateDate = Date
                
              End With
              SaveHLT.update
            
             Call ClearForm(Me)
             Call HLTListing(HLTCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             HLT.Close
             Set SaveHLT = Nothing
        End If
        EditFlag = False
     End If
   End If
End Sub

Private Sub HLTListing(Optional strAppend As String)
    Dim rslistHLT As New ADODB.Recordset
    Dim rsListHLT2 As New ADODB.Recordset
    Dim HLTMaster As ListItem
    Dim sqlHLT As String
    Dim TotalHLTCount As Integer
         
    
    TotalHLTCount = 0
    lstHLTList.ListItems.Clear
    sqlHLT = "select HLTCode, HLTDesc From HLT where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sqlHLT = sqlHLT + strAppend
    End If
    
    rslistHLT.Open sqlHLT, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rslistHLT.RecordCount = 0 Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    
    Do Until rslistHLT.EOF
        With rslistHLT
            Set HLTMaster = lstHLTList.ListItems.Add(, , !HLTCode)
                HLTMaster.SubItems(1) = !HLTDesc
            .MoveNext
        End With
        TotalHLTCount = TotalHLTCount + 1
        txtTotalRecord = TotalHLTCount
    Loop
    End If
    rslistHLT.Close
    Set rslistHLT = Nothing
        
End Sub


Private Sub lstHLTList_Click()
Dim StrSql As String
    
    If lstHLTList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtHLTcode = lstHLTList.SelectedItem.Text
        txtHLTcode.Enabled = False
        txtHLTdesc = lstHLTList.SelectedItem.SubItems(1)
        txtHLTdesc.Enabled = False
            
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstHLTList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim StrSql As String
    
    If lstHLTList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtHLTcode = lstHLTList.SelectedItem.Text
        txtHLTcode.Enabled = False
        txtHLTdesc = lstHLTList.SelectedItem.SubItems(1)
        txtHLTdesc.Enabled = False
            
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstHLTList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim StrSql As String
    
    If lstHLTList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtHLTcode = lstHLTList.SelectedItem.Text
        txtHLTcode.Enabled = False
        txtHLTdesc = lstHLTList.SelectedItem.SubItems(1)
        txtHLTdesc.Enabled = False
            
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstHLTList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstHLTList.SortKey = ColumnHeader.Index - 1
    lstHLTList.Sorted = True
    If lstHLTList.SortOrder = lvwAscending Then
       lstHLTList.SortOrder = lvwDescending
    Else
       lstHLTList.SortOrder = lvwAscending
    End If
    
    txtHLTcode.Enabled = False
End Sub
'########################### End Health ###########################'

'####################### Start Insurance Type #######################'
Private Sub DeleteINS()
    
    Dim rsDeleteINS As New ADODB.Recordset
    Dim rsCheckINS As New ADODB.Recordset
    Dim rsCheckBNF As New ADODB.Recordset
    Dim rsCheckPRM As New ADODB.Recordset
    Dim rsCheckMBMSEQ As New ADODB.Recordset
    Dim rsCheckPRV As New ADODB.Recordset
    Dim rsCheckMCO As New ADODB.Recordset
    Dim StrSql, sqlINS, sqlBNF, sqlPRM, sqlMBMSEQ, sqlPRV, sqlMCO As String
    Dim DeleteRecord
    EditFlag = False
        
    If txtINScode = "" Then
       MsgBox "Please Enter Insured Type", vbOKOnly + vbCritical, DialogBoxTitle
       'txtINScode.SetFocus
    ElseIf lstINSList.ListItems.Count = 0 Then
       MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    sqlINS = "Select INSCode From MBM where INSCode = '" & RemStr(txtINScode) & "'" 'ChkStr(lstINSList.SelectedItem.Text) & "'"
    sqlBNF = "Select BNFCode From BNF where INSCode = '" & RemStr(txtINScode) & "'" 'ChkStr(lstINSList.SelectedItem.Text) & "'"
    sqlPRM = "Select PRMCode From PRM where INSCode = '" & RemStr(txtINScode) & "'" 'ChkStr(lstINSList.SelectedItem.Text) & "'"
    sqlMBMSEQ = "Select SEQCode From MBMSEQ where INSCode = '" & RemStr(txtINScode) & "'" 'ChkStr(lstINSList.SelectedItem.Text) & "'"
    sqlPRV = "Select PRVIndex From PRV where INSCode = '" & RemStr(txtINScode) & "'" 'ChkStr(lstINSList.SelectedItem.Text) & "'"
    sqlMCO = "Select MCOCode From MCO where INSCode = '" & ChkStr(lstINSList.SelectedItem.Text) & "'"
    
    rsCheckBNF.MaxRecords = 1
    rsCheckBNF.Open sqlBNF, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
'    rsCheckMCO.MaxRecords = 1
'    rsCheckMCO.Open sqlMCO, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    rsCheckINS.MaxRecords = 1
    rsCheckINS.Open sqlINS, MediConnCISDB, adOpenKeyset, adLockOptimistic
       
    rsCheckPRM.MaxRecords = 1
    rsCheckPRM.Open sqlPRM, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
'    rsCheckMBMSEQ.MaxRecords = 1
'    rsCheckMBMSEQ.Open sqlMBMSEQ, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
'    rsCheckPRV.MaxRecords = 1
'    rsCheckPRV.Open sqlPRV, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If rsCheckBNF.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in Benefit table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
'    ElseIf rsCheckMCO.RecordCount Then
'        MsgBox "This record cannot be deleted because it is used in MCO table!", vbCritical, "Error in deleting record..."
'        Call ClearForm(Me)
'        CmdDelete.Enabled = False
'        CmdEdit.Enabled = False
    ElseIf rsCheckINS.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in MBM table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
'    ElseIf rsCheckMBMSEQ.RecordCount Then
'        MsgBox "This record cannot be deleted because it is used in MBMSEQ table!", vbCritical, "Error in deleting record..."
'        Call ClearForm(Me)
'        CmdDelete.Enabled = False
'        CmdEdit.Enabled = False
    ElseIf rsCheckPRM.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in PRM table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
'    ElseIf rsCheckPRV.RecordCount Then
'        MsgBox "This record cannot be deleted because it is used in PRV table!", vbCritical, "Error in deleting record..."
'        Call ClearForm(Me)
'        CmdDelete.Enabled = False
'        CmdEdit.Enabled = False
    ElseIf rsCheckINS.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in Member table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
'            And rsCheckINS.RecordCount = 0 Then
             StrSql = "Delete From INS Where INSCode = '" & ChkStr(lstINSList.SelectedItem.Text) & "'"
             rsDeleteINS.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call ClearForm(Me)
             Call INSListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             txtINScode.Enabled = True
             txtINScode.SetFocus
        Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
        End If
    End If
 End If
    Exit Sub

End Sub
    
Private Sub SearchINS()
    Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(txtINScode)) Then
        strAppend = strAppend + " And INSCode = '" & ChkStr(txtINScode) & "'"
    End If
    
    INSCriteria = strAppend
    Call INSListing(strAppend)
    Screen.MousePointer = Default
    
End Sub

Private Sub SaveINS()
Dim RecordSaved
Dim INS As New ADODB.Recordset
Dim SaveINS As New ADODB.Recordset
Dim StrSql As String
    
    If txtINScode = "" Then
        MsgBox "Please Enter Insured Type", vbOKOnly + vbCritical, DialogBoxTitle
        txtINScode.SetFocus
    ElseIf txtINSdesc = "" Then
        MsgBox "Please Enter Insured Type Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtINSdesc.SetFocus
    Else
        StrSql = "Select * From INS where INSCode = '" & Trim(txtINScode) & "'"
        INS.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If INS.RecordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            INS.Close
            txtINScode.SetFocus
            txtINScode.SelStart = 0
            txtINScode.SelLength = Len(txtINScode)
            Set INS = Nothing
            Set SaveINS = Nothing
            Exit Sub
            
        Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                   SaveINS.Open "INS", MediConnCISDB, adOpenKeyset, adLockOptimistic
                   SaveINS.AddNew
                Else
                   StrSql = "Select * From INS Where INSCode = '" & ChkStr(txtINScode) & "'"
                   SaveINS.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                End If
                With SaveINS
                     !INSCode = ChkStr(txtINScode)
                     !InsDesc = ChkStr(txtINSdesc)
                     !INSLastUpdateUser = Logon
                     !INSLastUpdateDate = Date
                     
                     
                End With
                SaveINS.update
                Call ClearForm(Me)
                Call INSListing(INSCriteria)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
        
                SaveINS.Close
            End If
            EditFlag = False
      End If
  End If
End Sub

Private Sub PrintINS()
    RptInsList.Show
End Sub

Private Sub INSListing(Optional strAppend As String)
Dim rsListINS As New ADODB.Recordset
Dim INSMaster As ListItem
Dim sqlINS As String
Dim TotalINSCount As Integer
        
  
    TotalINSCount = 0
    lstINSList.ListItems.Clear
    sqlINS = "select INSCode, INSDesc From INS where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sqlINS = sqlINS + strAppend
    End If
    
    rsListINS.Open sqlINS, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If rsListINS.RecordCount = 0 Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
    Do Until rsListINS.EOF
        With rsListINS
            Set INSMaster = lstINSList.ListItems.Add(, , !INSCode)
                INSMaster.SubItems(1) = !InsDesc
            .MoveNext
        End With
        TotalINSCount = TotalINSCount + 1
        txtTotalRecord = TotalINSCount
    Loop
    End If
    rsListINS.Close
    Set rsListINS = Nothing
    
End Sub

Private Sub lstINSList_Click()
Dim StrSql As String
    
    If lstINSList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        txtINScode = lstINSList.SelectedItem.Text
        txtINScode.Enabled = False
        txtINSdesc = lstINSList.SelectedItem.SubItems(1)
        txtINSdesc.Enabled = False
        Call EntriesEnable(False)
        CmdDelete.Enabled = True
        CmdSave.Enabled = False
    End If
End Sub

Private Sub lstINSList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim StrSql As String
    
    If lstINSList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        txtINScode = lstINSList.SelectedItem.Text
        txtINScode.Enabled = False
        txtINSdesc = lstINSList.SelectedItem.SubItems(1)
        txtINSdesc.Enabled = False
        Call EntriesEnable(False)
        CmdDelete.Enabled = True
        CmdSave.Enabled = False
    End If
End Sub

Private Sub lstINSList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim StrSql As String
    
    If lstINSList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        txtINScode = lstINSList.SelectedItem.Text
        txtINScode.Enabled = False
        txtINSdesc = lstINSList.SelectedItem.SubItems(1)
        txtINSdesc.Enabled = False
        Call EntriesEnable(False)
        CmdDelete.Enabled = True
        CmdSave.Enabled = False
    End If
End Sub

Private Sub lstINSList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstINSList.SortKey = ColumnHeader.Index - 1
    lstINSList.Sorted = True
    If lstINSList.SortOrder = lvwAscending Then
       lstINSList.SortOrder = lvwDescending
    Else
       lstINSList.SortOrder = lvwAscending
    End If
    
    txtINScode.Enabled = False
End Sub
'######################## End Insurance Type ########################'

'########################## Start Plan Code ##########################'
Private Sub DeletePLN()
    'On Error GoTo DeleleErr
    
    Dim rsDeletePLN As New ADODB.Recordset
    Dim rsCheckMBM As New ADODB.Recordset
    Dim StrSql As String
    Dim sqlMBM As String
    Dim DeleteRecord
    EditFlag = False
    
    
    'If TxtPLNCode = "" Then
    '   MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
      ' TxtPLNCode.SetFocus
    If lstPlanList.ListItems.Count = 0 Then
       MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    sqlMBM = "Select MBMNumber From MBMCLNOthers where PLNType = '" & RemStr(lstPlanList.SelectedItem.SubItems(1)) & "'"
    
    rsCheckMBM.MaxRecords = 1
    rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rsCheckMBM.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in Membership table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    
        If DeleteRecord = vbYes Then
'        And rsCheckMBM.RecordCount = 0 Then
             StrSql = "Delete From PLN Where PLNCode = '" & ChkStr(lstPlanList.SelectedItem.SubItems(1)) & "'"
             rsDeletePLN.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             Call ClearForm(Me)
             Call PLNListing(PLNCriteria)
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             CmdAdd.Enabled = True
             CmdSave.Enabled = True
        Else
             CmdDelete.Enabled = False
             CmdEdit.Enabled = False
        End If
    End If
    Exit Sub
  End If
End Sub
    
Private Sub SearchPLN()
    Dim strAppend As String
        
    Screen.MousePointer = vbHourglass
    
    strAppend = ""
        
    If Len(Trim(txtPLN1)) Then
        strAppend = strAppend + " And PLNCode like '%" & ChkStr(txtPLN1) & "%'"
    End If
        
    If Len(Trim(cboPLNPayor)) Then
        strAppend = strAppend + " And PAYCode = '" & Mid(cboPLNPayor, 1, 2) & "'"
    End If
    
    If Len(Trim(cboPLNPayor2)) Then
        strAppend = strAppend + " And HISPAYCode = '" & Mid(cboPLNPayor2, 1, 2) & "'"
    End If
   
    If Len(Trim(cboPLNHLTCode)) Then
        strAppend = strAppend + " And HLTCode = '" & Mid(cboPLNHLTCode, 1, 1) & "'"
    End If
   
    If Len(Trim(txtGrpCom)) Then
        strAppend = strAppend + " And CoGRPCode = '" & Mid(txtGrpCom, 1, 4) & "'"
    End If
    
    If Len(Trim(cboPLNCLNContract)) Then
        strAppend = strAppend + " And ClnContract = '" & Trim(cboPLNCLNContract) & "'"
    End If
    
    If Len(Trim(cboPLNAnnLimit)) Then
        strAppend = strAppend + " And ANNLmtInd = '" & Mid(cboPLNAnnLimit, 1, 1) & "'"
    End If
    
    If Len(Trim(cboPLNVisitLimit)) Then
        strAppend = strAppend + " And VISLmtInd = '" & Mid(cboPLNVisitLimit, 1, 1) & "'"
    End If
    
    If Len(Trim(cboPLNPremium)) Then
        strAppend = strAppend + " And PRMInd = '" & Mid(cboPLNPremium, 1, 1) & "'"
    End If
    
    If Len(Trim(cboPLNMCO)) Then
        strAppend = strAppend + " And MCOInd = '" & Mid(cboPLNMCO, 1, 1) & "'"
    End If
        
    If Len(Trim(CboPDentalStatus)) Then
        strAppend = strAppend + " And PDentalStatus = '" & Mid(CboPDentalStatus, 1, 1) & "'"
    End If
    
    If Len(Trim(CboPDentalLimit)) Then
        strAppend = strAppend + " And PDentalLimit = '" & Mid(CboPDentalLimit, 1, 1) & "'"
    End If
    
    If Len(Trim(CboPSpectStatus)) Then
        strAppend = strAppend + " And PSpecStatus = '" & Mid(CboPSpectStatus, 1, 1) & "'"
    End If
    
    If Len(Trim(CboPSpectLimit)) Then
        strAppend = strAppend + " And PSpecLimit = '" & Mid(CboPSpectLimit, 1, 1) & "'"
    End If
    
    If Len(Trim(CboPMEPStatus)) Then
        strAppend = strAppend + " And PMEPStatus = '" & Mid(CboPMEPStatus, 1, 1) & "'"
    End If
    
    If Len(Trim(CboSDentalStatus)) Then
        strAppend = strAppend + " And SDentalStatus = '" & Mid(CboSDentalStatus, 1, 1) & "'"
    End If
    
    If Len(Trim(CboSDentalLimit)) Then
        strAppend = strAppend + " And SDentalLimit = '" & Mid(CboSDentalLimit, 1, 1) & "'"
    End If
    
    If Len(Trim(CboSSpecStatus)) Then
        strAppend = strAppend + " And SSpecStatus = '" & Mid(CboSSpecStatus, 1, 1) & "'"
    End If
    
    If Len(Trim(CboSSpecLimit)) Then
        strAppend = strAppend + " And SSpecLimit = '" & Mid(CboSSpecLimit, 1, 1) & "'"
    End If
    
    If Len(Trim(CboSMEPStatus)) Then
        strAppend = strAppend + " And SMEPStatus = '" & Mid(CboSMEPStatus, 1, 1) & "'"
    End If
    
    If Len(Trim(CboSAnnLimit)) Then
        strAppend = strAppend + " And ANNLmtIndSupp = '" & Mid(CboSAnnLimit, 1, 1) & "'"
    End If
    
    If Len(Trim(CboSVisitLimit)) Then
        strAppend = strAppend + " And VISLmtIndSupp = '" & Mid(CboSVisitLimit, 1, 1) & "'"
    End If
    
    PLNCriteria = strAppend
    Call PLNListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub SavePLN()
Dim RecordSaved
Dim CompanyGrp As New ADODB.Recordset
Dim PLN As New ADODB.Recordset
Dim PLN1 As New ADODB.Recordset
Dim PLN2 As New ADODB.Recordset
Dim PLN3 As New ADODB.Recordset
Dim PLN4 As New ADODB.Recordset
Dim PLN5 As New ADODB.Recordset
Dim PLN6 As New ADODB.Recordset
Dim SavePLN As New ADODB.Recordset
Dim StrSql As String
Dim StrSql1 As String
Dim StrSql2 As String
Dim strsql3 As String
Dim strsql4 As String
Dim strsql5 As String
Dim strsql6 As String
Dim ListCount As Integer
Dim ListLoop As Integer
Dim PLNSeq As String
Strgrpcode = ""
    
    If Trim(cboPLNHLTCode) = "" Then
       MsgBox "Please Enter Health Code ", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(cboPLNHLTCode, 1, 1)) = "N" And Trim(cboPLNPayor) = "" Then
       MsgBox "Please Enter Payor Code ", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(cboPLNHLTCode, 1, 1)) = "N" And Trim(CboNoofPlan) = "" Then
       MsgBox "Please Enter No of Plan Code ", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(cboPLNCLNContract) = "" Then
       MsgBox "Please Enter Clinic Contract", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(cboPLNAnnLimit) = "" Then
       MsgBox "Please Enter Principal Annual Limit", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(cboPLNVisitLimit) = "" Then
       MsgBox "Please Enter Principal Visit Limit", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(cboPLNPremium) = "" Then
       MsgBox "Please Enter Premium", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(cboPLNMCO) = "" Then
       MsgBox "Please Enter MCO", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(CboSAnnLimit) = "" Then
       MsgBox "Please Enter Supplementary Annual Limit", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(CboSVisitLimit) = "" Then
       MsgBox "Please Enter Supplementary Visit Limit", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(CboDCStatus) = "" Then
       MsgBox "Please Enter DC Status", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(txtGrpCom) = "" Then
       MsgBox "Please Enter Group Company", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(cboPLNHLTCode, 1, 1)) = "S" And Trim(txtPLN1) = "" Then
       MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(cboOPDependent) = "" Then
       MsgBox "Please Select OP Dependent", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        If Trim(Mid(cboPLNHLTCode, 1, 1)) = "N" Then
        
            StrSql = "Select * From GRPCompany Where GRPCoName = '" & Trim(txtGrpCom) & "'"
            CompanyGrp.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
            If CompanyGrp.RecordCount > 0 Then
                MsgBox "Record Duplicated! Please change the corresponding field", vbCritical, "Error"
                CompanyGrp.Close
                txtGrpCom.SetFocus
                txtGrpCom.SelStart = 0
                txtGrpCom.SelLength = Len(txtGrpCom)
                Set CompanyGrp = Nothing
                Exit Sub
            Else
                CompanyGrp.Close
                Set CompanyGrp = Nothing
        
                RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
                            
                If RecordSaved = vbYes Then
                    
                    ListCount = Trim(CboNoofPlan)
                    For ListLoop = 1 To ListCount
                        
                        SavePLN.Open "PLN", MediConnCISDB, adOpenKeyset, adLockOptimistic
                        SavePLN.AddNew
                        PLNSeq = GeneratePLNSeq(Mid(txtGrpCom, 1, 1))
                        
                        With SavePLN
                            !PLNCode = ChkStr(PLNSeq)
                            
                            If ListLoop = "1" Then
                                !PlnDesc = ChkStr(txtDesc1)
                            ElseIf ListLoop = "2" Then
                                !PlnDesc = ChkStr(txtDesc2)
                            ElseIf ListLoop = "3" Then
                                !PlnDesc = ChkStr(txtDesc3)
                            ElseIf ListLoop = "4" Then
                                !PlnDesc = ChkStr(txtDesc4)
                            ElseIf ListLoop = "5" Then
                                !PlnDesc = ChkStr(txtDesc5)
                            ElseIf ListLoop = "6" Then
                                !PlnDesc = ChkStr(txtDesc6)
                            End If
                            
                            !HLTCode = Trim(Mid(cboPLNHLTCode, 1, IIf(InStr(1, cboPLNHLTCode, "-") = 0, 1, InStr(1, cboPLNHLTCode, "-") - 1)))
                            !PAYCode = Trim(Mid(cboPLNPayor, 1, IIf(InStr(1, cboPLNPayor, "-") = 0, 4, InStr(1, cboPLNPayor, "-") - 1)))
                            !HISPAYCode = Trim(Mid(cboPLNPayor2, 1, IIf(InStr(1, cboPLNPayor2, "-") = 0, 4, InStr(1, cboPLNPayor2, "-") - 1)))
                            '!GrpCompany = ChkStr(txtGrpCom)
                            !GRPCompany = Trim(Mid(txtGrpCom, 1, IIf(InStr(1, txtGrpCom, "-") = 0, 4, InStr(1, txtGrpCom, "-") - 1)))
                            Strgrpcode = Trim(Mid(txtGrpCom, 1, IIf(InStr(1, txtGrpCom, "-") = 0, 4, InStr(1, txtGrpCom, "-") - 1)))
                            !AnnLmtIND = Trim(Mid(cboPLNAnnLimit, 1, 1))
                            !VisLmtIND = Trim(Mid(cboPLNVisitLimit, 1, 1))
                            !ANNLmtIndSupp = Trim(Mid(CboSAnnLimit, 1, 1))
                            !VISLmtIndSupp = Trim(Mid(CboSVisitLimit, 1, 1))
                            !PRMInd = Trim(Mid(cboPLNPremium, 1, 1))
                            !MCOInd = Trim(Mid(cboPLNMCO, 1, 1))
                            !ClnContract = Trim(cboPLNCLNContract)
                            !DCStatus = Trim(Mid(CboDCStatus, 1, 1))
                            !PDentalStatus = "N"
                            !PSpecStatus = "N"
                            !PMEPStatus = "N"
                            !PDentalLimit = "N"
                            !PSpecLimit = "N"
                            !SDentalStatus = "N"
                            !SSpecStatus = "N"
                            !SMEPStatus = "N"
                            !SDentalLimit = "N"
                            !SSpecLimit = "N"
                            '!PDentalStatus = Trim(Mid(CboPDentalStatus, 1, 1))
                            '!PSpecStatus = Trim(Mid(CboPSpectStatus, 1, 1))
                            '!PMEPStatus = Trim(Mid(CboPMEPStatus, 1, 1))
                            '!PDentalLimit = Trim(Mid(CboPDentalLimit, 1, 1))
                            '!PSpecLimit = Trim(Mid(CboPSpectLimit, 1, 1))
                            '!SDentalStatus = Trim(Mid(CboSDentalStatus, 1, 1))
                            '!SSpecStatus = Trim(Mid(CboSSpecStatus, 1, 1))
                            '!SMEPStatus = Trim(Mid(CboSMEPStatus, 1, 1))
                            '!SDentalLimit = Trim(Mid(CboSDentalLimit, 1, 1))
                            '!SSpecLimit = Trim(Mid(CboSSpecLimit, 1, 1))
                            !Insured = Mid(cmbInsured, 1, 1)
                            !HRApproval = Mid(cmbHRApproval, 1, 1)
                            
                            !OPFamilyStatus = cmbOPStatus
                            If IsNumeric(Trim(txtPSPLimit)) Then
                                !PrinSPLimit = CDbl(txtPSPLimit)
                            Else
                                !PrinSPLimit = 0
                            End If
                            If IsNumeric(txtPGPLimit) Then
                                !PrinGPLimit = CDbl(txtPGPLimit)
                            Else
                                !PrinGPLimit = 0
                            End If
                            If IsNumeric(txtsSPLimit) Then
                                !DepSPLimit = CDbl(txtsSPLimit)
                            Else
                                !DepSPLimit = 0
                            End If
                            If IsNumeric(txtSGPLimit) Then
                                !DepGPLimit = CDbl(txtSGPLimit)
                            Else
                                !DepGPLimit = 0
                            End If
                            
                            'CHEng 2Feb2017
                            !PLNSpecialRemark = ChkStr(txtSpecialNote.Text)
                            !isopdependent = Mid(cboOPDependent, 1, 1)
                            
                            'PrinSPLimit
                            'PrinGPLimit
                            'DepSPLimit
                            'DepGPLimit
                            
                            !PLNLastUpdateUser = Logon
                            !PLNLastUpdateDate = Date
                            SavePLN.update
                            SavePLN.Close
                            Set SavePLN = Nothing
                        End With
                
                    Next ListLoop
                    
                    Call ClearForm(Me)
                    Call PLNListing(PLNCriteria)
                    
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                
                End If
            
            End If
            
        Else
            
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
                            
            If RecordSaved = vbYes Then
                    
                ListCount = 6
                For ListLoop = 1 To ListCount
                    
                    If ListLoop = 1 Then
                        If Trim(txtPLN1) <> "" Then
                            StrSql1 = " Select PLNCode, HLTCode, PAYCode From PLN where PLNCode = '" & Trim(txtPLN1) & "'" & _
                                      " AND HLTCode = '" & Mid(cboPLNHLTCode, 1, 1) & "' AND PAYCODE = '" & Mid(cboPLNPayor, 1, 2) & "'"
                            PLN1.Open StrSql1, medixmasterconn, adOpenKeyset, adLockOptimistic
                            If PLN1.RecordCount <> 0 Then
                               MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                               PLN1.Close
                               txtPLN1.SetFocus
                               txtPLN1.SelStart = 0
                               txtPLN1.SelLength = Len(txtPLN1)
                               Set PLN1 = Nothing
                               Exit Sub
                            Else
                                SavePLN.Open "PLN", medixmasterconn, adOpenKeyset, adLockOptimistic
                                SavePLN.AddNew
                                SavePLN!PLNCode = ChkStr(txtPLN1)
                                SavePLN!PLNDescription = ChkStr(txtDesc1)
                            End If
                        End If
                    ElseIf ListLoop = 2 Then
                        If Trim(txtPLN2) <> "" Then
                            StrSql2 = " Select PLNCode, HLTCode, PAYCode From PLN where PLNCode = '" & Trim(txtPLN2) & "'" & _
                                      " AND HLTCode = '" & Mid(cboPLNHLTCode, 1, 1) & "' AND PAYCODE = '" & Mid(cboPLNPayor, 1, 2) & "'"
                            PLN2.Open StrSql2, medixmasterconn, adOpenKeyset, adLockOptimistic
                            If PLN2.RecordCount <> 0 Then
                               MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                               PLN2.Close
                               txtPLN2.SetFocus
                               txtPLN2.SelStart = 0
                               txtPLN2.SelLength = Len(txtPLN2)
                               Set PLN2 = Nothing
                               Exit Sub
                            Else
                                SavePLN.Open "PLN", medixmasterconn, adOpenKeyset, adLockOptimistic
                                SavePLN.AddNew
                                SavePLN!PLNCode = ChkStr(txtPLN2)
                                SavePLN!PLNDescription = ChkStr(txtDesc2)
                            End If
                        End If
                    ElseIf ListLoop = 3 Then
                        If Trim(txtPLN3) <> "" Then
                            strsql3 = " Select PLNCode, HLTCode, PAYCode From PLN where PLNCode = '" & Trim(txtPLN3) & "'" & _
                                      " AND HLTCode = '" & Mid(cboPLNHLTCode, 1, 1) & "' AND PAYCODE = '" & Mid(cboPLNPayor, 1, 2) & "'"
                            PLN3.Open strsql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                            If PLN3.RecordCount <> 0 Then
                               MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                               PLN3.Close
                               txtPLN3.SetFocus
                               txtPLN3.SelStart = 0
                               txtPLN3.SelLength = Len(txtPLN3)
                               Set PLN3 = Nothing
                               Exit Sub
                            Else
                                SavePLN.Open "PLN", medixmasterconn, adOpenKeyset, adLockOptimistic
                                SavePLN.AddNew
                                SavePLN!PLNCode = ChkStr(txtPLN3)
                                SavePLN!PLNDescription = ChkStr(txtDesc3)
                            End If
                        End If
                    ElseIf ListLoop = 4 Then
                        If Trim(txtPLN4) <> "" Then
                            strsql4 = " Select PLNCode, HLTCode, PAYCode From PLN where PLNCode = '" & Trim(txtPLN4) & "'" & _
                                      " AND HLTCode = '" & Mid(cboPLNHLTCode, 1, 1) & "' AND PAYCODE = '" & Mid(cboPLNPayor, 1, 2) & "'"
                            PLN4.Open strsql4, medixmasterconn, adOpenKeyset, adLockOptimistic
                            If PLN4.RecordCount <> 0 Then
                               MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                               PLN4.Close
                               txtPLN4.SetFocus
                               txtPLN4.SelStart = 0
                               txtPLN4.SelLength = Len(txtPLN1)
                               Set PLN4 = Nothing
                               Exit Sub
                            Else
                                SavePLN.Open "PLN", medixmasterconn, adOpenKeyset, adLockOptimistic
                                SavePLN.AddNew
                                SavePLN!PLNCode = ChkStr(txtPLN4)
                                SavePLN!PLNDescription = ChkStr(txtDesc4)
                            End If
                        End If
                    ElseIf ListLoop = 5 Then
                        If Trim(txtPLN5) <> "" Then
                            strsql5 = " Select PLNCode, HLTCode, PAYCode From PLN where PLNCode = '" & Trim(txtPLN5) & "'" & _
                                      " AND HLTCode = '" & Mid(cboPLNHLTCode, 1, 1) & "' AND PAYCODE = '" & Mid(cboPLNPayor, 1, 2) & "'"
                            PLN5.Open strsql5, medixmasterconn, adOpenKeyset, adLockOptimistic
                            If PLN5.RecordCount <> 0 Then
                               MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                               PLN5.Close
                               txtPLN5.SetFocus
                               txtPLN5.SelStart = 0
                               txtPLN5.SelLength = Len(txtPLN5)
                               Set PLN5 = Nothing
                               Exit Sub
                            Else
                                SavePLN.Open "PLN", medixmasterconn, adOpenKeyset, adLockOptimistic
                                SavePLN.AddNew
                                SavePLN!PLNCode = ChkStr(txtPLN5)
                                SavePLN!PLNDescription = ChkStr(txtDesc5)
                            End If
                        End If

                    ElseIf ListLoop = 6 Then
                        If Trim(txtPLN6) <> "" Then
                            strsql6 = " Select PLNCode, HLTCode, PAYCode From PLN where PLNCode = '" & Trim(txtPLN6) & "'" & _
                                      " AND HLTCode = '" & Mid(cboPLNHLTCode, 1, 1) & "' AND PAYCODE = '" & Mid(cboPLNPayor, 1, 2) & "'"
                            PLN6.Open strsql6, medixmasterconn, adOpenKeyset, adLockOptimistic
                            If PLN6.RecordCount <> 0 Then
                               MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                               PLN6.Close
                               txtPLN6.SetFocus
                               txtPLN6.SelStart = 0
                               txtPLN6.SelLength = Len(txtPLN6)
                               Set PLN6 = Nothing
                               Exit Sub
                            Else
                                SavePLN.Open "PLN", medixmasterconn, adOpenKeyset, adLockOptimistic
                                SavePLN.AddNew
                                SavePLN!PLNCode = ChkStr(txtPLN6)
                                SavePLN!PLNDescription = ChkStr(txtDesc6)
                            End If
                        End If
                    End If
                            
                    
                    With SavePLN
                        !HLTCode = Trim(Mid(cboPLNHLTCode, 1, IIf(InStr(1, cboPLNHLTCode, "-") = 0, 1, InStr(1, cboPLNHLTCode, "-") - 1)))
                        !PAYCode = Trim(Mid(cboPLNPayor, 1, IIf(InStr(1, cboPLNPayor, "-") = 0, 4, InStr(1, cboPLNPayor, "-") - 1)))
                        !HISPAYCode = Trim(Mid(cboPLNPayor2, 1, IIf(InStr(1, cboPLNPayor2, "-") = 0, 4, InStr(1, cboPLNPayor2, "-") - 1)))
                        !GRPCompany = ChkStr(txtGrpCom)
                        !AnnLmtIND = Trim(Mid(cboPLNAnnLimit, 1, 1))
                        !VisLmtIND = Trim(Mid(cboPLNVisitLimit, 1, 1))
                        !ANNLmtIndSupp = Trim(Mid(CboSAnnLimit, 1, 1))
                        !VISLmtIndSupp = Trim(Mid(CboSVisitLimit, 1, 1))
                        !PRMInd = Trim(Mid(cboPLNPremium, 1, 1))
                        !MCOInd = Trim(Mid(cboPLNMCO, 1, 1))
                        !ClnContract = Trim(cboPLNCLNContract)
                        !DCStatus = Trim(Mid(CboDCStatus, 1, 1))
                        !PDentalStatus = "N"
                        !PSpecStatus = "N"
                        !PMEPStatus = "N"
                        !PDentalLimit = "N"
                        !PSpecLimit = "N"
                        !SDentalStatus = "N"
                        !SSpecStatus = "N"
                        !SMEPStatus = "N"
                        !SDentalLimit = "N"
                        !SSpecLimit = "N"
                        '!PDentalStatus = Trim(Mid(CboPDentalStatus, 1, 1))
                        '!PSpecStatus = Trim(Mid(CboPSpectStatus, 1, 1))
                        '!PMEPStatus = Trim(Mid(CboPMEPStatus, 1, 1))
                        '!PDentalLimit = Trim(Mid(CboPDentalLimit, 1, 1))
                        '!PSpecLimit = Trim(Mid(CboPSpectLimit, 1, 1))
                        '!SDentalStatus = Trim(Mid(CboSDentalStatus, 1, 1))
                        '!SSpecStatus = Trim(Mid(CboSSpecStatus, 1, 1))
                        '!SMEPStatus = Trim(Mid(CboSMEPStatus, 1, 1))
                        '!SDentalLimit = Trim(Mid(CboSDentalLimit, 1, 1))
                        '!SSpecLimit = Trim(Mid(CboSSpecLimit, 1, 1))
                        !Insured = Mid(cmbInsured, 1, 1)
                        !HRApproval = Mid(cmbHRApproval, 1, 1)
                        
                            !OPFamilyStatus = cmbOPStatus
                            If IsNumeric(Trim(txtPSPLimit)) Then
                                !PrinSPLimit = CDbl(txtPSPLimit)
                            Else
                                !PrinSPLimit = 0
                            End If
                            If IsNumeric(txtPGPLimit) Then
                                !PrinGPLimit = CDbl(txtPGPLimit)
                            Else
                                !PrinGPLimit = 0
                            End If
                            If IsNumeric(txtsSPLimit) Then
                                !DepSPLimit = CDbl(txtsSPLimit)
                            Else
                                !DepSPLimit = 0
                            End If
                            If IsNumeric(txtSGPLimit) Then
                                !DepGPLimit = CDbl(txtSGPLimit)
                            Else
                                !DepGPLimit = 0
                            End If
                        !PLNSpecialRemark = ChkStr(txtSpecialNote.Text) 'CHEng 2Feb2017
                        !PLNLastUpdateUser = Logon
                        !PLNLastUpdateDate = Date
                        !isopdependent = Mid(cboOPDependent, 1, 1)
                        SavePLN.update
                        SavePLN.Close
                        Set SavePLN = Nothing
                    End With
                Next ListLoop
                
                Call ClearForm(Me)
                Call PLNListing(PLNCriteria)
                
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                
            End If
                
        End If
            
    End If
    
End Sub

Private Sub PrintPLN()
    RptPlanList.Show
End Sub

Private Sub PLNComboListing()
    Dim PAY As New ADODB.Recordset
    Dim PAY2 As New ADODB.Recordset
    Dim HLT As New ADODB.Recordset
    'Dim GrpCompany As New ADODB.Recordset
    Dim INS As New ADODB.Recordset
    Dim CSMD As New ADODB.Recordset
    
    INS.Open "INS", MediConnCISDB, adOpenKeyset, adLockOptimistic
    PAY.Open "Select PAYCode, PAYCompanyName from PAY", MediConnCISDB, adOpenKeyset, adLockOptimistic
    PAY2.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
    HLT.Open "Select HLTCOde, HLTDesc from HLT", MediConnCISDB, adOpenKeyset, adLockOptimistic
    CSMD.Open "Select CatName from CheckItems where CatCode = 'CSMD'", MediConnCISDB, adOpenKeyset, adLockOptimistic
    'GrpCompany.Open "Select CoGRPCode, CoGRPName from VIEW_GRPCompany", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    CboInsType.Clear
    cboPLNHLTCode.Clear
    cboPLNPayor.Clear
    cboPLNPayor2.Clear
    cboPLNCLNContract.Clear
    
    cboPLNAnnLimit.AddItem "Yes"
    cboPLNAnnLimit.AddItem "No"
    
    cboPLNVisitLimit.AddItem "Yes"
    cboPLNVisitLimit.AddItem "No"

    cboPLNPremium.AddItem "Yes"
    cboPLNPremium.AddItem "No"

    cboPLNMCO.AddItem "Yes"
    cboPLNMCO.AddItem "No"
    
    cboPLNDiag.AddItem "0 - By Plan"
    cboPLNDiag.AddItem "1 - General Diagnosis"
    
    '=========== Principal =============================
    CboPDentalStatus.AddItem "C - Combine"
    CboPDentalStatus.AddItem "S - Stand Alone"
    CboPDentalStatus.AddItem "N - Not Applicable"
    
    CboPDentalLimit.AddItem "0 - Extraction Limit"
    CboPDentalLimit.AddItem "1 - Others Limit"
    CboPDentalLimit.AddItem "2 - No"
    CboPDentalLimit.AddItem "3 - Yes"
    
    CboPSpectStatus.AddItem "C - Combine"
    CboPSpectStatus.AddItem "S - Stand Alone"
    CboPSpectStatus.AddItem "N - Not Applicable"
    
    CboPSpectLimit.AddItem "0 - First Limit"
    CboPSpectLimit.AddItem "1 - Others Limit"
    CboPSpectLimit.AddItem "2 - No"
    CboPSpectLimit.AddItem "3 - Yes"
    
    CboPMEPStatus.AddItem "C - Combine"
    CboPMEPStatus.AddItem "S - Stand Alone"
    CboPMEPStatus.AddItem "N - Not Applicable"
    
    CboNoofPlan.AddItem "1"
    CboNoofPlan.AddItem "2"
    CboNoofPlan.AddItem "3"
    CboNoofPlan.AddItem "4"
    CboNoofPlan.AddItem "5"
    CboNoofPlan.AddItem "6"
    
    '=========== Principal =============================
    
    '=========== Supplementary =============================
    
    CboSAnnLimit.AddItem "Yes"
    CboSAnnLimit.AddItem "No"
    
    CboSVisitLimit.AddItem "Yes"
    CboSVisitLimit.AddItem "No"
    
    CboSDentalStatus.AddItem "C - Combine"
    CboSDentalStatus.AddItem "S - Stand Alone"
    CboSDentalStatus.AddItem "N - Not Applicable"
    
    CboSDentalLimit.AddItem "0 - No Extraction Limit"
    CboSDentalLimit.AddItem "1 - No Others Limit"
    CboSDentalLimit.AddItem "2 - No"
    CboSDentalLimit.AddItem "3 - Yes"
    
    CboSSpecStatus.AddItem "C - Combine"
    CboSSpecStatus.AddItem "S - Stand Alone"
    CboSSpecStatus.AddItem "N - Not Applicable"
    
    CboSSpecLimit.AddItem "0 - No Extraction Limit"
    CboSSpecLimit.AddItem "1 - No Others Limit"
    CboSSpecLimit.AddItem "2 - No"
    CboSSpecLimit.AddItem "3 - Yes"
    
    CboSMEPStatus.AddItem "C - Combine"
    CboSMEPStatus.AddItem "S - Stand Alone"
    CboSMEPStatus.AddItem "N - Not Applicable"
    
    CboDCStatus.AddItem "N - No"
    CboDCStatus.AddItem "Y - Yes"
    
    cboOPDependent.AddItem "N - No"
    cboOPDependent.AddItem "Y - Yes"
    '=========== Supplementary =============================
    
    Do Until HLT.EOF
       cboPLNHLTCode.AddItem HLT!HLTCode & " - " & HLT!HLTDesc
       HLT.MoveNext
    Loop
    
    Do Until INS.EOF
       CboInsType.AddItem INS!INSCode & " - " & INS!InsDesc
       INS.MoveNext
    Loop
  
    Do Until PAY.EOF
       cboPLNPayor.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
       PAY.MoveNext
    Loop
    
    Do Until PAY2.EOF
       cboPLNPayor2.AddItem PAY2!PAYCode & " - " & PAY2!PAYCompanyName
       PAY2.MoveNext
    Loop
    
    'Do Until GrpCompany.EOF
       'txtGrpCom.AddItem Trim(GrpCompany!CoGRPCode) & " - " & Trim(GrpCompany!CoGRPName)
       'GrpCompany.MoveNext
    'Loop
    
    
    Do Until CSMD.EOF
        cboPLNCLNContract.AddItem CSMD!CatName
        CSMD.MoveNext
    Loop
    
    INS.Close
    Set INS = Nothing
    HLT.Close
    Set HLT = Nothing
    PAY.Close
    Set PAY = Nothing
    PAY2.Close
    Set PAY2 = Nothing
    CSMD.Close
    Set CSMD = Nothing
    'GrpCompany.Close
    'Set GrpCompany = Nothing
    
End Sub

Private Sub PLNListing(Optional strAppend As String)
    Dim rsListPlan As New ADODB.Recordset
    Dim TotalPLNCount As Integer
    Dim PLNMaster As ListItem
    Dim sqlPLN As String
    
    TotalPLNCount = 0
    lstPlanList.ListItems.Clear
    
    If Len(Strgrpcode) Then
        strAppend = strAppend + " And CoGRPCode= '" & Trim(Strgrpcode) & "'"
    End If
    
    sqlPLN = "select * From VIEW_PLNListing Where 0 = 0 "
    If Len(Trim(strAppend)) Then
        sqlPLN = sqlPLN + strAppend
    End If
    rsListPlan.Open sqlPLN, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rsListPlan.RecordCount = 0 Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
    Do Until rsListPlan.EOF
         Set PLNMaster = lstPlanList.ListItems.Add(, , rsListPlan!HLTCode)
             PLNMaster.SubItems(1) = rsListPlan!PLNCode
             
             If rsListPlan!PlnDesc <> "" Then
                PLNMaster.SubItems(2) = rsListPlan!PlnDesc
             End If
             PLNMaster.SubItems(3) = IIf(Trim(rsListPlan!PAYCode) <> "", Trim(rsListPlan!PAYCode), "")
             PLNMaster.SubItems(4) = IIf(Trim(rsListPlan!HISPAYCode) <> "", Trim(rsListPlan!HISPAYCode), "")
             PLNMaster.SubItems(5) = IIf(Trim(rsListPlan!CoGRPCode) <> "", rsListPlan!CoGRPCode, "")
             PLNMaster.SubItems(6) = IIf(Trim(rsListPlan!CoGRPName) <> "", rsListPlan!CoGRPName, "")
             PLNMaster.SubItems(7) = IIf(Trim(rsListPlan!ClnContract) <> "", rsListPlan!ClnContract, "")
             PLNMaster.SubItems(8) = IIf(Trim(rsListPlan!Insured) <> "", rsListPlan!Insured, "")
             PLNMaster.SubItems(9) = IIf(Trim(rsListPlan!HRApproval) <> "", rsListPlan!HRApproval, "")
             
             PLNMaster.SubItems(10) = IIf(IsNumeric(rsListPlan!PrinSPLimit), rsListPlan!PrinSPLimit, 0)
             PLNMaster.SubItems(11) = IIf(IsNumeric(rsListPlan!PrinGPLimit), rsListPlan!PrinGPLimit, 0)
             PLNMaster.SubItems(12) = IIf(IsNumeric(rsListPlan!DepSPLimit), rsListPlan!DepSPLimit, 0)
             PLNMaster.SubItems(13) = IIf(IsNumeric(rsListPlan!DepGPLimit), rsListPlan!DepGPLimit, 0)
             PLNMaster.SubItems(14) = IIf(Len(Trim(rsListPlan!OPFamilyStatus)), rsListPlan!OPFamilyStatus, "")
             PLNMaster.SubItems(15) = IIf(Len(Trim(rsListPlan!PLNSpecialRemark)), rsListPlan!PLNSpecialRemark, "") 'CHEng 2Feb2017
             PLNMaster.SubItems(16) = IIf(Len(Trim(rsListPlan!isopdependent)), rsListPlan!isopdependent, "")
             rsListPlan.MoveNext
             TotalPLNCount = TotalPLNCount + 1
             txtTotalRecord = TotalPLNCount
    Loop
    rsListPlan.Close
    Set rsListPlan = Nothing
    End If
    
End Sub
Private Sub lstPlanList_Click()
Dim StrSql As String
Dim StrSql1 As String
Dim PLN As New ADODB.Recordset
Dim ListLoop As Integer
Dim ListCount As Integer
Dim rst3 As New ADODB.Recordset
Dim StrSql2 As String


    If lstPlanList.ListItems.Count <> 0 Then
        CmdAdd.Enabled = True
        CmdEdit.Enabled = False
        CmdDelete.Enabled = True
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        
        StrSql1 = "Select * from PLN where PLNCode = '" & lstPlanList.SelectedItem.SubItems(1) & "'"
       
        PLN.Open StrSql1, MediConnCISDB, adOpenKeyset, adLockOptimistic
       
        cboPLNAnnLimit = PLN!AnnLmtIND
        cboPLNVisitLimit = PLN!VisLmtIND
        cboPLNPremium = PLN!PRMInd
        cboPLNMCO = PLN!MCOInd
        If Trim(PLN!PDentalStatus) <> "" Then
            CboPDentalStatus = PLN!PDentalStatus
        End If
        
        If Trim(PLN!PDentalLimit) <> "" Then
            CboPDentalLimit = PLN!PDentalLimit
        End If
        
        If Trim(PLN!PSpecStatus) <> "" Then
            CboPSpectStatus = PLN!PSpecStatus
        End If
       
        If Trim(PLN!PSpecLimit) <> "" Then
            CboPSpectLimit = PLN!PSpecLimit
        End If
        
        If Trim(PLN!PMEPStatus) <> "" Then
            CboPMEPStatus = PLN!PMEPStatus
        End If
        
        If Trim(PLN!ANNLmtIndSupp) <> "" Then
            CboSAnnLimit = PLN!ANNLmtIndSupp
        End If
        
        If Trim(PLN!VISLmtIndSupp) <> "" Then
            CboSVisitLimit = PLN!VISLmtIndSupp
        End If
       
        If Trim(PLN!SDentalStatus) <> "" Then
            CboSDentalStatus = PLN!SDentalStatus
        End If
        
        If Trim(PLN!SDentalLimit) <> "" Then
            CboSDentalLimit = PLN!SDentalLimit
        End If
       
        If Trim(PLN!SSpecStatus) <> "" Then
            CboSSpecStatus = PLN!SSpecStatus
        End If
        
        If Trim(PLN!SSpecLimit) <> "" Then
            CboSSpecLimit = PLN!SSpecLimit
        End If
        
        If Trim(PLN!SMEPStatus) <> "" Then
            CboSMEPStatus = PLN!SMEPStatus
        End If
       
        txtPLN1 = PLN!PLNCode
        If Trim(PLN!PlnDesc) <> "" Then
             txtDesc1 = PLN!PlnDesc
        End If
              
        txtPSPLimit = IIf(IsNumeric(PLN!PrinSPLimit), PLN!PrinSPLimit, 0)
        txtPGPLimit = IIf(IsNumeric(PLN!PrinGPLimit), PLN!PrinGPLimit, 0)
        txtsSPLimit = IIf(IsNumeric(PLN!DepSPLimit), PLN!DepSPLimit, 0)
        txtSGPLimit = IIf(IsNumeric(PLN!DepGPLimit), PLN!DepGPLimit, 0)
        If IsNull(PLN!OPFamilyStatus) Then
        
        Else
            cmbOPStatus = (PLN!OPFamilyStatus)
        End If
     
        StrSql2 = "Select * from PLN Where PLNCode = '" & lstPlanList.SelectedItem.SubItems(1) & "'"
        rst3.Open StrSql2, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If Trim(rst3!HLTCode) <> "" Then
            cboPLNHLTCode = rst3!HLTCode
        End If
        
        If Trim(rst3!PAYCode) <> "" Then
            cboPLNPayor = rst3!PAYCode
        End If
        
        If Trim(rst3!HISPAYCode) <> "" Then
            cboPLNPayor2 = rst3!HISPAYCode
        End If
        
        If Trim(rst3!GRPCompany) <> "" Then
            txtGrpCom = rst3!GRPCompany & " - " & Trim(lstPlanList.SelectedItem.SubItems(6))
        End If
        
        If Trim(rst3!ClnContract) <> "" Then
            cboPLNCLNContract = rst3!ClnContract
        End If
        
        If Trim(rst3!AnnLmtIND) <> "" Then
            cboPLNAnnLimit = rst3!AnnLmtIND
        End If
      
        If Trim(rst3!VisLmtIND) <> "" Then
            cboPLNVisitLimit = rst3!VisLmtIND
        End If
      
        If Trim(rst3!PRMInd) <> "" Then
            cboPLNPremium = rst3!PRMInd
        End If
        
        If Trim(rst3!MCOInd) <> "" Then
            cboPLNMCO = rst3!MCOInd
        End If
        
        If Trim(rst3!PDentalStatus) <> "" Then
            CboPDentalStatus = rst3!PDentalStatus
        End If
        
        If Trim(rst3!PDentalLimit) <> "" Then
            CboPDentalLimit = rst3!PDentalLimit
        End If
        
        If Trim(rst3!PSpecStatus) <> "" Then
            CboPSpectStatus = rst3!PSpecStatus
        End If
       
        If Trim(rst3!PSpecLimit) <> "" Then
            CboPSpectLimit = rst3!PSpecLimit
        End If
        
        If Trim(rst3!PMEPStatus) <> "" Then
            CboPMEPStatus = rst3!PMEPStatus
        End If
        
        If Trim(rst3!ANNLmtIndSupp) <> "" Then
            CboSAnnLimit = rst3!ANNLmtIndSupp
        End If
        
        If Trim(rst3!VISLmtIndSupp) <> "" Then
            CboSVisitLimit = rst3!VISLmtIndSupp
        End If
       
        If Trim(rst3!SDentalStatus) <> "" Then
            CboSDentalStatus = rst3!SDentalStatus
        End If
        
        If Trim(rst3!SDentalLimit) <> "" Then
            CboSDentalLimit = rst3!SDentalLimit
        End If
       
        If Trim(rst3!SSpecStatus) <> "" Then
            CboSSpecStatus = rst3!SSpecStatus
        End If
        
        If Trim(rst3!SSpecLimit) <> "" Then
            CboSSpecLimit = rst3!SSpecLimit
        End If
        
        If Trim(rst3!SMEPStatus) <> "" Then
            CboSMEPStatus = rst3!SMEPStatus
        End If
        
         If Trim(rst3!Insured) = "Y" Then
            cmbInsured = "Y-Yes"
         ElseIf Trim(rst3!Insured) = "N" Then
            cmbInsured = "N-No"
        End If
         If Trim(rst3!HRApproval) = "Y" Then
            cmbHRApproval = "Y-Yes"
         ElseIf Trim(rst3!HRApproval) = "N" Then
            cmbHRApproval = "N-No"
         End If
      '  MediConnCISDB.Close
      '  Set MediConnCISDB = Nothing
        
        If Trim(rst3!PLNSpecialRemark) <> "" Then
            txtSpecialNote.Text = Trim(rst3!PLNSpecialRemark)
        End If
        
        Call EntriesEnable(False)
        txtSpecialNote.Enabled = True
        CmdDelete.Enabled = True
        CmdSave.Enabled = False
    End If
End Sub

Private Sub lstPlanList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim StrSql As String
Dim StrSql1 As String
Dim PLN As New ADODB.Recordset
Dim ListLoop As Integer
Dim ListCount As Integer
Dim rst3 As New ADODB.Recordset
Dim StrSql2 As String


    If lstPlanList.ListItems.Count <> 0 Then
        CmdAdd.Enabled = True
        CmdEdit.Enabled = False
        CmdDelete.Enabled = True
        Call ClearForm(Me)
       
        'If Mid(lstPlanList.SelectedItem.Text, 1, 1) = "N" Then
            'strsql1 = "Select PLNCOde, * from PLN where PAYCode = '" & lstPlanList.SelectedItem.SubItems(3) & "'"
        'Else
            'strsql1 = "Select PLNCOde, * from PLN where HLTCode = '" & Trim(lstPlanList.SelectedItem.Text) & "'"
        'End If
        
        StrSql1 = "Select * from PLN where PLNCode = '" & lstPlanList.SelectedItem.SubItems(1) & "'"
       
        'strsql1 = strsql1 & "ORDER BY RIGHT(PLNCOde, 2)"
       
        PLN.Open StrSql1, MediConnCISDB, adOpenKeyset, adLockOptimistic
       
        cboPLNAnnLimit = PLN!AnnLmtIND
        cboPLNVisitLimit = PLN!VisLmtIND
        cboPLNPremium = PLN!PRMInd
        cboPLNMCO = PLN!MCOInd
        If Trim(PLN!PDentalStatus) <> "" Then
            CboPDentalStatus = PLN!PDentalStatus
        End If
        
        If Trim(PLN!PDentalLimit) <> "" Then
            CboPDentalLimit = PLN!PDentalLimit
        End If
        
        If Trim(PLN!PSpecStatus) <> "" Then
            CboPSpectStatus = PLN!PSpecStatus
        End If
       
        If Trim(PLN!PSpecLimit) <> "" Then
            CboPSpectLimit = PLN!PSpecLimit
        End If
        
        If Trim(PLN!PMEPStatus) <> "" Then
            CboPMEPStatus = PLN!PMEPStatus
        End If
        
        If Trim(PLN!ANNLmtIndSupp) <> "" Then
            CboSAnnLimit = PLN!ANNLmtIndSupp
        End If
        
        If Trim(PLN!VISLmtIndSupp) <> "" Then
            CboSVisitLimit = PLN!VISLmtIndSupp
        End If
       
        If Trim(PLN!SDentalStatus) <> "" Then
            CboSDentalStatus = PLN!SDentalStatus
        End If
        
        If Trim(PLN!SDentalLimit) <> "" Then
            CboSDentalLimit = PLN!SDentalLimit
        End If
       
        If Trim(PLN!SSpecStatus) <> "" Then
            CboSSpecStatus = PLN!SSpecStatus
        End If
        
        If Trim(PLN!SSpecLimit) <> "" Then
            CboSSpecLimit = PLN!SSpecLimit
        End If
        
        If Trim(PLN!SMEPStatus) <> "" Then
            CboSMEPStatus = PLN!SMEPStatus
        End If
              
       'If PLN.RecordCount > 6 Then
         'PLN.Filter = " PLNCode >='" & lstPlanList.SelectedItem.SubItems(1) & "'"
      
       'End If
       'Do Until PLN.EOF
       'ListCount = IIf(PLN.RecordCount >= 6, 6, PLN.RecordCount)
       'For ListLoop = 1 To ListCount
           'If ListLoop = "1" Then
               txtPLN1 = PLN!PLNCode
               If Trim(PLN!PlnDesc) <> "" Then
                    txtDesc1 = PLN!PlnDesc
               End If
               'If Trim(PLN!NewPLNCode) <> "" Then
                'txtNewPLN1 = PLN!NewPLNCode
               'End If
                'If Len(PLN!PLNGenWP) Then
                    'TxtGenWP1 = PLN!PLNGenWP
                'End If
                'If Len(PLN!PLNPreWP) Then
                    'txtPREWP1 = PLN!PLNPreWP
                'End If
                'If Len(PLN!PLNSpeWP) Then
                    'TxtSpeWP1 = PLN!PLNSpeWP
                'End If

               'If Trim(PLN!NewPLNDescription) <> "" Then
                'txtNewDecs1 = PLN!NewPLNDescription
               'End If
           'ElseIf ListLoop = "2" Then
               'txtPLN2 = PLN!PLNCode
                'If Trim(PLN!PLNDesc) <> "" Then
                    'txtDecs2 = PLN!PLNDesc
                'End If
                'If Trim(PLN!NewPLNCode) <> "" Then
                    'txtNewPLN2 = PLN!NewPLNCode
                'End If
                'If Len(PLN!PLNGenWP) Then
                    'TxtGenWP2 = PLN!PLNGenWP
                'End If
                'If Len(PLN!PLNPreWP) Then
                    'txtPREWP2 = PLN!PLNPreWP
                'End If
                'If Len(PLN!PLNSpeWP) Then
                    'TxtSpeWP2 = PLN!PLNSpeWP
                'End If
               
               'If Trim(PLN!NewPLNDescription) <> "" Then
                'txtNewDecs2 = PLN!NewPLNDescription
               'End If
           'ElseIf ListLoop = "3" Then
               'txtPLN3 = PLN!PLNCode
                'If Trim(PLN!PLNDesc) <> "" Then
                 'txtDesc3 = PLN!PLNDesc
                'End If
                'If Trim(PLN!NewPLNCode) <> "" Then
                 'txtNewPLN3 = PLN!NewPLNCode
                'End If
                'If Len(PLN!PLNGenWP) Then
                    'TxtGenWP3 = PLN!PLNGenWP
                'End If
                'If Len(PLN!PLNPreWP) Then
                    'txtPREWP3 = PLN!PLNPreWP
                'End If
                'If Len(PLN!PLNSpeWP) Then
                    'TxtSpeWP3 = PLN!PLNSpeWP
                'End If
               
               'If Trim(PLN!NewPLNDescription) <> "" Then
                'txtNewDecs3 = PLN!NewPLNDescription
               'End If
           'ElseIf ListLoop = "4" Then
               'txtPLN4 = PLN!PLNCode
                'If Trim(PLN!PLNDesc) <> "" Then
                  'txtDecs4 = PLN!PLNDesc
                'End If
                'If Trim(PLN!NewPLNCode) <> "" Then
                 'txtNewPLN4 = PLN!NewPLNCode
                'End If
                'If Len(PLN!PLNGenWP) Then
                    'TxtGenWP4 = PLN!PLNGenWP
                'End If
                'If Len(PLN!PLNPreWP) Then
                    'txtPREWP4 = PLN!PLNPreWP
                'End If
                'If Len(PLN!PLNSpeWP) Then
                    'TxtSpeWP4 = PLN!PLNSpeWP
                'End If
               
               'If Trim(PLN!NewPLNDescription) <> "" Then
                'txtNewDecs4 = PLN!NewPLNDescription
               'End If
           'ElseIf ListLoop = "5" Then
               'txtPLN5 = PLN!PLNCode
                'If Trim(PLN!PLNDesc) <> "" Then
                'txtDesc5 = PLN!PLNDesc
                'End If
                'If Trim(PLN!NewPLNCode) <> "" Then
                 'txtNewPLN5 = PLN!NewPLNCode
                'End If
                'If Len(PLN!PLNGenWP) Then
                    'TxtGenWP5 = PLN!PLNGenWP
                'End If
                'If Len(PLN!PLNPreWP) Then
                    'txtPREWP5 = PLN!PLNPreWP
                'End If
                'If Len(PLN!PLNSpeWP) Then
                    'TxtSpeWP5 = PLN!PLNSpeWP
                'End If
               
               'If Trim(PLN!NewPLNDescription) <> "" Then
                'txtNewDecs5 = PLN!NewPLNDescription
               'End If
           'ElseIf ListLoop = "6" Then
               'txtPLN6 = PLN!PLNCode
                'If Trim(PLN!PLNDesc) <> "" Then
                 'txtDecs6 = PLN!PLNDesc
                'End If
                'If Trim(PLN!NewPLNCode) <> "" Then
                 'txtNewPLN6 = PLN!NewPLNCode
                'End If
                'If Len(PLN!PLNGenWP) Then
                    'TxtGenWP6 = PLN!PLNGenWP
                'End If
                'If Len(PLN!PLNPreWP) Then
                    'txtPREWP6 = PLN!PLNPreWP
                'End If
                'If Len(PLN!PLNSpeWP) Then
                    'TxtSpeWP6 = PLN!PLNSpeWP
                'End If
               
               'If Trim(PLN!NewPLNDescription) <> "" Then
                'txtNewDecs6 = PLN!NewPLNDescription
               'End If
      'End If
     
       'PLN.MoveNext
       
       'Next ListLoop
       'If ListLoop > 6 Then
         'Exit Do
       'End If
       'Loop
      
        StrSql2 = "Select * from PLN Where PLNCode = '" & lstPlanList.SelectedItem.SubItems(1) & "'"
        rst3.Open StrSql2, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If Trim(rst3!HLTCode) <> "" Then
            cboPLNHLTCode = rst3!HLTCode
        End If
        
        If Trim(rst3!PAYCode) <> "" Then
            cboPLNPayor = rst3!PAYCode
        End If
        
        If Trim(rst3!HISPAYCode) <> "" Then
            cboPLNPayor2 = rst3!HISPAYCode
        End If
        
        If Trim(rst3!GRPCompany) <> "" Then
            txtGrpCom = rst3!GRPCompany
        End If
        
        If Trim(rst3!ClnContract) <> "" Then
            cboPLNCLNContract = rst3!ClnContract
        End If
        
        If Trim(rst3!AnnLmtIND) <> "" Then
            cboPLNAnnLimit = rst3!AnnLmtIND
        End If
      
        If Trim(rst3!VisLmtIND) <> "" Then
            cboPLNVisitLimit = rst3!VisLmtIND
        End If
      
        If Trim(rst3!PRMInd) <> "" Then
            cboPLNPremium = rst3!PRMInd
        End If
        
        If Trim(rst3!MCOInd) <> "" Then
            cboPLNMCO = rst3!MCOInd
        End If
        
        If Trim(rst3!PDentalStatus) <> "" Then
            CboPDentalStatus = rst3!PDentalStatus
        End If
        
        If Trim(rst3!PDentalLimit) <> "" Then
            CboPDentalLimit = rst3!PDentalLimit
        End If
        
        If Trim(rst3!PSpecStatus) <> "" Then
            CboPSpectStatus = rst3!PSpecStatus
        End If
       
        If Trim(rst3!PSpecLimit) <> "" Then
            CboPSpectLimit = rst3!PSpecLimit
        End If
        
        If Trim(rst3!PMEPStatus) <> "" Then
            CboPMEPStatus = rst3!PMEPStatus
        End If
        
        If Trim(rst3!ANNLmtIndSupp) <> "" Then
            CboSAnnLimit = rst3!ANNLmtIndSupp
        End If
        
        If Trim(rst3!VISLmtIndSupp) <> "" Then
            CboSVisitLimit = rst3!VISLmtIndSupp
        End If
       
        If Trim(rst3!SDentalStatus) <> "" Then
            CboSDentalStatus = rst3!SDentalStatus
        End If
        
        If Trim(rst3!SDentalLimit) <> "" Then
            CboSDentalLimit = rst3!SDentalLimit
        End If
       
        If Trim(rst3!SSpecStatus) <> "" Then
            CboSSpecStatus = rst3!SSpecStatus
        End If
        
        If Trim(rst3!SSpecLimit) <> "" Then
            CboSSpecLimit = rst3!SSpecLimit
        End If
        
        If Trim(rst3!SMEPStatus) <> "" Then
            CboSMEPStatus = rst3!SMEPStatus
        End If
     
       Call EntriesEnable(False)
       CmdDelete.Enabled = True
       CmdSave.Enabled = False
    End If
End Sub

Private Sub lstPlanList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim StrSql As String
Dim StrSql1 As String
Dim PLN As New ADODB.Recordset
Dim ListLoop As Integer
Dim ListCount As Integer
Dim rst3 As New ADODB.Recordset
Dim StrSql2 As String


    If lstPlanList.ListItems.Count <> 0 Then
        CmdAdd.Enabled = True
        CmdEdit.Enabled = False
        CmdDelete.Enabled = True
        Call ClearForm(Me)
       
        'If Mid(lstPlanList.SelectedItem.Text, 1, 1) = "N" Then
            'strsql1 = "Select PLNCOde, * from PLN where PAYCode = '" & lstPlanList.SelectedItem.SubItems(3) & "'"
        'Else
            'strsql1 = "Select PLNCOde, * from PLN where HLTCode = '" & Trim(lstPlanList.SelectedItem.Text) & "'"
        'End If
        
        StrSql1 = "Select * from PLN where PLNCode = '" & lstPlanList.SelectedItem.SubItems(1) & "'"
       
        'strsql1 = strsql1 & "ORDER BY RIGHT(PLNCOde, 2)"
       
        PLN.Open StrSql1, MediConnCISDB, adOpenKeyset, adLockOptimistic
       
        cboPLNAnnLimit = PLN!AnnLmtIND
        cboPLNVisitLimit = PLN!VisLmtIND
        cboPLNPremium = PLN!PRMInd
        cboPLNMCO = PLN!MCOInd
        If Trim(PLN!PDentalStatus) <> "" Then
            CboPDentalStatus = PLN!PDentalStatus
        End If
        
        If Trim(PLN!PDentalLimit) <> "" Then
            CboPDentalLimit = PLN!PDentalLimit
        End If
        
        If Trim(PLN!PSpecStatus) <> "" Then
            CboPSpectStatus = PLN!PSpecStatus
        End If
       
        If Trim(PLN!PSpecLimit) <> "" Then
            CboPSpectLimit = PLN!PSpecLimit
        End If
        
        If Trim(PLN!PMEPStatus) <> "" Then
            CboPMEPStatus = PLN!PMEPStatus
        End If
        
        If Trim(PLN!ANNLmtIndSupp) <> "" Then
            CboSAnnLimit = PLN!ANNLmtIndSupp
        End If
        
        If Trim(PLN!VISLmtIndSupp) <> "" Then
            CboSVisitLimit = PLN!VISLmtIndSupp
        End If
       
        If Trim(PLN!SDentalStatus) <> "" Then
            CboSDentalStatus = PLN!SDentalStatus
        End If
        
        If Trim(PLN!SDentalLimit) <> "" Then
            CboSDentalLimit = PLN!SDentalLimit
        End If
       
        If Trim(PLN!SSpecStatus) <> "" Then
            CboSSpecStatus = PLN!SSpecStatus
        End If
        
        If Trim(PLN!SSpecLimit) <> "" Then
            CboSSpecLimit = PLN!SSpecLimit
        End If
        
        If Trim(PLN!SMEPStatus) <> "" Then
            CboSMEPStatus = PLN!SMEPStatus
        End If
              
       'If PLN.RecordCount > 6 Then
         'PLN.Filter = " PLNCode >='" & lstPlanList.SelectedItem.SubItems(1) & "'"
      
       'End If
       'Do Until PLN.EOF
       'ListCount = IIf(PLN.RecordCount >= 6, 6, PLN.RecordCount)
       'For ListLoop = 1 To ListCount
           'If ListLoop = "1" Then
               txtPLN1 = PLN!PLNCode
               If Trim(PLN!PlnDesc) <> "" Then
                    txtDesc1 = PLN!PlnDesc
               End If
               'If Trim(PLN!NewPLNCode) <> "" Then
                'txtNewPLN1 = PLN!NewPLNCode
               'End If
                'If Len(PLN!PLNGenWP) Then
                    'TxtGenWP1 = PLN!PLNGenWP
                'End If
                'If Len(PLN!PLNPreWP) Then
                    'txtPREWP1 = PLN!PLNPreWP
                'End If
                'If Len(PLN!PLNSpeWP) Then
                    'TxtSpeWP1 = PLN!PLNSpeWP
                'End If

               'If Trim(PLN!NewPLNDescription) <> "" Then
                'txtNewDecs1 = PLN!NewPLNDescription
               'End If
           'ElseIf ListLoop = "2" Then
               'txtPLN2 = PLN!PLNCode
                'If Trim(PLN!PLNDesc) <> "" Then
                    'txtDecs2 = PLN!PLNDesc
                'End If
                'If Trim(PLN!NewPLNCode) <> "" Then
                    'txtNewPLN2 = PLN!NewPLNCode
                'End If
                'If Len(PLN!PLNGenWP) Then
                    'TxtGenWP2 = PLN!PLNGenWP
                'End If
                'If Len(PLN!PLNPreWP) Then
                    'txtPREWP2 = PLN!PLNPreWP
                'End If
                'If Len(PLN!PLNSpeWP) Then
                    'TxtSpeWP2 = PLN!PLNSpeWP
                'End If
               
               'If Trim(PLN!NewPLNDescription) <> "" Then
                'txtNewDecs2 = PLN!NewPLNDescription
               'End If
           'ElseIf ListLoop = "3" Then
               'txtPLN3 = PLN!PLNCode
                'If Trim(PLN!PLNDesc) <> "" Then
                 'txtDesc3 = PLN!PLNDesc
                'End If
                'If Trim(PLN!NewPLNCode) <> "" Then
                 'txtNewPLN3 = PLN!NewPLNCode
                'End If
                'If Len(PLN!PLNGenWP) Then
                    'TxtGenWP3 = PLN!PLNGenWP
                'End If
                'If Len(PLN!PLNPreWP) Then
                    'txtPREWP3 = PLN!PLNPreWP
                'End If
                'If Len(PLN!PLNSpeWP) Then
                    'TxtSpeWP3 = PLN!PLNSpeWP
                'End If
               
               'If Trim(PLN!NewPLNDescription) <> "" Then
                'txtNewDecs3 = PLN!NewPLNDescription
               'End If
           'ElseIf ListLoop = "4" Then
               'txtPLN4 = PLN!PLNCode
                'If Trim(PLN!PLNDesc) <> "" Then
                  'txtDecs4 = PLN!PLNDesc
                'End If
                'If Trim(PLN!NewPLNCode) <> "" Then
                 'txtNewPLN4 = PLN!NewPLNCode
                'End If
                'If Len(PLN!PLNGenWP) Then
                    'TxtGenWP4 = PLN!PLNGenWP
                'End If
                'If Len(PLN!PLNPreWP) Then
                    'txtPREWP4 = PLN!PLNPreWP
                'End If
                'If Len(PLN!PLNSpeWP) Then
                    'TxtSpeWP4 = PLN!PLNSpeWP
                'End If
               
               'If Trim(PLN!NewPLNDescription) <> "" Then
                'txtNewDecs4 = PLN!NewPLNDescription
               'End If
           'ElseIf ListLoop = "5" Then
               'txtPLN5 = PLN!PLNCode
                'If Trim(PLN!PLNDesc) <> "" Then
                'txtDesc5 = PLN!PLNDesc
                'End If
                'If Trim(PLN!NewPLNCode) <> "" Then
                 'txtNewPLN5 = PLN!NewPLNCode
                'End If
                'If Len(PLN!PLNGenWP) Then
                    'TxtGenWP5 = PLN!PLNGenWP
                'End If
                'If Len(PLN!PLNPreWP) Then
                    'txtPREWP5 = PLN!PLNPreWP
                'End If
                'If Len(PLN!PLNSpeWP) Then
                    'TxtSpeWP5 = PLN!PLNSpeWP
                'End If
               
               'If Trim(PLN!NewPLNDescription) <> "" Then
                'txtNewDecs5 = PLN!NewPLNDescription
               'End If
           'ElseIf ListLoop = "6" Then
               'txtPLN6 = PLN!PLNCode
                'If Trim(PLN!PLNDesc) <> "" Then
                 'txtDecs6 = PLN!PLNDesc
                'End If
                'If Trim(PLN!NewPLNCode) <> "" Then
                 'txtNewPLN6 = PLN!NewPLNCode
                'End If
                'If Len(PLN!PLNGenWP) Then
                    'TxtGenWP6 = PLN!PLNGenWP
                'End If
                'If Len(PLN!PLNPreWP) Then
                    'txtPREWP6 = PLN!PLNPreWP
                'End If
                'If Len(PLN!PLNSpeWP) Then
                    'TxtSpeWP6 = PLN!PLNSpeWP
                'End If
               
               'If Trim(PLN!NewPLNDescription) <> "" Then
                'txtNewDecs6 = PLN!NewPLNDescription
               'End If
      'End If
     
       'PLN.MoveNext
       
       'Next ListLoop
       'If ListLoop > 6 Then
         'Exit Do
       'End If
       'Loop
      
        StrSql2 = "Select * from PLN Where PLNCode = '" & lstPlanList.SelectedItem.SubItems(1) & "'"
        rst3.Open StrSql2, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If Trim(rst3!HLTCode) <> "" Then
            cboPLNHLTCode = rst3!HLTCode
        End If
        
        If Trim(rst3!PAYCode) <> "" Then
            cboPLNPayor = rst3!PAYCode
        End If
        
        If Trim(rst3!HISPAYCode) <> "" Then
            cboPLNPayor2 = rst3!HISPAYCode
        End If
        
        If Trim(rst3!GRPCompany) <> "" Then
            txtGrpCom = rst3!GRPCompany
        End If
        
        If Trim(rst3!ClnContract) <> "" Then
            cboPLNCLNContract = rst3!ClnContract
        End If
        
        If Trim(rst3!AnnLmtIND) <> "" Then
            cboPLNAnnLimit = rst3!AnnLmtIND
        End If
      
        If Trim(rst3!VisLmtIND) <> "" Then
            cboPLNVisitLimit = rst3!VisLmtIND
        End If
      
        If Trim(rst3!PRMInd) <> "" Then
            cboPLNPremium = rst3!PRMInd
        End If
        
        If Trim(rst3!MCOInd) <> "" Then
            cboPLNMCO = rst3!MCOInd
        End If
        
        If Trim(rst3!PDentalStatus) <> "" Then
            CboPDentalStatus = rst3!PDentalStatus
        End If
        
        If Trim(rst3!PDentalLimit) <> "" Then
            CboPDentalLimit = rst3!PDentalLimit
        End If
        
        If Trim(rst3!PSpecStatus) <> "" Then
            CboPSpectStatus = rst3!PSpecStatus
        End If
       
        If Trim(rst3!PSpecLimit) <> "" Then
            CboPSpectLimit = rst3!PSpecLimit
        End If
        
        If Trim(rst3!PMEPStatus) <> "" Then
            CboPMEPStatus = rst3!PMEPStatus
        End If
        
        If Trim(rst3!ANNLmtIndSupp) <> "" Then
            CboSAnnLimit = rst3!ANNLmtIndSupp
        End If
        
        If Trim(rst3!VISLmtIndSupp) <> "" Then
            CboSVisitLimit = rst3!VISLmtIndSupp
        End If
       
        If Trim(rst3!SDentalStatus) <> "" Then
            CboSDentalStatus = rst3!SDentalStatus
        End If
        
        If Trim(rst3!SDentalLimit) <> "" Then
            CboSDentalLimit = rst3!SDentalLimit
        End If
       
        If Trim(rst3!SSpecStatus) <> "" Then
            CboSSpecStatus = rst3!SSpecStatus
        End If
        
        If Trim(rst3!SSpecLimit) <> "" Then
            CboSSpecLimit = rst3!SSpecLimit
        End If
        
        If Trim(rst3!SMEPStatus) <> "" Then
            CboSMEPStatus = rst3!SMEPStatus
        End If
     
       Call EntriesEnable(False)
       CmdDelete.Enabled = True
       CmdSave.Enabled = False
    End If
End Sub

Private Sub lstPlanList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstPlanList.SortKey = ColumnHeader.Index - 1
    lstPlanList.Sorted = True
    If lstPlanList.SortOrder = lvwAscending Then
       lstPlanList.SortOrder = lvwDescending
    Else
       lstPlanList.SortOrder = lvwAscending
    End If
 
End Sub

Private Sub PLNDisEnable()
    'txtPLN1.Enabled = False
    'txtPLN2.Enabled = False
    'txtPLN3.Enabled = False
    'txtPLN4.Enabled = False
    'txtPLN5.Enabled = False
    'txtPLN6.Enabled = False
    
    'txtNewPLN1.Enabled = False
    'txtNewPLN2.Enabled = False
    'txtNewPLN3.Enabled = False
    'txtNewPLN4.Enabled = False
    'txtNewPLN5.Enabled = False
    'txtNewPLN6.Enabled = False
    'TxtGenWP1.Enabled = False
    'txtPREWP1.Enabled = False
    'TxtSpeWP1.Enabled = False
    'TxtGenWP2.Enabled = False
    'txtPREWP2.Enabled = False
    'TxtSpeWP2.Enabled = False
    'TxtGenWP3.Enabled = False
    'txtPREWP3.Enabled = False
    'TxtSpeWP3.Enabled = False
    'TxtGenWP4.Enabled = False
    'txtPREWP4.Enabled = False
    'TxtSpeWP4.Enabled = False
    'TxtGenWP5.Enabled = False
    'txtPREWP5.Enabled = False
    'TxtSpeWP5.Enabled = False
    'TxtGenWP6.Enabled = False
    'txtPREWP6.Enabled = False
    'TxtSpeWP6.Enabled = False
End Sub
'########################## End Plan Code ##########################'

'############################### Start Premium ###############################'
Private Sub DeletePRM()
'On Error GoTo DeleleErr
Dim rsDeletePRM As New ADODB.Recordset
Dim rsCheckPRM As New ADODB.Recordset
Dim StrSql As String
Dim sqlPRM As String
Dim DeleteRecord

EditFlag = False
    
 If cboPmINSCode = "" Then
    MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
    'cboPmINSCode.SetFocus
    ElseIf lstPmPRMList.ListItems.Count = 0 Then
       MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
   sqlPRM = "Select MBMNumber From MBM"
   
   sqlPRM = sqlPRM & " Where (INSCode = '" & Trim(cboPmINSCode) & "') AND (HLTCode = '" & Trim(cboPmHLTCode) & "')" & _
   " AND (PLNCode = '" & Trim(cboPmPLNCode) & "') and (AGECode = '" & Trim(cboPmAgeCode) & "') "
   If ChkStr(cboPmHLTCode) <> "S" Then
        sqlPRM = sqlPRM & " and (PAYCode = '" & Trim(cboPmPAYCode) & "') "
    End If
    
    
    rsCheckPRM.MaxRecords = 1
    rsCheckPRM.Open sqlPRM, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If rsCheckPRM.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in Membership table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    
        If DeleteRecord = vbYes And rsCheckPRM.RecordCount = 0 Then
             StrSql = "Delete From PRM Where PRMCode = '" & ChkStr(lstPmPRMList.SelectedItem.Text) & "'"
             rsDeletePRM.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             Call ClearForm(Me)
             Call PRMListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             CmdAdd.Enabled = True
             CmdSave.Enabled = True
        Else
             CmdDelete.Enabled = False
             CmdEdit.Enabled = False
        End If
    End If
    Exit Sub
  End If
End Sub


Private Sub SearchPRM()
Dim strAppend As String
Dim location As Integer
    
    Screen.MousePointer = vbHourglass
    
    strAppend = ""
    
    If Len(Trim(cboPmINSCode)) Then
        strAppend = strAppend + " And INSCode = '" & Trim(Mid(cboPmINSCode, 1, IIf(InStr(1, cboPmINSCode, "-") = 0, 4, InStr(1, cboPmINSCode, "-") - 1))) & "'"
        'strAppend = strAppend + " And INSCode = '" & Trim(Mid(cboPmINSCode, 1, 2)) & "'"
    End If
    If Len(Trim(cboPmPAYCode)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(cboPmPAYCode, 1, IIf(InStr(1, cboPmPAYCode, "-") = 0, 4, InStr(1, cboPmPAYCode, "-") - 1))) & "'"
        'strAppend = strAppend + " And PAYCode = '" & Trim(Mid(cboPmPAYCode, 1, 2)) & "'"
    End If
    If Len(Trim(cboPmHLTCode)) Then
        strAppend = strAppend + " And HLTCode = '" & Trim(Mid(cboPmHLTCode, 1, IIf(InStr(1, cboPmHLTCode, "-") = 0, 4, InStr(1, cboPmHLTCode, "-") - 1))) & "'"
        'strAppend = strAppend + " And HLTCode = '" & Trim(Mid(cboPmHLTCode, 1, 2)) & "'"
    End If
    If Len(Trim(cboPmAgeCode)) Then
        strAppend = strAppend + " And AgeCode = '" & Trim(Mid(cboPmAgeCode, 1, IIf(InStr(1, cboPmAgeCode, "-") = 0, 4, InStr(1, cboPmAgeCode, "-") - 1))) & "'"
        'strAppend = strAppend + " And AgeCode = '" & Trim(Mid(cboPmAgeCode, 1, 2)) & "'"
    End If
    If Len(Trim(cboPmPLNCode)) Then
        location = InStr(1, cboPmPLNCode, "-") - 1
        If location = -1 Then
            location = 2
        End If
        strAppend = strAppend + " And PLNCOde = '" & Trim(Mid(cboPmPLNCode, 1, location)) & "'"
    End If
   
    PRMCriteria = strAppend
    Call PRMListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub PrintPRM()
    RptPremiumList.Show
End Sub

Private Sub SavePRM()
Dim PRM As New ADODB.Recordset
Dim SavePRM As New ADODB.Recordset
Dim RecordSaved
Dim StrSql As String
Dim ListLoop As Integer
Dim ListCount As Integer
Dim Plan As String

Plan = ""

    If Len(Trim(cboPmINSCode)) = 0 Then
        MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboPmINSCode.SetFocus
    ElseIf Len(Trim(cboPmHLTCode)) = 0 Then
        MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboPmHLTCode.SetFocus
    'ElseIf Len(Trim(cboPmAgeCode)) = 0 Then
    '    MsgBox "Please Enter Age Code", vbOKOnly + vbCritical, DialogBoxTitle
    '    cboPmAgeCode.SetFocus
    ElseIf Len(Trim(cboPmPLNCode)) = 0 Then
        MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboPmPLNCode.SetFocus
    ElseIf txtPRMEffdate = "__/__/____" Or Not IsDate(txtPRMEffdate) Then
        MsgBox "Please validate your Premium Effective date", vbOKOnly + vbCritical, DialogBoxTitle
        txtPRMEffdate.SetFocus
    'ElseIf Mid(Combo3, 1, 1) = "B" Or Mid(Combo3, 1, 1) = "C" Then
    '    If Text2 = "" Then
    '        MsgBox "Please Enter Std Supp Premium ", vbOKOnly + vbCritical, DialogBoxTitle
    '        Text2.SetFocus
    '    ElseIf Text3 = "" Then
    '        MsgBox "Please Enter Std Supp Premium ", vbOKOnly + vbCritical, DialogBoxTitle
    '        Text3.SetFocus
    '    ElseIf Text4 = "" Then
    '        MsgBox "Please Enter Std Supp Premium ", vbOKOnly + vbCritical, DialogBoxTitle
    '        Text4.SetFocus
    '    ElseIf Text5 = "" Then
    '        MsgBox "Please Enter Std Supp Premium ", vbOKOnly + vbCritical, DialogBoxTitle
    '        Text5.SetFocus
    '    ElseIf Text6 = "" Then
    '        MsgBox "Please Enter Std Supp Premium ", vbOKOnly + vbCritical, DialogBoxTitle
    '        Text6.SetFocus
    '    ElseIf Text7 = "" Then
    '        MsgBox "Please Enter Std Supp Premium ", vbOKOnly + vbCritical, DialogBoxTitle
    '        Text7.SetFocus
    '    'End If
    '
    'ElseIf Len(Trim(txtPmPRMAmount)) = 0 Then
    '    MsgBox "Please Enter Premium Amount", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtPmPRMAmount.SetFocus
    Else
        Dim locationPLN As Integer
        Dim locationPAY As Integer
        StrSql = "Select * From PRM "
        locationPLN = InStr(1, cboPmPLNCode, "-") - 1
        If locationPLN = -1 Then
            locationPLN = 4
        End If
        
        locationPAY = InStr(1, cboPmPAYCode, "-") - 1
        If locationPAY = -1 Then
            locationPAY = 2
        End If
        StrSql = StrSql & " Where " & _
                " INSCode = '" & Trim(Mid(cboPmINSCode, 1, 2)) & "' " & _
                " AND HLTCode = '" & Trim(Mid(cboPmHLTCode, 1, 2)) & "' " & _
                " AND PLNCode = '" & Trim(Mid(cboPmPLNCode, 1, locationPLN)) & "' " & _
                " and AGECode = '" & Trim(Mid(cboPmAgeCode, 1, 2)) & "' " & _
                " and PRMeffdate = '" & Format(txt01, "mm/dd/yyyy") & "'"
                'and PRMeffdate = '" & Format(txtPRMEffdate, "mm/dd/yyyy") & "'"
'" PRMCode =" & txtPrmCode
        If ChkStr(Mid(cboPmHLTCode, 1, 2)) <> "S" Then
               StrSql = StrSql & " and (PAYCode = '" & Trim(Mid(cboPmPAYCode, 1, locationPAY)) & "') "
        End If
        
        PRM.MaxRecords = 1
        PRM.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
            If PRM.RecordCount > 0 And EditFlag = False Then
               MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
               PRM.Close
               cboPmINSCode.SetFocus
               Set PRM = Nothing
               Exit Sub
            
            Else
            RecordSaved = MsgBox("Do you want to save this record?", vbYesNo + vbExclamation)
            
            If RecordSaved = vbYes Then
               
               If EditFlag = False Then
                  SavePRM.Open "PRM", MediConnCISDB, adOpenKeyset, adLockOptimistic
                  ListCount = 7
                  For ListLoop = 1 To ListCount
                  SavePRM.AddNew
                    
                    
                    With SavePRM
                    
                    Plan = Trim(Mid(cboPmPLNCode, 1, IIf(InStr(1, cboPmPLNCode, "-") = 0, 4, InStr(1, cboPmPLNCode, "-") - 1)))
                    
                        !INSCode = ChkStr(Mid(cboPmINSCode, 1, 2))
                        !PAYCode = ChkStr(Mid(cboPmPAYCode, 1, 2))
                        !HLTCode = ChkStr(Mid(cboPmHLTCode, 1, 2))
                        !AgeCode = ChkStr(Mid(cboPmAgeCode, 1, 2))
                        
                        If Mid(cboPmHLTCode, 1, 1) = "S" Then
                            !PLNCode = ChkStr(Plan)
                        Else
                            !PLNCode = ChkStr(Plan)
                        End If
                        
                        If Mid(Combo3, 1, 1) = "A" Then
                            !SuppPrmStatus = "A"
                            
                            If ListLoop = 1 Then
                                If txt01 <> "" Then
                                    !PRMAmount = txt01
                                    !AgeCode = "01"
                                End If
                            ElseIf ListLoop = 2 Then
                                If txt02 <> "" Then
                                    !PRMAmount = txt02
                                    !AgeCode = "02"
                                End If
                            ElseIf ListLoop = 3 Then
                                If txt03 <> "" Then
                                    !PRMAmount = txt03
                                    !AgeCode = "03"
                                End If
                            ElseIf ListLoop = 4 Then
                                If txt04 <> "" Then
                                    !PRMAmount = txt04
                                    !AgeCode = "04"
                                End If
                            ElseIf ListLoop = 5 Then
                                If txt05 <> "" Then
                                    !PRMAmount = txt05
                                    !AgeCode = "05"
                                End If
                            ElseIf ListLoop = 6 Then
                                If txt06 <> "" Then
                                    !PRMAmount = txt06
                                    !AgeCode = "06"
                                End If
                            ElseIf ListLoop = 7 Then
                                If txt07 <> "" Then
                                    !PRMAmount = txt07
                                    !AgeCode = "07"
                                End If
                            End If
                            
                        ElseIf Mid(Combo3, 1, 1) = "B" Then
                            !SuppPrmStatus = "B"
                            
                            If ListLoop = 1 Then
                                If txt01 <> "" Then
                                    !PRMAmount = txt01
                                    If Text2 = "" Then
                                        !SuppPrmAmt = Text2
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AgeCode = "01"
                                End If
                            ElseIf ListLoop = 2 Then
                                If txt02 <> "" Then
                                    !PRMAmount = txt02
                                    If Text3 = "" Then
                                        !SuppPrmAmt = Text3
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AgeCode = "02"
                                End If
                            ElseIf ListLoop = 3 Then
                                If txt03 <> "" Then
                                    !PRMAmount = txt03
                                    If Text4 = "" Then
                                        !SuppPrmAmt = Text4
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AgeCode = "03"
                                End If
                            ElseIf ListLoop = 4 Then
                                If txt04 <> "" Then
                                    !PRMAmount = txt04
                                    If Text5 = "" Then
                                        !SuppPrmAmt = Text5
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AgeCode = "04"
                                End If
                            ElseIf ListLoop = 5 Then
                                If txt05 <> "" Then
                                    !PRMAmount = txt05
                                    If Text6 = "" Then
                                        !SuppPrmAmt = Text6
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AgeCode = "05"
                                End If
                            ElseIf ListLoop = 6 Then
                                If txt06 <> "" Then
                                    !PRMAmount = txt06
                                    If Text7 = "" Then
                                        !SuppPrmAmt = Text7
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AgeCode = "06"
                                End If
                            ElseIf ListLoop = 7 Then
                                If txt07 <> "" Then
                                    !PRMAmount = txt07
                                    If Text1 = "" Then
                                        !SuppPrmAmt = Text1
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AgeCode = "07"
                                End If
                            End If
                        
                        ElseIf Mid(Combo3, 1, 1) = "C" Then
                            !SuppPrmStatus = "C"
                            
                            If ListLoop = 1 Then
                                If txt01 <> "" Then
                                    !PRMAmount = txt01
                                    If Text2 = "" Then
                                        !SuppPrmAmt = Text2
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AgeCode = "01"
                                End If
                            ElseIf ListLoop = 2 Then
                                If txt02 <> "" Then
                                    !PRMAmount = txt02
                                    If Text3 = "" Then
                                        !SuppPrmAmt = Text3
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AgeCode = "02"
                                End If
                            ElseIf ListLoop = 3 Then
                                If txt03 <> "" Then
                                    !PRMAmount = txt03
                                    If Text4 = "" Then
                                        !SuppPrmAmt = Text4
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AgeCode = "03"
                                End If
                            ElseIf ListLoop = 4 Then
                                If txt04 <> "" Then
                                    !PRMAmount = txt04
                                    If Text5 = "" Then
                                        !SuppPrmAmt = Text5
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AgeCode = "04"
                                End If
                            ElseIf ListLoop = 5 Then
                                If txt05 <> "" Then
                                    !PRMAmount = txt05
                                    If Text6 = "" Then
                                        !SuppPrmAmt = Text6
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AgeCode = "05"
                                End If
                            ElseIf ListLoop = 6 Then
                                If txt06 <> "" Then
                                    !PRMAmount = txt06
                                    If Text7 = "" Then
                                        !SuppPrmAmt = Text7
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AgeCode = "06"
                                End If
                            ElseIf ListLoop = 7 Then
                                If txt07 <> "" Then
                                    !PRMAmount = txt07
                                    If Text1 = "" Then
                                        !SuppPrmAmt = Text1
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AgeCode = "07"
                                End If
                            End If
                        End If
                        
                        
                        !PRMEffDAte = CDate(txtPRMEffdate)
                        If Len(txtPREVersion) Then
                            !PRMVersion = txtPREVersion
                        End If
                        
                        SavePRM!PRMLastUpdateUser = Logon
                        SavePRM!PRMLastUpdateDate = Date
                    End With
                    SavePRM.update
                 'End If
                 Next ListLoop

               Else
                  SavePRM.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                  SavePRM.update
               End If
               Call ClearForm(Me)
                
               Call PRMListing(PRMCriteria)
                
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                SavePRM.Close
                Set SavePRM = Nothing
            End If
            
            EditFlag = False
        End If
        End If
        'End If
        cboPmAgeCode.Enabled = True

  'End If
End Sub

Private Sub lstPMPRMList_Click()
Dim PRM As New ADODB.Recordset
Dim StrSql As String
Dim AGE01, AGE02, AGE03, AGE04, AGE05, AGE06, AGE07 As String
ListClick = True
    
   
    If lstPmPRMList.ListItems.Count <> 0 Then
       ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
          Call setCmdButton(Me, False)
          Call ClearForm(Me)
             
          txtPRMEffdate.Enabled = False
          FrameAGEBand.Enabled = False
          Combo3.Enabled = False
          txtPREVersion.Enabled = False
          
          StrSql = "Select * from PRM where INSCode = '" & lstPmPRMList.SelectedItem.SubItems(1) & "' And PLNCode = '" & lstPmPRMList.SelectedItem.SubItems(5) & "'"
          PRM.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
          Do Until PRM.EOF
              If PRM!AgeCode = "01" Then
                  AGE01 = PRM!PRMAmount
              ElseIf PRM!AgeCode = "02" Then
                  AGE02 = PRM!PRMAmount
              ElseIf PRM!AgeCode = "03" Then
                  AGE03 = PRM!PRMAmount
              ElseIf PRM!AgeCode = "04" Then
                  AGE04 = PRM!PRMAmount
              ElseIf PRM!AgeCode = "05" Then
                  AGE05 = PRM!PRMAmount
              ElseIf PRM!AgeCode = "06" Then
                  AGE06 = PRM!PRMAmount
              ElseIf PRM!AgeCode = "07" Then
                  AGE07 = PRM!PRMAmount
         
         End If
         PRM.MoveNext
          Loop
         PRM.Close
         Set PRM = Nothing
           
                      
          'if lstPmPRMList.SelectedItem.SubItems(1) =
          txtPrmCode = lstPmPRMList.SelectedItem.Text
          cboPmINSCode.Text = lstPmPRMList.SelectedItem.SubItems(1)
          cboPmINSCode.Enabled = False
          cboPmHLTCode.Text = lstPmPRMList.SelectedItem.SubItems(3)
          cboPmHLTCode.Enabled = False
          cboPmPAYCode.Text = lstPmPRMList.SelectedItem.SubItems(2)
          cboPmPAYCode.Enabled = False
          cboPmAgeCode.Text = lstPmPRMList.SelectedItem.SubItems(4)
          cboPmAgeCode.Enabled = False
          cboPmPLNCode.Text = lstPmPRMList.SelectedItem.SubItems(5)
          cboPmPLNCode.Enabled = False
          txt01 = Format(AGE01, "#,###0.00")
          txt02 = Format(AGE02, "#,###0.00")
          txt03 = Format(AGE03, "#,###0.00")
          txt04 = Format(AGE04, "#,###0.00")
          txt05 = Format(AGE05, "#,###0.00")
          txt06 = Format(AGE06, "#,###0.00")
          txt07 = Format(AGE07, "#,###0.00")
          FrameAGEBand.Enabled = False
          
          AGE01 = 0
          AGE02 = 0
          AGE03 = 0
          AGE04 = 0
          AGE05 = 0
          AGE06 = 0
          AGE07 = 0
          'txtPmPRMAmount.Text = Format(lstPmPRMList.SelectedItem.SubItems(6), "#,###0.00")
          'txtPmPRMAmount.Enabled = False
          
          StrSql = "Select * from PRM where INSCode = '" & lstPmPRMList.SelectedItem.SubItems(1) & "' And PLNCode = '" & lstPmPRMList.SelectedItem.SubItems(5) & "'"
          PRM.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
          Do Until PRM.EOF
              If PRM!AgeCode = "01" Then
                  If Len(PRM!SuppPrmAmt) Then
                      AGE01 = PRM!SuppPrmAmt
                  End If
              ElseIf PRM!AgeCode = "02" Then
                  If Len(PRM!SuppPrmAmt) Then
                      AGE02 = PRM!SuppPrmAmt
                  End If
              ElseIf PRM!AgeCode = "03" Then
                  If Len(PRM!SuppPrmAmt) Then
                      AGE03 = PRM!SuppPrmAmt
                  End If
              ElseIf PRM!AgeCode = "04" Then
                  If Len(PRM!SuppPrmAmt) Then
                      AGE04 = PRM!SuppPrmAmt
                  End If
              ElseIf PRM!AgeCode = "05" Then
                  If Len(PRM!SuppPrmAmt) Then
                      AGE05 = PRM!SuppPrmAmt
                  End If
              ElseIf PRM!AgeCode = "06" Then
                  If Len(PRM!SuppPrmAmt) Then
                      AGE06 = PRM!SuppPrmAmt
                  End If
              ElseIf PRM!AgeCode = "07" Then
                  If Len(PRM!SuppPrmAmt) Then
                      AGE07 = PRM!SuppPrmAmt
                  End If
         
         End If
         PRM.MoveNext
          Loop
         PRM.Close
         Set PRM = Nothing
          
          Text2 = Format(AGE01, "#,###0.00")
          Text3 = Format(AGE02, "#,###0.00")
          Text4 = Format(AGE03, "#,###0.00")
          Text5 = Format(AGE04, "#,###0.00")
          Text6 = Format(AGE05, "#,###0.00")
          Text7 = Format(AGE06, "#,###0.00")
          Text1 = Format(AGE07, "#,###0.00")
          Frame6.Enabled = False
          
          txtPRMEffdate.mask = ""
          txtPRMEffdate = lstPmPRMList.SelectedItem.SubItems(8)
          txtPRMEffdate.mask = "##/##/####"
          txtPREVersion = lstPmPRMList.SelectedItem.SubItems(9)
          
          If lstPmPRMList.SelectedItem.SubItems(10) = "A" Then
              Combo3.ListIndex = 0
          ElseIf lstPmPRMList.SelectedItem.SubItems(10) = "B" Then
              Combo3.ListIndex = 1
          ElseIf lstPmPRMList.SelectedItem.SubItems(10) = "C" Then
              Combo3.ListIndex = 2
          End If
          'Call EntriesEnable(False)
          CmdAdd.Enabled = True
        '  cmdSave.Enabled = False
         ' cmdDelete.Enabled = False
     '   MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
    End If
End Sub

Private Sub lstPmPRMList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstPmPRMList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstPmPRMList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstPmPRMList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstPmPRMList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstPmPRMList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstPmPRMList, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstPmPRMList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstPmPRMList, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstPmPRMList, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstPmPRMList, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    End Select
End Sub

Private Sub lstPmPRMList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim PRM As New ADODB.Recordset
Dim StrSql As String
Dim AGE01, AGE02, AGE03, AGE04, AGE05, AGE06, AGE07 As String
ListClick = True
    
       
    If lstPmPRMList.ListItems.Count <> 0 Then
    
    Call setCmdButton(Me, False)
    Call ClearForm(Me)
       
    txtPRMEffdate.Enabled = False
    FrameAGEBand.Enabled = False
    Combo3.Enabled = False
    txtPREVersion.Enabled = False
    
    StrSql = "Select * from PRM where INSCode = '" & lstPmPRMList.SelectedItem.SubItems(1) & "' And PLNCode = '" & lstPmPRMList.SelectedItem.SubItems(5) & "'"
    PRM.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do Until PRM.EOF
        If PRM!AgeCode = "01" Then
            AGE01 = PRM!PRMAmount
        ElseIf PRM!AgeCode = "02" Then
            AGE02 = PRM!PRMAmount
        ElseIf PRM!AgeCode = "03" Then
            AGE03 = PRM!PRMAmount
        ElseIf PRM!AgeCode = "04" Then
            AGE04 = PRM!PRMAmount
        ElseIf PRM!AgeCode = "05" Then
            AGE05 = PRM!PRMAmount
        ElseIf PRM!AgeCode = "06" Then
            AGE06 = PRM!PRMAmount
        ElseIf PRM!AgeCode = "07" Then
            AGE07 = PRM!PRMAmount
   
   End If
   PRM.MoveNext
    Loop
   PRM.Close
   Set PRM = Nothing
   
                
    'if lstPmPRMList.SelectedItem.SubItems(1) =
    txtPrmCode = lstPmPRMList.SelectedItem.Text
    cboPmINSCode.Text = lstPmPRMList.SelectedItem.SubItems(1)
    cboPmINSCode.Enabled = False
    cboPmHLTCode.Text = lstPmPRMList.SelectedItem.SubItems(3)
    cboPmHLTCode.Enabled = False
    cboPmPAYCode.Text = lstPmPRMList.SelectedItem.SubItems(2)
    cboPmPAYCode.Enabled = False
    cboPmAgeCode.Text = lstPmPRMList.SelectedItem.SubItems(4)
    cboPmAgeCode.Enabled = False
    cboPmPLNCode.Text = lstPmPRMList.SelectedItem.SubItems(5)
    cboPmPLNCode.Enabled = False
    txt01 = Format(AGE01, "#,###0.00")
    txt02 = Format(AGE02, "#,###0.00")
    txt03 = Format(AGE03, "#,###0.00")
    txt04 = Format(AGE04, "#,###0.00")
    txt05 = Format(AGE05, "#,###0.00")
    txt06 = Format(AGE06, "#,###0.00")
    txt07 = Format(AGE07, "#,###0.00")
    FrameAGEBand.Enabled = False
    'txtPmPRMAmount.Text = Format(lstPmPRMList.SelectedItem.SubItems(6), "#,###0.00")
    'txtPmPRMAmount.Enabled = False
    
    AGE01 = 0
    AGE02 = 0
    AGE03 = 0
    AGE04 = 0
    AGE05 = 0
    AGE06 = 0
    AGE07 = 0
    
    StrSql = "Select * from PRM where INSCode = '" & lstPmPRMList.SelectedItem.SubItems(1) & "' And PLNCode = '" & lstPmPRMList.SelectedItem.SubItems(5) & "'"
    PRM.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do Until PRM.EOF
        If PRM!AgeCode = "01" Then
            If Len(PRM!SuppPrmAmt) Then
                AGE01 = PRM!SuppPrmAmt
            End If
        ElseIf PRM!AgeCode = "02" Then
            If Len(PRM!SuppPrmAmt) Then
                AGE02 = PRM!SuppPrmAmt
            End If
        ElseIf PRM!AgeCode = "03" Then
            If Len(PRM!SuppPrmAmt) Then
                AGE03 = PRM!SuppPrmAmt
            End If
        ElseIf PRM!AgeCode = "04" Then
            If Len(PRM!SuppPrmAmt) Then
                AGE04 = PRM!SuppPrmAmt
            End If
        ElseIf PRM!AgeCode = "05" Then
            If Len(PRM!SuppPrmAmt) Then
                AGE05 = PRM!SuppPrmAmt
            End If
        ElseIf PRM!AgeCode = "06" Then
            If Len(PRM!SuppPrmAmt) Then
                AGE06 = PRM!SuppPrmAmt
            End If
        ElseIf PRM!AgeCode = "07" Then
            If Len(PRM!SuppPrmAmt) Then
                AGE07 = PRM!SuppPrmAmt
            End If
   
   End If
   PRM.MoveNext
    Loop
   PRM.Close
   Set PRM = Nothing
    
    Text2 = Format(AGE01, "#,###0.00")
    Text3 = Format(AGE02, "#,###0.00")
    Text4 = Format(AGE03, "#,###0.00")
    Text5 = Format(AGE04, "#,###0.00")
    Text6 = Format(AGE05, "#,###0.00")
    Text7 = Format(AGE06, "#,###0.00")
    Text1 = Format(AGE07, "#,###0.00")
    Frame6.Enabled = False
    
    txtPRMEffdate.mask = ""
    txtPRMEffdate = lstPmPRMList.SelectedItem.SubItems(8)
    txtPRMEffdate.mask = "##/##/####"
    txtPREVersion = lstPmPRMList.SelectedItem.SubItems(9)
    
    If lstPmPRMList.SelectedItem.SubItems(10) = "A" Then
        Combo3.ListIndex = 0
    ElseIf lstPmPRMList.SelectedItem.SubItems(10) = "B" Then
        Combo3.ListIndex = 1
    ElseIf lstPmPRMList.SelectedItem.SubItems(10) = "C" Then
        Combo3.ListIndex = 2
    End If
    
    'Call EntriesEnable(False)
    CmdAdd.Enabled = True
  '  cmdSave.Enabled = False
   ' cmdDelete.Enabled = False
    End If
End Sub

Private Sub lstPmPRMList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim PRM As New ADODB.Recordset
Dim StrSql As String
Dim AGE01, AGE02, AGE03, AGE04, AGE05, AGE06, AGE07 As String
ListClick = True
    
      
    If lstPmPRMList.ListItems.Count <> 0 Then
    
    Call setCmdButton(Me, False)
    Call ClearForm(Me)
       
    txtPRMEffdate.Enabled = False
    FrameAGEBand.Enabled = False
    Combo3.Enabled = False
    txtPREVersion.Enabled = False
    
    StrSql = "Select * from PRM where INSCode = '" & lstPmPRMList.SelectedItem.SubItems(1) & "' And PLNCode = '" & lstPmPRMList.SelectedItem.SubItems(5) & "'"
    PRM.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do Until PRM.EOF
        If PRM!AgeCode = "01" Then
            AGE01 = PRM!PRMAmount
        ElseIf PRM!AgeCode = "02" Then
            AGE02 = PRM!PRMAmount
        ElseIf PRM!AgeCode = "03" Then
            AGE03 = PRM!PRMAmount
        ElseIf PRM!AgeCode = "04" Then
            AGE04 = PRM!PRMAmount
        ElseIf PRM!AgeCode = "05" Then
            AGE05 = PRM!PRMAmount
        ElseIf PRM!AgeCode = "06" Then
            AGE06 = PRM!PRMAmount
        ElseIf PRM!AgeCode = "07" Then
            AGE07 = PRM!PRMAmount
   
   End If
   PRM.MoveNext
    Loop
   PRM.Close
   Set PRM = Nothing
   
                
    'if lstPmPRMList.SelectedItem.SubItems(1) =
    txtPrmCode = lstPmPRMList.SelectedItem.Text
    cboPmINSCode.Text = lstPmPRMList.SelectedItem.SubItems(1)
    cboPmINSCode.Enabled = False
    cboPmHLTCode.Text = lstPmPRMList.SelectedItem.SubItems(3)
    cboPmHLTCode.Enabled = False
    cboPmPAYCode.Text = lstPmPRMList.SelectedItem.SubItems(2)
    cboPmPAYCode.Enabled = False
    cboPmAgeCode.Text = lstPmPRMList.SelectedItem.SubItems(4)
    cboPmAgeCode.Enabled = False
    cboPmPLNCode.Text = lstPmPRMList.SelectedItem.SubItems(5)
    cboPmPLNCode.Enabled = False
    txt01 = Format(AGE01, "#,###0.00")
    txt02 = Format(AGE02, "#,###0.00")
    txt03 = Format(AGE03, "#,###0.00")
    txt04 = Format(AGE04, "#,###0.00")
    txt05 = Format(AGE05, "#,###0.00")
    txt06 = Format(AGE06, "#,###0.00")
    txt07 = Format(AGE07, "#,###0.00")
    FrameAGEBand.Enabled = False
    'txtPmPRMAmount.Text = Format(lstPmPRMList.SelectedItem.SubItems(6), "#,###0.00")
    'txtPmPRMAmount.Enabled = False
    AGE01 = 0
    AGE02 = 0
    AGE03 = 0
    AGE04 = 0
    AGE05 = 0
    AGE06 = 0
    AGE07 = 0
    
    StrSql = "Select * from PRM where INSCode = '" & lstPmPRMList.SelectedItem.SubItems(1) & "' And PLNCode = '" & lstPmPRMList.SelectedItem.SubItems(5) & "'"
    PRM.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do Until PRM.EOF
        If PRM!AgeCode = "01" Then
            If Len(PRM!SuppPrmAmt) Then
                AGE01 = PRM!SuppPrmAmt
            End If
        ElseIf PRM!AgeCode = "02" Then
            If Len(PRM!SuppPrmAmt) Then
                AGE02 = PRM!SuppPrmAmt
            End If
        ElseIf PRM!AgeCode = "03" Then
            If Len(PRM!SuppPrmAmt) Then
                AGE03 = PRM!SuppPrmAmt
            End If
        ElseIf PRM!AgeCode = "04" Then
            If Len(PRM!SuppPrmAmt) Then
                AGE04 = PRM!SuppPrmAmt
            End If
        ElseIf PRM!AgeCode = "05" Then
            If Len(PRM!SuppPrmAmt) Then
                AGE05 = PRM!SuppPrmAmt
            End If
        ElseIf PRM!AgeCode = "06" Then
            If Len(PRM!SuppPrmAmt) Then
                AGE06 = PRM!SuppPrmAmt
            End If
        ElseIf PRM!AgeCode = "07" Then
            If Len(PRM!SuppPrmAmt) Then
                AGE07 = PRM!SuppPrmAmt
            End If
   
   End If
   PRM.MoveNext
    Loop
   PRM.Close
   Set PRM = Nothing
    
    Text2 = Format(AGE01, "#,###0.00")
    Text3 = Format(AGE02, "#,###0.00")
    Text4 = Format(AGE03, "#,###0.00")
    Text5 = Format(AGE04, "#,###0.00")
    Text6 = Format(AGE05, "#,###0.00")
    Text7 = Format(AGE06, "#,###0.00")
    Text1 = Format(AGE07, "#,###0.00")
    Frame6.Enabled = False
    
    txtPRMEffdate.mask = ""
    txtPRMEffdate = lstPmPRMList.SelectedItem.SubItems(8)
    txtPRMEffdate.mask = "##/##/####"
    txtPREVersion = lstPmPRMList.SelectedItem.SubItems(9)
    
    If lstPmPRMList.SelectedItem.SubItems(10) = "A" Then
        Combo3.ListIndex = 0
    ElseIf lstPmPRMList.SelectedItem.SubItems(10) = "B" Then
        Combo3.ListIndex = 1
    ElseIf lstPmPRMList.SelectedItem.SubItems(10) = "C" Then
        Combo3.ListIndex = 2
    End If
    'Call EntriesEnable(False)
    CmdAdd.Enabled = True
  '  cmdSave.Enabled = False
   ' cmdDelete.Enabled = False
    End If
End Sub

'List down all the premium list into the comboBox
Private Sub PRMComboListing()
Dim INS As New ADODB.Recordset
Dim PAY As New ADODB.Recordset
Dim HLT As New ADODB.Recordset
Dim AGE As New ADODB.Recordset
Dim PLN As New ADODB.Recordset
    
  INS.Open "INS", MediConnCISDB, adOpenKeyset, adLockOptimistic
  PAY.Open "PAY", MediConnCISDB, adOpenKeyset, adLockOptimistic
  HLT.Open "HLT", MediConnCISDB, adOpenKeyset, adLockOptimistic
  AGE.Open "Age", MediConnCISDB, adOpenKeyset, adLockOptimistic
  PLN.Open "PLN", MediConnCISDB, adOpenKeyset, adLockOptimistic
    
  cboPmINSCode.Clear
  cboPmHLTCode.Clear
  cboPmPAYCode.Clear
  cboPmPLNCode.Clear
  cboPmAgeCode.Clear
  Combo3.Clear
    
  
  Do Until INS.EOF
     cboPmINSCode.AddItem INS!INSCode & " - " & INS!InsDesc
     INS.MoveNext
  Loop
  
  Do Until HLT.EOF
     cboPmHLTCode.AddItem HLT!HLTCode & " - " & HLT!HLTDesc
     HLT.MoveNext
  Loop
  
  Do Until PAY.EOF
     cboPmPAYCode.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
     PAY.MoveNext
  Loop
    
  Do Until AGE.EOF
     cboPmAgeCode.AddItem AGE!AgeCode & " - " & AGE!AGEDesc
     AGE.MoveNext
  Loop
    
  Do Until PLN.EOF
     cboPmPLNCode.AddItem PLN!PLNCode & " - " & PLN!PlnDesc
     PLN.MoveNext
  Loop
  
  Combo3.AddItem "A" & " - " & "Share Limit for Family"
  Combo3.AddItem "B" & " - " & "Share Limit for Supplementary"
  Combo3.AddItem "C" & " - " & "Separate Limit for Supplementary"
  Combo3.Text = "A" & " - " & "Share Limit for Family"
    
    INS.Close
    Set INS = Nothing
    PAY.Close
    Set PAY = Nothing
    HLT.Close
    Set HLT = Nothing
    AGE.Close
    Set AGE = Nothing
    PLN.Close
    Set PLN = Nothing
End Sub

'List down all the premium list into the listbox
Private Sub PRMListing(Optional strAppend As String)
    On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim PRM As ListItem
    Dim sql As String
    Dim TotalPRMCount As Integer
    
    TotalPRMCount = 0
    lstPmPRMList.ListItems.Clear
        
    sql = "SELECT PRMCode, INSCode, " & _
      " PAYCode, HLTCode, PRMVersion, SuppPrmAmt, " & _
      " AGECode, PLNCode, PRMAmount , PRMEffDate, SuppPrmStatus " & _
      " From PRM where 0=0 "
      
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    
    Do Until rst2.EOF
            Set PRM = lstPmPRMList.ListItems.Add(, , rst2!PRMCode)
                PRM.SubItems(1) = rst2!INSCode
                PRM.SubItems(2) = rst2!PAYCode
                PRM.SubItems(3) = rst2!HLTCode
                PRM.SubItems(4) = rst2!AgeCode
                PRM.SubItems(5) = rst2!PLNCode
                PRM.SubItems(6) = Format(rst2!PRMAmount, "#,##0.00")
                PRM.SubItems(7) = Format(rst2!SuppPrmAmt, "#,##0.00")
                PRM.SubItems(8) = rst2!PRMEffDAte
                PRM.SubItems(9) = rst2!PRMVersion
                PRM.SubItems(10) = rst2!SuppPrmStatus

                rst2.MoveNext
           TotalPRMCount = TotalPRMCount + 1
           txtTotalRecord = TotalPRMCount
           Loop

    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

'############################### End Premium ###############################'

'############################## Start Annual ##############################'
Private Sub SaveANN()
Dim Annual As New ADODB.Recordset
Dim SaveAnnual As New ADODB.Recordset
Dim RecordSaved
Dim StrSql As String
    
    If Len(Trim(CboPlnCode)) = 0 Then
        MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboPlnCode.SetFocus
    ElseIf Len(Trim(CboHltCode)) = 0 Then
        MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboHltCode.SetFocus
    ElseIf Len(Trim(CboInsType)) = 0 Then
        MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboInsType.SetFocus
    ElseIf Len(Trim(TxtAnnualLimit)) = 0 Then
        MsgBox "Please Enter Principal Annual Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtAnnualLimit.SetFocus
    ElseIf Len(Trim(TxtPDentalAnnLmt)) = 0 Then
        MsgBox "Please Enter Principal Dental Annual Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtPDentalAnnLmt.SetFocus
    ElseIf Len(Trim(TxtPSpecAnnLmt)) = 0 Then
        MsgBox "Please Enter Principal Spectacles Annual Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtPSpecAnnLmt.SetFocus
    ElseIf Len(Trim(TxtPMEPAnnLmt)) = 0 Then
        MsgBox "Please Enter Principal MEP Annual Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtPMEPAnnLmt.SetFocus
    ElseIf txtALEffDate = "__/__/____" Then
        MsgBox "Please Enter Annual Limit Effective Date", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf txtANNVersion = "" Then
        MsgBox "Please Enter Annual Limit Version", vbOKOnly + vbCritical, DialogBoxTitle
        txtANNVersion.SetFocus
    ElseIf Mid(Combo1, 1, 1) = "B" And TxtSupplyAL = "" Then
        MsgBox "Please Enter Supp Premium Limit", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Mid(Combo1, 1, 1) = "B" And TxtSDentalAnnLmt = "" Then
        MsgBox "Please Enter Supplementary Dental Annual Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtSDentalAnnLmt.SetFocus
    ElseIf Mid(Combo1, 1, 1) = "B" And TxtSSpecAnnLmt = "" Then
        MsgBox "Please Enter Supplementary Spectacles Annual Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtSSpecAnnLmt.SetFocus
    ElseIf Mid(Combo1, 1, 1) = "B" And TxtSMEPAnnLmt = "" Then
        MsgBox "Please Enter Supplementary MEP Annual Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtSMEPAnnLmt.SetFocus
    ElseIf Mid(Combo1, 1, 1) = "C" And TxtSupplyAL = "" Then
        MsgBox "Please Enter Supp Premium Limit", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Mid(Combo1, 1, 1) = "C" And TxtSDentalAnnLmt = "" Then
        MsgBox "Please Enter Supplementary Dental Annual Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtSDentalAnnLmt.SetFocus
    ElseIf Mid(Combo1, 1, 1) = "C" And TxtSSpecAnnLmt = "" Then
        MsgBox "Please Enter Supplementary Spectacles Annual Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtSSpecAnnLmt.SetFocus
    ElseIf Mid(Combo1, 1, 1) = "C" And TxtSMEPAnnLmt = "" Then
        MsgBox "Please Enter Supplementary MEP Annual Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtSMEPAnnLmt.SetFocus
    Else
        
        StrSql = " Select * From AnnualLimit Where " & _
                 " INSCode = '" & Trim(Mid(CboInsType, 1, 1)) & "' " & _
                 " AND HLTCode = '" & Trim(Mid(CboHltCode, 1, 1)) & "' " & _
                 " AND PLNCode = '" & Trim(Mid(CboPlnCode, 1, 4)) & "' " & _
                 " AND AnnLimitVersion = '" & Trim(txtANNVersion) & "'" & _
                 " AND AnnEffDate = '" & Format(txtALEffDate, "mm/dd/yyyy") & "'"
        Annual.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
      If Annual.RecordCount > 0 And EditFlag = False Then
        MsgBox "Record Duplicated! " & vbNewLine & _
                " Please Change The Corresponding Field", vbCritical, "Error"
        Annual.Close
        CboInsType.SetFocus
        CboInsType.SelStart = 0
        CboInsType.SelLength = Len(CboInsType)
        Set Annual = Nothing
        Exit Sub
      Else
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                    SaveAnnual.Open "AnnualLimit", MediConnCISDB, adOpenKeyset, adLockOptimistic
                    SaveAnnual.AddNew
                Else
                    
                    StrSql = " Select * From AnnualLimit Where AnnualIndex = '" & ChkStr(lstAnnualList.SelectedItem.SubItems(1)) & "'"
                    
                    SaveAnnual.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                    
                End If
            
                With SaveAnnual
                    If InStr(1, CboPlnCode, "-") > 0 Then
                        !PLNCode = Trim(Mid(CboPlnCode, 1, InStr(1, CboPlnCode, "-") - 1))
                    Else
                        !PLNCode = Trim(Mid(CboPlnCode, 1, 4))
                    End If
                    
                    !INSCode = Mid(CboInsType, 1, 1)
                    !HLTCode = Mid(CboHltCode, 1, 1)
                    !AnnLimit = TxtAnnualLimit
                    !PDentalAnnLimit = TxtPDentalAnnLmt
                    !PSpecAnnLimit = TxtPSpecAnnLmt
                    !PMEPAnnLimit = TxtPMEPAnnLmt
                    
                    If Mid(Combo1, 1, 1) = "A" Then
                        !SuppLimitStatus = "A"
                    ElseIf Mid(Combo1, 1, 1) = "B" Then
                        !SuppLimitStatus = "B"
                        !SuppLimit = TxtSupplyAL
                        !SDentalAnnLimit = TxtSDentalAnnLmt
                        !SSpecAnnLimit = TxtSSpecAnnLmt
                        !SMEPAnnLimit = TxtSMEPAnnLmt
                    ElseIf Mid(Combo1, 1, 1) = "C" Then
                        !SuppLimitStatus = "C"
                        !SuppLimit = TxtSupplyAL
                        !SDentalAnnLimit = TxtSDentalAnnLmt
                        !SSpecAnnLimit = TxtSSpecAnnLmt
                        !SMEPAnnLimit = TxtSMEPAnnLmt
                    End If
                                        
                    !AnnEffDate = IIf(txtALEffDate = "__/__/____", "", CDate(txtALEffDate))
                    !AnnLimitVersion = txtANNVersion
                    !AnnLastUpdateUser = Logon
                    !AnnLastUpdateDate = Date
                End With
                SaveAnnual.update
                Call ClearForm(Me)
                Call AnnualListing(AnnualCriteria)
                    If EditFlag = False Then
                        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    Else
                        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
 
                SaveAnnual.Close
                Set SaveAnnual = Nothing
            End If
        EditFlag = False
   End If
  End If
  'End If
  Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub DeleteANN()
Dim DeleteAnnual As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim StrSql, sqlMBM As String
Dim DeleteRecord

EditFlag = False
If CboPlnCode = "" Then
   MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
   'CboPlnCode.SetFocus
ElseIf lstAnnualList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBM where PLNCode = '" & RemStr(CboPlnCode) & "'" 'ChkStr(lstAnnualList.SelectedItem.Text) & "'"
   
   rsCheckMBM.MaxRecords = 1
   rsCheckMBM.Open sqlMBM, MediConnCISDB, adOpenKeyset, adLockOptimistic
   
     
   If rsCheckMBM.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.RecordCount = 0 Then
           StrSql = " Delete From AnnualLimit Where AnnualIndex = '" & ChkStr(lstAnnualList.SelectedItem.SubItems(1)) & "'"
                    
           DeleteAnnual.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call AnnualListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           CboPlnCode.SetFocus
         End If
     End If
End If
Exit Sub
End Sub

Private Sub SearchANN()
    Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(CboPlnCode)) Then
        strAppend = strAppend + " And PLNCode = '" & Trim(Mid(CboPlnCode, 1, IIf(InStr(1, CboPlnCode, "-") = 0, 4, InStr(1, CboPlnCode, "-") - 1))) & "'"
    End If
    If Len(Trim(CboHltCode)) Then
        strAppend = strAppend + " And HLTCode = '" & Trim(Mid(CboHltCode, 1, IIf(InStr(1, CboHltCode, "-") = 0, 4, InStr(1, CboHltCode, "-") - 1))) & "'"
        'strAppend = strAppend + " And HLTCode = '" & Trim(Mid(CboHltCode, 1, 2)) & "'"
    End If
    
    If Len(Trim(CboInsType)) Then
        strAppend = strAppend + " And INSCode = '" & Trim(Mid(CboInsType, 1, IIf(InStr(1, CboInsType, "-") = 0, 4, InStr(1, CboInsType, "-") - 1))) & "'"
        'strAppend = strAppend + " And INSCode = '" & Trim(Mid(CboInsType, 1, 2)) & "'"
    End If
    
    If Len(Trim(CboPayorCode)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayorCode, 1, IIf(InStr(1, CboPayorCode, "-") = 0, 4, InStr(1, CboPayorCode, "-") - 1))) & "'"
    End If

    If Len(Trim(CboGrpCom)) Then
        strAppend = strAppend + " AND GRPCompany like '" & RemStr(Mid(CboGrpCom, 1, IIf(InStr(1, CboGrpCom, "-") = 0, 4, InStr(1, CboGrpCom, "-") - 1))) & "%'"
    End If
    
    AnnualCriteria = strAppend
    Call AnnualListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub AnnualListing(Optional strAppend As String)
'On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim AnnualMaster As ListItem
    Dim sql As String
    Dim TotalAnnualCount As Integer
    
    
    TotalAnnualCount = 0
    lstAnnualList.ListItems.Clear
    
    sql = " Select * From VIEW_ANN_PAY_PLN where 0=0 "
    
    If Len(Trim(strAppend)) Then
    sql = sql + strAppend
    End If

    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set AnnualMaster = lstAnnualList.ListItems.Add(, , rst2!HLTCode)
                AnnualMaster.SubItems(1) = IIf(Len(rst2!AnnualIndex), rst2!AnnualIndex, "")
                AnnualMaster.SubItems(2) = IIf(Len(rst2!PLNCode), rst2!PLNCode, "")
                AnnualMaster.SubItems(3) = IIf(Len(rst2!INSCode), rst2!INSCode, "")
                AnnualMaster.SubItems(4) = Format(rst2!AnnLimit, "#,##0.00")
                AnnualMaster.SubItems(5) = Format(rst2!PDentalAnnLimit, "#,##0.00")
                AnnualMaster.SubItems(6) = Format(rst2!PSpecAnnLimit, "#,##0.00")
                AnnualMaster.SubItems(7) = Format(rst2!PMEPAnnLimit, "#,##0.00")
                AnnualMaster.SubItems(8) = Format(rst2!SuppLimit, "#,##0.00")
                AnnualMaster.SubItems(9) = Format(rst2!SDentalAnnLimit, "#,##0.00")
                AnnualMaster.SubItems(10) = Format(rst2!SSpecAnnLimit, "#,##0.00")
                AnnualMaster.SubItems(11) = Format(rst2!SMEPAnnLimit, "#,##0.00")
                AnnualMaster.SubItems(12) = IIf(Len(Trim(rst2!AnnEffDate)) >= 1, rst2!AnnEffDate, "")
                If Len(rst2!AnnLimitVersion) Then
                   AnnualMaster.SubItems(13) = rst2!AnnLimitVersion
                End If
                If Len(rst2!PAYCode) Then
                    AnnualMaster.SubItems(14) = rst2!PAYCode
                End If
                If Len(rst2!GRPCompany) Then
                   AnnualMaster.SubItems(15) = rst2!GRPCompany
                End If
                If Len(rst2!SuppLimitStatus) Then
                   AnnualMaster.SubItems(16) = rst2!SuppLimitStatus
                End If
                
                rst2.MoveNext
        TotalAnnualCount = TotalAnnualCount + 1
        txtTotalRecord = TotalAnnualCount
        Loop
        
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub AnnualComboListing()

    Dim INS As New ADODB.Recordset
    Dim HLT As New ADODB.Recordset
    Dim PLN As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    Dim GrpCom As New ADODB.Recordset
    
    PLN.Open "select PLNCode, PLNDesc  from PLN", MediConnCISDB, adOpenKeyset, adLockOptimistic
    INS.Open "select INSCOde, INSDesc from INS", MediConnCISDB, adOpenKeyset, adLockOptimistic
    HLT.Open "select HLTCOde, HLTDesc from HLT", MediConnCISDB, adOpenKeyset, adLockOptimistic
    PAY.Open "select PAYCOde, PAYCompanyName from PAY", MediConnCISDB, adOpenKeyset, adLockOptimistic
    GrpCom.Open "select Distinct GRPCompany from PLN ", MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    CboPlnCode.Clear
    CboHltCode.Clear
    CboInsType.Clear
    CboPayorCode.Clear
    CboGrpCom.Clear
    Combo1.Clear
    
    Do Until INS.EOF
        CboInsType.AddItem INS!INSCode & " - " & INS!InsDesc
        INS.MoveNext
    Loop
    Do Until HLT.EOF
        CboHltCode.AddItem HLT!HLTCode & " - " & HLT!HLTDesc
        HLT.MoveNext
    Loop
    Do Until PLN.EOF
        CboPlnCode.AddItem PLN!PLNCode & " - " & PLN!PlnDesc
        PLN.MoveNext
    Loop
    Do Until PAY.EOF
        CboPayorCode.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
        PAY.MoveNext
    Loop
    Do Until GrpCom.EOF
        If Len(GrpCom!GRPCompany) Then
            'If (GRPCOM!GRPCompany) <> CboGrpCom.ListIndex Then
            'GroupBy (GRPCOM!GRPCompany)
           CboGrpCom.AddItem GrpCom!GRPCompany
           
        End If
        'End If
        GrpCom.MoveNext
    Loop
    
    Combo1.AddItem "A" & " - " & "Share Limit For Family"
    Combo1.AddItem "B" & " - " & "Share Limit For Supplementary"
    Combo1.AddItem "C" & " - " & "Separate Limit For Supplementary"
    Combo1.Text = "A" & " - " & "Share Limit For Family"
        
    PLN.Close
    Set PLN = Nothing
    INS.Close
    Set INS = Nothing
    HLT.Close
    Set HLT = Nothing
    PAY.Close
    Set PAY = Nothing
    GrpCom.Close
    Set GrpCom = Nothing
End Sub

Private Sub VisitComboListing()
Dim INS As New ADODB.Recordset
Dim HLT As New ADODB.Recordset
Dim PLN As New ADODB.Recordset
Dim PAY As New ADODB.Recordset
Dim GrpCom As New ADODB.Recordset
    
    PLN.Open "select PLNCode, PLNDesc  from PLN", MediConnCISDB, adOpenKeyset, adLockOptimistic
    INS.Open "select INSCOde, INSDesc from INS", MediConnCISDB, adOpenKeyset, adLockOptimistic
    HLT.Open "select HLTCOde, HLTDesc from HLT", MediConnCISDB, adOpenKeyset, adLockOptimistic
    PAY.Open "select PAYCOde, PAYCompanyName from PAY", MediConnCISDB, adOpenKeyset, adLockOptimistic
    GrpCom.Open "select Distinct GRPCompany from PLN ", MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    CboVPLNCode.Clear
    CboVHltCode.Clear
    CboVINSCode.Clear
    CboVPAYCode.Clear
    CboVGrpCom.Clear
    
    Do Until INS.EOF
        CboVINSCode.AddItem INS!INSCode & " - " & INS!InsDesc
        INS.MoveNext
    Loop
    Do Until HLT.EOF
        CboVHltCode.AddItem HLT!HLTCode & " - " & HLT!HLTDesc
        HLT.MoveNext
    Loop
    Do Until PLN.EOF
        CboVPLNCode.AddItem PLN!PLNCode & " - " & PLN!PlnDesc
        PLN.MoveNext
    Loop
    Do Until PAY.EOF
        CboVPAYCode.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
        PAY.MoveNext
    Loop
    Do Until GrpCom.EOF
        If Len(GrpCom!GRPCompany) Then
            'If (GRPCOM!GRPCompany) <> CboGrpCom.ListIndex Then
            'GroupBy (GRPCOM!GRPCompany)
           CboVGrpCom.AddItem GrpCom!GRPCompany
           
        End If
        'End If
        GrpCom.MoveNext
    Loop
    
    CboVSuppStatus.AddItem "A" & " - " & "Share Limit For Family"
    CboVSuppStatus.AddItem "B" & " - " & "Share Limit For Supplementary"
    CboVSuppStatus.AddItem "C" & " - " & "Separate Limit For Supplementary"
    CboVSuppStatus.Text = "A" & " - " & "Share Limit For Family"
        
    PLN.Close
    Set PLN = Nothing
    INS.Close
    Set INS = Nothing
    HLT.Close
    Set HLT = Nothing
    PAY.Close
    Set PAY = Nothing
    GrpCom.Close
    Set GrpCom = Nothing
End Sub

Private Sub lstAnnualList_KeyDown(KeyCode As Integer, Shift As Integer)

    If lstAnnualList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboHltCode = lstAnnualList.SelectedItem.Text
        CboPlnCode = lstAnnualList.SelectedItem.SubItems(2)
        CboInsType = lstAnnualList.SelectedItem.SubItems(3)
        TxtAnnualLimit = lstAnnualList.SelectedItem.SubItems(4)
        TxtPDentalAnnLmt = lstAnnualList.SelectedItem.SubItems(5)
        TxtPSpecAnnLmt = lstAnnualList.SelectedItem.SubItems(6)
        TxtPMEPAnnLmt = lstAnnualList.SelectedItem.SubItems(7)
        TxtSupplyAL = lstAnnualList.SelectedItem.SubItems(8)
        TxtSDentalAnnLmt = lstAnnualList.SelectedItem.SubItems(9)
        TxtSSpecAnnLmt = lstAnnualList.SelectedItem.SubItems(10)
        TxtSMEPAnnLmt = lstAnnualList.SelectedItem.SubItems(11)
        txtALEffDate = lstAnnualList.SelectedItem.SubItems(12)
        txtANNVersion = lstAnnualList.SelectedItem.SubItems(13)
        CboPayorCode = lstAnnualList.SelectedItem.SubItems(14)
        CboGrpCom = lstAnnualList.SelectedItem.SubItems(15)
        
        
        If lstAnnualList.SelectedItem.SubItems(16) = "A" Then
            Combo1.ListIndex = 0
        ElseIf lstAnnualList.SelectedItem.SubItems(16) = "B" Then
            Combo1.ListIndex = 1
        ElseIf lstAnnualList.SelectedItem.SubItems(16) = "C" Then
            Combo1.ListIndex = 2
        End If
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstAnnualList_KeyUp(KeyCode As Integer, Shift As Integer)

    If lstAnnualList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboHltCode = lstAnnualList.SelectedItem.Text
        CboPlnCode = lstAnnualList.SelectedItem.SubItems(2)
        CboInsType = lstAnnualList.SelectedItem.SubItems(3)
        TxtAnnualLimit = lstAnnualList.SelectedItem.SubItems(4)
        TxtPDentalAnnLmt = lstAnnualList.SelectedItem.SubItems(5)
        TxtPSpecAnnLmt = lstAnnualList.SelectedItem.SubItems(6)
        TxtPMEPAnnLmt = lstAnnualList.SelectedItem.SubItems(7)
        TxtSupplyAL = lstAnnualList.SelectedItem.SubItems(8)
        TxtSDentalAnnLmt = lstAnnualList.SelectedItem.SubItems(9)
        TxtSSpecAnnLmt = lstAnnualList.SelectedItem.SubItems(10)
        TxtSMEPAnnLmt = lstAnnualList.SelectedItem.SubItems(11)
        txtALEffDate = lstAnnualList.SelectedItem.SubItems(12)
        txtANNVersion = lstAnnualList.SelectedItem.SubItems(13)
        CboPayorCode = lstAnnualList.SelectedItem.SubItems(14)
        CboGrpCom = lstAnnualList.SelectedItem.SubItems(15)
        
        
        If lstAnnualList.SelectedItem.SubItems(16) = "A" Then
            Combo1.ListIndex = 0
        ElseIf lstAnnualList.SelectedItem.SubItems(16) = "B" Then
            Combo1.ListIndex = 1
        ElseIf lstAnnualList.SelectedItem.SubItems(16) = "C" Then
            Combo1.ListIndex = 2
        End If
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstAnnualList_Click()

    If lstAnnualList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboHltCode = lstAnnualList.SelectedItem.Text
        CboPlnCode = lstAnnualList.SelectedItem.SubItems(2)
        CboInsType = lstAnnualList.SelectedItem.SubItems(3)
        TxtAnnualLimit = lstAnnualList.SelectedItem.SubItems(4)
        TxtPDentalAnnLmt = lstAnnualList.SelectedItem.SubItems(5)
        TxtPSpecAnnLmt = lstAnnualList.SelectedItem.SubItems(6)
        TxtPMEPAnnLmt = lstAnnualList.SelectedItem.SubItems(7)
        TxtSupplyAL = lstAnnualList.SelectedItem.SubItems(8)
        TxtSDentalAnnLmt = lstAnnualList.SelectedItem.SubItems(9)
        TxtSSpecAnnLmt = lstAnnualList.SelectedItem.SubItems(10)
        TxtSMEPAnnLmt = lstAnnualList.SelectedItem.SubItems(11)
        txtALEffDate = lstAnnualList.SelectedItem.SubItems(12)
        txtANNVersion = lstAnnualList.SelectedItem.SubItems(13)
        CboPayorCode = lstAnnualList.SelectedItem.SubItems(14)
        CboGrpCom = lstAnnualList.SelectedItem.SubItems(15)
        
        
        If lstAnnualList.SelectedItem.SubItems(16) = "A" Then
            Combo1.ListIndex = 0
        ElseIf lstAnnualList.SelectedItem.SubItems(16) = "B" Then
            Combo1.ListIndex = 1
        ElseIf lstAnnualList.SelectedItem.SubItems(16) = "C" Then
            Combo1.ListIndex = 2
        End If
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
End Sub
Private Sub lstAnnualList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstAnnualList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstAnnualList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstAnnualList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstAnnualList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstAnnualList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstAnnualList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstAnnualList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstAnnualList, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstAnnualList, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstAnnualList, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 10
        SortListView lstAnnualList, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
    Case 10
        SortListView lstAnnualList, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    Case 11
        SortListView lstAnnualList, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
    Case 12
        SortListView lstAnnualList, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
    Case 13
        SortListView lstAnnualList, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 14
        SortListView lstAnnualList, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 15
        SortListView lstAnnualList, ColumnHeader.Index, ldtString, m_blnDirection(17)
    End Select
End Sub

Private Sub PrintANN()
    RptAnnualList.Show
End Sub

Private Sub PrintVisit()
    RptVisitList.Show
End Sub
'############################### End Annual ###############################'

'############################## Start MCO ##############################'
Private Sub SaveMCO()
Dim RecordSaved
Dim tmpAdultAmt As Double
Dim tmpChildAmt As Double
Dim tmpSuppAdultAmt As Double
Dim tmpSuppChildAmt As Double
Dim tmpInsType As String
    
    If CboMCOHlt = "" Then
        MsgBox "Please Enter Health Code", vbOKOnly + vbCritical
        CboMCOHlt.SetFocus
    ElseIf CboMCOPln = "" Then
        MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical
        CboMCOPln.SetFocus
    ElseIf TxtIAAmt = "" And TxtICAmt = "" And TxtFAAmt = "" And TxtFCAmt = "" _
           And TxtGAAmt = "" And TxtGCAmt = "" And TxtHAAmt = "" And TxtHCAmt = "" Then
        MsgBox "Please Enter MCO Fees ", vbOKOnly + vbCritical
        TxtIAAmt.SetFocus
    ElseIf txtEffDate = "__/__/____" Or Not IsDate(txtEffDate) Then
        MsgBox "Please validate your MCO Effective date", vbOKOnly + vbCritical
        txtEffDate.SetFocus
    Else
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
' --------- Individual
            tmpInsType = "I"
            tmpAdultAmt = CDbl(IIf(IsNumeric(TxtIAAmt), TxtIAAmt, 0))
            tmpChildAmt = CDbl(IIf(IsNumeric(TxtICAmt), TxtICAmt, 0))
            tmpSuppAdultAmt = 0
            tmpSuppChildAmt = 0
            If tmpAdultAmt > 0 Or tmpChildAmt > 0 Then
                Call SaveMCO1(tmpInsType, tmpAdultAmt, tmpChildAmt, tmpSuppAdultAmt, tmpSuppChildAmt)
            End If
' --------- Family
            tmpInsType = "F"
            tmpAdultAmt = CDbl(IIf(IsNumeric(TxtFAAmt), TxtFAAmt, 0))
            tmpChildAmt = CDbl(IIf(IsNumeric(TxtFCAmt), TxtFCAmt, 0))
            tmpSuppAdultAmt = CDbl(IIf(IsNumeric(TxtSFAAmt), TxtSFAAmt, 0))
            tmpSuppChildAmt = CDbl(IIf(IsNumeric(TxtSFCAmt), TxtSFCAmt, 0))
            If tmpAdultAmt > 0 Or tmpChildAmt > 0 Or tmpSuppAdultAmt > 0 Or tmpSuppChildAmt > 0 Then
                Call SaveMCO1(tmpInsType, tmpAdultAmt, tmpChildAmt, tmpSuppAdultAmt, tmpSuppChildAmt)
            End If
' --------- Group Individual
            tmpInsType = "G"
            tmpAdultAmt = CDbl(IIf(IsNumeric(TxtGAAmt), TxtGAAmt, 0))
            tmpChildAmt = CDbl(IIf(IsNumeric(TxtGCAmt), TxtGCAmt, 0))
            tmpSuppAdultAmt = 0
            tmpSuppChildAmt = 0
            If tmpAdultAmt > 0 Or tmpChildAmt > 0 Then
                Call SaveMCO1(tmpInsType, tmpAdultAmt, tmpChildAmt, tmpSuppAdultAmt, tmpSuppChildAmt)
            End If
' --------- Group Family
            tmpInsType = "H"
            tmpAdultAmt = CDbl(IIf(IsNumeric(TxtHAAmt), TxtHAAmt, 0))
            tmpChildAmt = CDbl(IIf(IsNumeric(TxtHCAmt), TxtHCAmt, 0))
            tmpSuppAdultAmt = CDbl(IIf(IsNumeric(TxtSHAAmt), TxtSHAAmt, 0))
            tmpSuppChildAmt = CDbl(IIf(IsNumeric(TxtSHCAmt), TxtSHCAmt, 0))
            If tmpAdultAmt > 0 Or tmpChildAmt > 0 Or tmpSuppAdultAmt > 0 Or tmpSuppChildAmt > 0 Then
                Call SaveMCO1(tmpInsType, tmpAdultAmt, tmpChildAmt, tmpSuppAdultAmt, tmpSuppChildAmt)
            End If
            MsgBox "Record Saved ! ", vbOKOnly
        End If
        
    End If
End Sub

Private Sub DeleteMCO()

Dim DeleteMCO As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim StrSql, sqlMBM As String
Dim DeleteRecord

EditFlag = False
If CboMCOHlt = "" Then
   MsgBox "Please Enter Health Type", vbOKOnly + vbCritical, DialogBoxTitle
   'CboMCOHlt.SetFocus
ElseIf lstMCOList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBM where INSCode = '" & ChkStr(lstMCOList.SelectedItem.Text) & "'" & _
            " and MBMPayorEffdate >='" & Format(lstMCOList.SelectedItem.SubItems(6), "mm/dd/yyyy") & "'"
   
   rsCheckMBM.MaxRecords = 1
    rsCheckMBM.Open sqlMBM, MediConnCISDB, adOpenKeyset, adLockOptimistic
   
     
   If rsCheckMBM.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.RecordCount = 0 Then
           StrSql = "Delete From MCO Where MCOcode = '" & ChkStr(lstMCOList.SelectedItem.Text) & "'"
           DeleteMCO.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call MCOListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtINScode.SetFocus
         End If
     End If
End If
Exit Sub
End Sub

Private Sub SearchMCO()
    Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(CboMCOHlt)) Then
        strAppend = strAppend + " AND HLTCode = '" & Mid(CboMCOHlt, 1, 1) & "'"
    End If
    
    If Len(Trim(CboMCOPln)) Then
        strAppend = strAppend + " AND PLNCode = '" & Trim(CboMCOPln) & "'"
    End If
    

    If txtEffDate <> "__/__/____" Then
        strAppend = strAppend + " AND MCOEffDate = '" & Format(CDate(txtEffDate), "MM/DD/YYYY") & "'"
    End If
        
    MCOCriteria = strAppend
    Call MCOListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub MCOListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
    Dim MCOMaster As ListItem
    Dim sql As String
    Dim TotalMCOCount As Integer
    
    TotalMCOCount = 0
    lstMCOList.ListItems.Clear
    sql = "Select * from MCO where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
    sql = sql + strAppend
    End If

    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found", vbOKOnly + vbInformation
        
    Else
        Do Until rst2.EOF
            Set MCOMaster = lstMCOList.ListItems.Add(, , rst2!HLTCode)
                MCOMaster.SubItems(1) = Trim(rst2!INSCode)
                If Len(rst2!PLNCode) Then
                    MCOMaster.SubItems(2) = Trim(rst2!PLNCode)
                End If
                If Len(rst2!SuppMCOStatus) Then
                    MCOMaster.SubItems(3) = rst2!SuppMCOStatus
                End If
                MCOMaster.SubItems(4) = Format(rst2!MCOAmtAdult, "#,##0.00")
                MCOMaster.SubItems(5) = Format(rst2!MCOAmtChild, "#,##0.00")
                MCOMaster.SubItems(6) = Format(rst2!MCOSuppAdultAmt, "#,##0.00")
                MCOMaster.SubItems(7) = Format(rst2!MCOSuppChildAmt, "#,##0.00")
                MCOMaster.SubItems(8) = IIf(Len(Trim(rst2!MCOEffDate)) >= 1, rst2!MCOEffDate, "")
                rst2.MoveNext
        TotalMCOCount = TotalMCOCount + 1
        txtTotalRecord = TotalMCOCount
        Loop
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub MCOComboListing()
Dim HLT As New ADODB.Recordset
Dim PLN As New ADODB.Recordset
    
    HLT.Open "HLT", MediConnCISDB, adOpenKeyset, adLockOptimistic
    PLN.Open "PLN", MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    CboMCOHlt.Clear
    CboMCOPln.Clear
    
    Do Until HLT.EOF
        CboMCOHlt.AddItem HLT!HLTCode & " - " & HLT!HLTDesc
        HLT.MoveNext
    Loop
    
    Do Until PLN.EOF
        CboMCOPln.AddItem PLN!PLNCode '& " - " & PLN!PLNDesc
        PLN.MoveNext
    Loop

    
    cboSuppMCOStatus.AddItem "A" & " - " & "Share Limit for Family"
    cboSuppMCOStatus.AddItem "B" & " - " & "Share Limit for Supplementary"
    cboSuppMCOStatus.AddItem "C" & " - " & "Separate Limit for Supplementary"
    cboSuppMCOStatus.Text = "C" & " - " & "Separate Limit for Supplementary"
    
    HLT.Close
    PLN.Close
End Sub

Private Sub lstMCOList_Click()
Dim MCO As New ADODB.Recordset
Dim StrSql As String
Dim Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Amt7, Amt8, Amt9, Amt10, Amt11, Amt12 As String
   
   Amt1 = 0
   Amt2 = 0
   Amt3 = 0
   Amt4 = 0
   Amt5 = 0
   Amt6 = 0
   Amt7 = 0
   Amt8 = 0
   Amt9 = 0
   Amt10 = 0
   Amt11 = 0
   Amt12 = 0
   
   If lstMCOList.ListItems.Count <> 0 Then
        Call ClearForm(Me)
        
      '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        
        StrSql = " Select * From MCO where PLNCode = '" & lstMCOList.SelectedItem.SubItems(2) & "'" & _
                 " AND INSCode = '" & lstMCOList.SelectedItem.SubItems(1) & "' AND SuppMCOStatus = '" & lstMCOList.SelectedItem.SubItems(3) & "'" & _
                 " AND MCOEffDate='" & Format(lstMCOList.SelectedItem.SubItems(8), "MM/DD/YYYY") & "'"
                 
        MCO.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        Do Until MCO.EOF
            If MCO!INSCode = "I" Then
                Amt1 = MCO!MCOAmtAdult
                Amt2 = MCO!MCOAmtChild
            ElseIf MCO!INSCode = "F" Then
                Amt3 = MCO!MCOAmtAdult
                Amt4 = MCO!MCOAmtChild
                Amt9 = MCO!MCOSuppAdultAmt
                Amt10 = MCO!MCOSuppChildAmt
            ElseIf MCO!INSCode = "G" Then
                Amt5 = MCO!MCOAmtAdult
                Amt6 = MCO!MCOAmtChild
            ElseIf MCO!INSCode = "H" Then
                Amt7 = MCO!MCOAmtAdult
                Amt8 = MCO!MCOAmtChild
                Amt11 = MCO!MCOSuppAdultAmt
                Amt12 = MCO!MCOSuppChildAmt
            End If
            MCO.MoveNext
            Loop
            MCO.Close
            Set MCO = Nothing
    
                
        CboMCOHlt = lstMCOList.SelectedItem.Text
        CboMCOPln = lstMCOList.SelectedItem.SubItems(2)
        If lstMCOList.SelectedItem.SubItems(3) = "A" Then
            cboSuppMCOStatus.ListIndex = 0
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "B" Then
            cboSuppMCOStatus.ListIndex = 1
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "C" Then
            cboSuppMCOStatus.ListIndex = 2
        End If
        
        TxtIAAmt = Format(Amt1, "#,##0.00")
        TxtICAmt = Format(Amt2, "#,##0.00")
        TxtFAAmt = Format(Amt3, "#,##0.00")
        TxtFCAmt = Format(Amt4, "#,##0.00")
        TxtSFAAmt = Format(Amt9, "#,##0.00")
        TxtSFCAmt = Format(Amt10, "#,##0.00")
        TxtGAAmt = Format(Amt5, "#,##0.00")
        TxtGCAmt = Format(Amt6, "#,##0.00")
        TxtHAAmt = Format(Amt7, "#,##0.00")
        TxtHCAmt = Format(Amt8, "#,##0.00")
        TxtSHAAmt = Format(Amt11, "#,##0.00")
        TxtSHCAmt = Format(Amt12, "#,##0.00")
        txtEffDate = lstMCOList.SelectedItem.SubItems(8)
        
     '   MediConnCISDB.Close
      '  Set MediConnCISDB = Nothing
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
End Sub

Private Sub lstMCOList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim MCO As New ADODB.Recordset
Dim StrSql As String
Dim Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Amt7, Amt8, Amt9, Amt10, Amt11, Amt12 As String
   
   Amt1 = 0
   Amt2 = 0
   Amt3 = 0
   Amt4 = 0
   Amt5 = 0
   Amt6 = 0
   Amt7 = 0
   Amt8 = 0
   Amt9 = 0
   Amt10 = 0
   Amt11 = 0
   Amt12 = 0
   
   If lstMCOList.ListItems.Count <> 0 Then
        Call ClearForm(Me)
        
        StrSql = " Select * From MCO where PLNCode = '" & lstMCOList.SelectedItem.SubItems(2) & "'" & _
                 " AND INSCode = '" & lstMCOList.SelectedItem.SubItems(1) & "' AND SuppMCOStatus = '" & lstMCOList.SelectedItem.SubItems(3) & "'"
        MCO.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        Do Until MCO.EOF
            If MCO!INSCode = "I" Then
                Amt1 = MCO!MCOAmtAdult
                Amt2 = MCO!MCOAmtChild
            ElseIf MCO!INSCode = "F" Then
                Amt3 = MCO!MCOAmtAdult
                Amt4 = MCO!MCOAmtChild
                Amt9 = MCO!MCOSuppAdultAmt
                Amt10 = MCO!MCOSuppChildAmt
            ElseIf MCO!INSCode = "G" Then
                Amt5 = MCO!MCOAmtAdult
                Amt6 = MCO!MCOAmtChild
            ElseIf MCO!INSCode = "H" Then
                Amt7 = MCO!MCOAmtAdult
                Amt8 = MCO!MCOAmtChild
                Amt11 = MCO!MCOSuppAdultAmt
                Amt12 = MCO!MCOSuppChildAmt
            End If
            MCO.MoveNext
            Loop
            MCO.Close
            Set MCO = Nothing
    
                
        CboMCOHlt = lstMCOList.SelectedItem.Text
        CboMCOPln = lstMCOList.SelectedItem.SubItems(2)
        If lstMCOList.SelectedItem.SubItems(3) = "A" Then
            cboSuppMCOStatus.ListIndex = 0
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "B" Then
            cboSuppMCOStatus.ListIndex = 1
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "C" Then
            cboSuppMCOStatus.ListIndex = 2
        End If
        
        TxtIAAmt = Format(Amt1, "#,##0.00")
        TxtICAmt = Format(Amt2, "#,##0.00")
        TxtFAAmt = Format(Amt3, "#,##0.00")
        TxtFCAmt = Format(Amt4, "#,##0.00")
        TxtSFAAmt = Format(Amt9, "#,##0.00")
        TxtSFCAmt = Format(Amt10, "#,##0.00")
        TxtGAAmt = Format(Amt5, "#,##0.00")
        TxtGCAmt = Format(Amt6, "#,##0.00")
        TxtHAAmt = Format(Amt7, "#,##0.00")
        TxtHCAmt = Format(Amt8, "#,##0.00")
        TxtSHAAmt = Format(Amt11, "#,##0.00")
        TxtSHCAmt = Format(Amt12, "#,##0.00")
        txtEffDate = lstMCOList.SelectedItem.SubItems(8)
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If

End Sub

Private Sub lstMCOList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim MCO As New ADODB.Recordset
Dim StrSql As String
Dim Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Amt7, Amt8, Amt9, Amt10, Amt11, Amt12 As String
   
   Amt1 = 0
   Amt2 = 0
   Amt3 = 0
   Amt4 = 0
   Amt5 = 0
   Amt6 = 0
   Amt7 = 0
   Amt8 = 0
   Amt9 = 0
   Amt10 = 0
   Amt11 = 0
   Amt12 = 0
   
   If lstMCOList.ListItems.Count <> 0 Then
        Call ClearForm(Me)
        
        StrSql = " Select * From MCO where PLNCode = '" & lstMCOList.SelectedItem.SubItems(2) & "'" & _
                 " AND INSCode = '" & lstMCOList.SelectedItem.SubItems(1) & "' AND SuppMCOStatus = '" & lstMCOList.SelectedItem.SubItems(3) & "'"
        MCO.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        Do Until MCO.EOF
            If MCO!INSCode = "I" Then
                Amt1 = MCO!MCOAmtAdult
                Amt2 = MCO!MCOAmtChild
            ElseIf MCO!INSCode = "F" Then
                Amt3 = MCO!MCOAmtAdult
                Amt4 = MCO!MCOAmtChild
                Amt9 = MCO!MCOSuppAdultAmt
                Amt10 = MCO!MCOSuppChildAmt
            ElseIf MCO!INSCode = "G" Then
                Amt5 = MCO!MCOAmtAdult
                Amt6 = MCO!MCOAmtChild
            ElseIf MCO!INSCode = "H" Then
                Amt7 = MCO!MCOAmtAdult
                Amt8 = MCO!MCOAmtChild
                Amt11 = MCO!MCOSuppAdultAmt
                Amt12 = MCO!MCOSuppChildAmt
            End If
            MCO.MoveNext
            Loop
            MCO.Close
            Set MCO = Nothing
    
                
        CboMCOHlt = lstMCOList.SelectedItem.Text
        CboMCOPln = lstMCOList.SelectedItem.SubItems(2)
        If lstMCOList.SelectedItem.SubItems(3) = "A" Then
            cboSuppMCOStatus.ListIndex = 0
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "B" Then
            cboSuppMCOStatus.ListIndex = 1
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "C" Then
            cboSuppMCOStatus.ListIndex = 2
        End If
        
        TxtIAAmt = Format(Amt1, "#,##0.00")
        TxtICAmt = Format(Amt2, "#,##0.00")
        TxtFAAmt = Format(Amt3, "#,##0.00")
        TxtFCAmt = Format(Amt4, "#,##0.00")
        TxtSFAAmt = Format(Amt9, "#,##0.00")
        TxtSFCAmt = Format(Amt10, "#,##0.00")
        TxtGAAmt = Format(Amt5, "#,##0.00")
        TxtGCAmt = Format(Amt6, "#,##0.00")
        TxtHAAmt = Format(Amt7, "#,##0.00")
        TxtHCAmt = Format(Amt8, "#,##0.00")
        TxtSHAAmt = Format(Amt11, "#,##0.00")
        TxtSHCAmt = Format(Amt12, "#,##0.00")
        txtEffDate = lstMCOList.SelectedItem.SubItems(8)
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If

End Sub

Private Sub lstMCOList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstMCOList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstMCOList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstMCOList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstMCOList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstMCOList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstMCOList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstMCOList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstMCOList, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstMCOList, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    End Select
End Sub
Private Sub PrintMCO()
    RptMCOList.Show
End Sub

Private Sub PrintCoPaymentMBM()
    RptCOPayMBM.Show
End Sub

'############################### End MCO ###############################'

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
    Clipboard.Clear
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

Private Sub CboVGrpCom_Change()
        If InStr(CboVGrpCom.Text, "null") Then
       CboVGrpCom.Enabled = False
    End If
End Sub

Private Sub CboVGrpCom_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboVGrpCom.Text, CboVGrpCom.SelStart) + Chr(KeyAscii) + Mid(CboVGrpCom.Text, CboVGrpCom.SelStart + CboVGrpCom.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboVGrpCom.ListCount - 1
 
        If NewText = LCase(Left(CboVGrpCom.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboVGrpCom.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboVGrpCom.Text = ValidValue
               CboVGrpCom.SelStart = 0
               CboVGrpCom.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboHLTCode_Change()
Dim PAY As New ADODB.Recordset
Dim StrSql As String
  
'    strsql = "Select  PAYCode, PAYCompanyName From PAY "
'    strsql = strsql + " Where HLTCode ='" & Mid(Trim(CboHltCode), 1, 1) & "'"
    
'    PAY.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
'    CboHltCode.Clear
'    Do Until PAY.EOF
'       CboHltCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
'       PAY.MoveNext
'    Loop
'    PAY.Close
'    Set PAY = Nothing
    
    If InStr(CboHltCode.Text, "null") Then
       CboHltCode.Enabled = False
    End If
   
   
End Sub

Private Sub CboHLTCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboHltCode.Text, CboHltCode.SelStart) + Chr(KeyAscii) + Mid(CboHltCode.Text, CboHltCode.SelStart + CboHltCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboHltCode.ListCount - 1
 
        If NewText = LCase(Left(CboHltCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboHltCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboHltCode.Text = ValidValue
               CboHltCode.SelStart = 0
               CboHltCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboVHLTCode_Change()
Dim PAY As New ADODB.Recordset
Dim StrSql As String

    If InStr(CboVHltCode.Text, "null") Then
       CboVHltCode.Enabled = False
    End If
   
End Sub

Private Sub cboVHLTcode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboVHltCode.Text, CboVHltCode.SelStart) + Chr(KeyAscii) + Mid(CboVHltCode.Text, CboVHltCode.SelStart + CboVHltCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboVHltCode.ListCount - 1
 
        If NewText = LCase(Left(CboVHltCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboVHltCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboVHltCode.Text = ValidValue
               CboVHltCode.SelStart = 0
               CboVHltCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboInsType_Change()
    If InStr(CboInsType.Text, "null") Then
       CboInsType.Enabled = False
    End If
End Sub

Private Sub CboInsType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboInsType.Text, CboInsType.SelStart) + Chr(KeyAscii) + Mid(CboInsType.Text, CboInsType.SelStart + CboInsType.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboInsType.ListCount - 1
 
        If NewText = LCase(Left(CboInsType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboInsType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboInsType.Text = ValidValue
               CboInsType.SelStart = 0
               CboInsType.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboVINSCode_Change()
    If InStr(CboVINSCode.Text, "null") Then
       CboVINSCode.Enabled = False
    End If
End Sub

Private Sub CboVINSCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboVINSCode.Text, CboVINSCode.SelStart) + Chr(KeyAscii) + Mid(CboVINSCode.Text, CboVINSCode.SelStart + CboVINSCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboVINSCode.ListCount - 1
 
        If NewText = LCase(Left(CboVINSCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboVINSCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboVINSCode.Text = ValidValue
               CboVINSCode.SelStart = 0
               CboVINSCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboMCOHlt_Change()
   
    If InStr(CboMCOHlt.Text, "null") Then
       CboMCOHlt.Enabled = False
    End If

End Sub

Private Sub CboMCOHlt_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboMCOHlt.Text, CboMCOHlt.SelStart) + Chr(KeyAscii) + Mid(CboMCOHlt.Text, CboMCOHlt.SelStart + CboMCOHlt.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboMCOHlt.ListCount - 1
 
        If NewText = LCase(Left(CboMCOHlt.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboMCOHlt.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboMCOHlt.Text = ValidValue
               CboMCOHlt.SelStart = 0
               CboMCOHlt.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPLNCode_Change()
    If InStr(CboPlnCode.Text, "null") Then
       CboPlnCode.Enabled = False
    End If
End Sub

Private Sub cboPLNCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPlnCode.Text, CboPlnCode.SelStart) + Chr(KeyAscii) + Mid(CboPlnCode.Text, CboPlnCode.SelStart + CboPlnCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPlnCode.ListCount - 1
 
        If NewText = LCase(Left(CboPlnCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPlnCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboPlnCode.Text = ValidValue
               CboPlnCode.SelStart = 0
               CboPlnCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboVPLNCode_Change()
    If InStr(CboVPLNCode.Text, "null") Then
       CboVPLNCode.Enabled = False
    End If
End Sub

Private Sub cboVPLNCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboVPLNCode.Text, CboVPLNCode.SelStart) + Chr(KeyAscii) + Mid(CboVPLNCode.Text, CboVPLNCode.SelStart + CboVPLNCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboVPLNCode.ListCount - 1
 
        If NewText = LCase(Left(CboVPLNCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboVPLNCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboVPLNCode.Text = ValidValue
               CboVPLNCode.SelStart = 0
               CboVPLNCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPLNHLTCode_Change()
Dim PAY As New ADODB.Recordset
Dim StrSql As String
    
    cboPLNPayor.Clear
        
    StrSql = "Select  PAYCode, PAYCompanyName From PAY "
    StrSql = StrSql + " Where HLTCode ='" & Mid(Trim(cboPLNHLTCode), 1, 1) & "'"
    
    PAY.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    cboPLNPayor.Clear
    Do Until PAY.EOF
       cboPLNPayor.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
    If InStr(cboPLNHLTCode.Text, "null") Then
       cboPLNHLTCode.Enabled = False
    End If
    
    If Mid(Trim(cboPLNHLTCode), 1, 1) = "N" Then
        txtPLN1.Enabled = False
        txtPLN2.Visible = False
        txtPLN3.Visible = False
        txtPLN4.Visible = False
        txtPLN5.Visible = False
        txtPLN6.Visible = False
    Else
        txtPLN1.Enabled = True
        txtPLN2.Visible = True
        txtPLN3.Visible = True
        txtPLN4.Visible = True
        txtPLN5.Visible = True
        txtPLN6.Visible = True
    End If
    
    
End Sub

Private Sub cboPLNHLTCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPLNHLTCode.Text, cboPLNHLTCode.SelStart) + Chr(KeyAscii) + Mid(cboPLNHLTCode.Text, cboPLNHLTCode.SelStart + cboPLNHLTCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPLNHLTCode.ListCount - 1
 
        If NewText = LCase(Left(cboPLNHLTCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPLNHLTCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboPLNHLTCode.Text = ValidValue
               cboPLNHLTCode.SelStart = 0
               cboPLNHLTCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub


Private Sub cboPayorCode_Change()
    If InStr(CboPayorCode.Text, "null") Then
       CboPayorCode.Enabled = False
    End If
End Sub

Private Sub cboPayorCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayorCode.Text, CboPayorCode.SelStart) + Chr(KeyAscii) + Mid(CboPayorCode.Text, CboPayorCode.SelStart + CboPayorCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPayorCode.ListCount - 1
 
        If NewText = LCase(Left(CboPayorCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayorCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboPayorCode.Text = ValidValue
               CboPayorCode.SelStart = 0
               CboPayorCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboVPAYCode_Change()
    If InStr(CboVPAYCode.Text, "null") Then
       CboVPAYCode.Enabled = False
    End If
End Sub

Private Sub CboVPAYCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboVPAYCode.Text, CboVPAYCode.SelStart) + Chr(KeyAscii) + Mid(CboVPAYCode.Text, CboVPAYCode.SelStart + CboVPAYCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboVPAYCode.ListCount - 1
 
        If NewText = LCase(Left(CboVPAYCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboVPAYCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboVPAYCode.Text = ValidValue
               CboVPAYCode.SelStart = 0
               CboVPAYCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPLNHLTCode_Click()
Dim PAY As New ADODB.Recordset
Dim StrSql As String
    
    cboPLNPayor.Clear
    
  '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    
    StrSql = "Select  PAYCode, PAYCompanyName From PAY "
    StrSql = StrSql + " Where HLTCode ='" & Mid(Trim(cboPLNHLTCode), 1, 1) & "'"
    
    PAY.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    cboPLNPayor.Clear
    Do Until PAY.EOF
       cboPLNPayor.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
    If InStr(cboPLNHLTCode.Text, "null") Then
       cboPLNHLTCode.Enabled = False
    End If
    
    If Mid(Trim(cboPLNHLTCode), 1, 1) = "N" Then
        txtPLN1.Enabled = False
        txtPLN2.Visible = False
        txtPLN3.Visible = False
        txtPLN4.Visible = False
        txtPLN5.Visible = False
        txtPLN6.Visible = False
    Else
        txtPLN1.Enabled = True
        txtPLN2.Visible = True
        txtPLN3.Visible = True
        txtPLN4.Visible = True
        txtPLN5.Visible = True
        txtPLN6.Visible = True
    End If
  '  MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing
End Sub

Private Sub cboPmHLTCode_Change()
Dim PAY As New ADODB.Recordset
Dim StrSql As String
    
    cboPmPAYCode.Clear
    

    StrSql = "Select  PAYCode, PAYCompanyName From PAY "
    StrSql = StrSql + " Where HLTCode ='" & Mid(Trim(cboPmHLTCode), 1, 1) & "'"
    
    PAY.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    cboPmPAYCode.Clear
    Do Until PAY.EOF
       cboPmPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
  
    If InStr(cboPmHLTCode.Text, "null") Then
       CboMCOHlt.Enabled = False
    End If
End Sub

Private Sub cboPmHLTCode_Click()
Dim PAY As New ADODB.Recordset
Dim StrSql As String
    
    cboPmPAYCode.Clear
    
    StrSql = "Select  PAYCode, PAYCompanyName From PAY "
    StrSql = StrSql + " Where HLTCode ='" & Mid(Trim(cboPmHLTCode), 1, 1) & "'"
    
    PAY.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    cboPmPAYCode.Clear
    Do Until PAY.EOF
       cboPmPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
End Sub

Private Sub cboPmHLTCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPmHLTCode.Text, cboPmHLTCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboPmHLTCode.Text, cboPmHLTCode.SelStart + cboPmHLTCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPmHLTCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboPmHLTCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPmHLTCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
              cboPmHLTCode.Text = ValidValue
              cboPmHLTCode.SelStart = 0
              cboPmHLTCode.SelLength = Len(ValidValue)
            End If
   End If

End Sub

Private Sub Combo3_Change()
    If Mid(Combo3, 1, 1) = "B" Or Mid(Combo3, 1, 1) = "C" Then
        Frame6.Enabled = True
    Else
        Frame6.Enabled = False
    End If
End Sub

Private Sub Combo3_Click()
    If Mid(Combo3, 1, 1) = "B" Or Mid(Combo3, 1, 1) = "C" Then
        Frame6.Enabled = True
    Else
        Frame6.Enabled = False
    End If
End Sub

Private Sub Check1_Click()
Dim AgeBand As String
    If Check1.Value = 1 Then
        If Text2 <> "" Then
            AgeBand = Text2
            Text3 = AgeBand
            Text4 = AgeBand
            Text5 = AgeBand
            Text6 = AgeBand
            Text7 = AgeBand
            Text1 = AgeBand
        End If
    End If

End Sub

Private Sub chkSTDPremium_Click()
Dim AgeBand As String
    If chkSTDPremium.Value = 1 Then
        If txt01 <> "" Then
            AgeBand = txt01
            txt02 = AgeBand
            txt03 = AgeBand
            txt04 = AgeBand
            txt05 = AgeBand
            txt06 = AgeBand
            txt07 = AgeBand
        End If
    End If

End Sub

Private Sub Combo1_Change()
    If Mid(Combo1, 1, 1) = "B" Or Mid(Combo1, 1, 1) = "C" Then
        TxtSupplyAL.Enabled = True
    Else
        TxtSupplyAL.Enabled = False
    End If
End Sub

Private Sub CboVSuppStatus_Click()
    If Mid(CboVSuppStatus, 1, 1) = "B" Or Mid(CboVSuppStatus, 1, 1) = "C" Then
        TxtVSuppVisit.Enabled = True
    Else
        TxtVSuppVisit.Enabled = False
    End If
End Sub

Private Sub CboVSuppStatus_Change()
    If Mid(CboVSuppStatus, 1, 1) = "B" Or Mid(CboVSuppStatus, 1, 1) = "C" Then
        TxtVSuppVisit.Enabled = True
    Else
        TxtVSuppVisit.Enabled = False
    End If
End Sub

Private Sub Combo1_Click()
    If Mid(Combo1, 1, 1) = "B" Or Mid(Combo1, 1, 1) = "C" Then
        TxtSupplyAL.Enabled = True
    Else
        TxtSupplyAL.Enabled = False
    End If
End Sub

Private Sub Text2_Change()
Dim AgeBand As String
    If Check1.Value = 1 Then
        If Text2 <> "" Then
            AgeBand = Text2
            Text3 = AgeBand
            Text4 = AgeBand
            Text5 = AgeBand
            Text6 = AgeBand
            Text7 = AgeBand
            Text1 = AgeBand
        ElseIf Text2 = "" Then
            Text3 = ""
            Text4 = ""
            Text5 = ""
            Text6 = ""
            Text7 = ""
            Text1 = ""
        End If
    End If
End Sub

Private Sub txt01_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txt02_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txt03_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txt04_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txt05_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txt06_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txt07_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txt01_Change()
Dim AgeBand As String
    If chkSTDPremium.Value = 1 Then
        If txt01 <> "" Then
            AgeBand = txt01
            txt02 = AgeBand
            txt03 = AgeBand
            txt04 = AgeBand
            txt05 = AgeBand
            txt06 = AgeBand
            txt07 = AgeBand
        ElseIf txt01 = "" Then
            txt01 = ""
            txt02 = ""
            txt03 = ""
            txt04 = ""
            txt05 = ""
            txt06 = ""
            txt07 = ""
        End If
    End If
End Sub

Private Sub txtEffDate_LostFocus()
    If txtEffDate <> "__/__/____" Then
        If IsDate(txtEffDate) = False Then
            MsgBox "Invalid date format!", vbCritical
        End If
    End If
End Sub

Private Sub txtALEffDate_LostFocus()
    If IsDate(txtALEffDate) Then
    Else
        If txtALEffDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtALEffDate.SetFocus
        End If
    End If
End Sub

Private Sub TxtPDentalAnnLmt_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtPMEPAnnLmt_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtPSpecAnnLmt_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtSDentalAnnLmt_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtSMEPAnnLmt_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtSSpecAnnLmt_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtVisitLimit_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtVLimitEffDate_LostFocus()
    If IsDate(TxtVLimitEffDate) Then
    Else
        If TxtVLimitEffDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtVLimitEffDate.SetFocus
        End If
    End If
End Sub


'Private Sub TxtGenWP1_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub


'Private Sub TxtGenWP2_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

'Private Sub TxtGenWP3_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub


'Private Sub TxtGenWP4_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

'Private Sub TxtGenWP5_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

'Private Sub TxtGenWP6_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

'Private Sub txtPREWP1_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub


'Private Sub txtPREWP2_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub


'Private Sub txtPREWP3_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub


'Private Sub txtPREWP4_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

'Private Sub txtPREWP5_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

'Private Sub txtPREWP6_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

Private Sub txtPRMEffdate_LostFocus()
    If txtPRMEffdate <> "__/__/____" Then
        If IsDate(txtPRMEffdate) = False Then
            MsgBox "Invalid date format!", vbCritical
        End If
    End If
End Sub

'Private Sub TxtSpeWP1_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

'Private Sub TxtSpeWP2_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

'Private Sub TxtSpeWP3_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

'Private Sub TxtSpeWP4_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

'Private Sub TxtSpeWP5_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

'Private Sub TxtSpeWP6_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

Private Sub TxtSupplyAL_KeyPress(KeyAscii As Integer)
Clipboard.Clear
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtAnnualLimit_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtDay_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub cboPLNPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPLNPayor.Text, cboPLNPayor.SelStart) & _
                Chr(KeyAscii) + Mid(cboPLNPayor.Text, cboPLNPayor.SelStart + cboPLNPayor.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPLNPayor.ListCount - 1
            
            If NewText = ChkStr(Left(cboPLNPayor.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPLNPayor.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboPLNPayor.Text = ValidValue
              cboPLNPayor.SelStart = 0
              cboPLNPayor.SelLength = Len(ValidValue)
              End If
End Sub

Private Sub cboPLNPayor2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPLNPayor2.Text, cboPLNPayor2.SelStart) & _
                Chr(KeyAscii) + Mid(cboPLNPayor2.Text, cboPLNPayor2.SelStart + cboPLNPayor2.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPLNPayor2.ListCount - 1
            
            If NewText = ChkStr(Left(cboPLNPayor2.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPLNPayor2.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboPLNPayor2.Text = ValidValue
              cboPLNPayor2.SelStart = 0
              cboPLNPayor2.SelLength = Len(ValidValue)
              End If
End Sub

Private Sub cboPmAgeCode_KeyPress(KeyAscii As Integer)
   ' If secondeditflag = False Then
    
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPmAgeCode.Text, cboPmAgeCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboPmAgeCode.Text, cboPmAgeCode.SelStart + cboPmAgeCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPmAgeCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboPmAgeCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPmAgeCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
              cboPmAgeCode.Text = ValidValue
              cboPmAgeCode.SelStart = 0
              cboPmAgeCode.SelLength = Len(ValidValue)
            End If
   End If
End Sub

Private Sub cboPmINSCode_KeyPress(KeyAscii As Integer)
  ' If secondeditflag = False Then
    Clipboard.Clear
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPmINSCode.Text, cboPmINSCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboPmINSCode.Text, cboPmINSCode.SelStart + cboPmINSCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPmINSCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboPmINSCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPmINSCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
              cboPmINSCode.Text = ValidValue
              cboPmINSCode.SelStart = 0
              cboPmINSCode.SelLength = Len(ValidValue)
            End If
   End If
'End If
End Sub

Private Sub cboPmPLNCode_KeyPress(KeyAscii As Integer)
   ' If secondeditflag = False Then
    
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPmPLNCode.Text, cboPmPLNCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboPmPLNCode.Text, cboPmPLNCode.SelStart + cboPmPLNCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPmPLNCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboPmPLNCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPmPLNCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
              cboPmPLNCode.Text = ValidValue
              cboPmPLNCode.SelStart = 0
              cboPmPLNCode.SelLength = Len(ValidValue)
            End If
   End If
'End If
End Sub

Private Sub cboPmPAYCode_KeyPress(KeyAscii As Integer)
  '  If secondeditflag = False Then
    
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPmPAYCode.Text, cboPmPAYCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboPmPAYCode.Text, cboPmPAYCode.SelStart + cboPmPAYCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPmPAYCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboPmPAYCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPmPAYCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
              cboPmPAYCode.Text = ValidValue
              cboPmPAYCode.SelStart = 0
              cboPmPAYCode.SelLength = Len(ValidValue)
            End If
   End If
End Sub

Private Sub TxtIAAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtICAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtFAAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtFCAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtGAAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtGCAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtHAAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtHCAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub CboMCOPln_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(CboMCOPln.Text, CboMCOPln.SelStart) & _
                Chr(KeyAscii) + Mid(CboMCOPln.Text, CboMCOPln.SelStart + CboMCOPln.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To CboMCOPln.ListCount - 1
            
            If NewText = ChkStr(Left(CboMCOPln.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = CboMCOPln.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
              CboMCOPln.Text = ValidValue
              CboMCOPln.SelStart = 0
              CboMCOPln.SelLength = Len(ValidValue)
            End If
   End If
End Sub

'############################## Start Visit ##############################'
Private Sub SaveVisit()
Dim Visit As New ADODB.Recordset
Dim SaveVisit As New ADODB.Recordset
Dim RecordSaved
Dim StrSql As String
    
    If Len(Trim(CboVPLNCode)) = 0 Then
        MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboVPLNCode.SetFocus
    ElseIf Len(Trim(CboVHltCode)) = 0 Then
        MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboVHltCode.SetFocus
    ElseIf Len(Trim(CboVINSCode)) = 0 Then
        MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboVINSCode.SetFocus
    ElseIf Len(Trim(TxtVisitLimit)) = 0 Then
        MsgBox "Please Enter Visit Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtVisitLimit.SetFocus
    ElseIf TxtVLimitEffDate = "__/__/____" Then
        MsgBox "Please Enter Visit Limit Effective Date", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf TxtVLimitVer = "" Then
        MsgBox "Please Enter Visit Limit Version", vbOKOnly + vbCritical, DialogBoxTitle
        TxtVLimitVer.SetFocus
    ElseIf Mid(CboVSuppStatus, 1, 1) = "B" And TxtVSuppVisit = "" Then
        MsgBox "Please Enter Supp Visit Limit", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Mid(CboVSuppStatus, 1, 1) = "C" And TxtVSuppVisit = "" Then
        MsgBox "Please Enter Supp Visit Limit", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        
        StrSql = "Select * From VisitLimit"
        StrSql = StrSql & " Where " & _
                " INSCode = '" & Trim(Mid(CboVINSCode, 1, 1)) & "' " & _
                " AND HLTCode = '" & Trim(Mid(CboVHltCode, 1, 1)) & "' " & _
                " AND PLNCode = '" & Trim(Mid(CboVPLNCode, 1, 4)) & "' " & _
                " AND VisitLimitVersion = '" & Trim(TxtVLimitVer) & "'" & _
                " AND VisitEffDate = '" & Format(TxtVLimitEffDate, "mm/dd/yyyy") & "'"
        
        Visit.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
      If Visit.RecordCount > 0 And EditFlag = False Then
        MsgBox "Record Duplicated! " & vbNewLine & _
                " Please Change The Corresponding Field", vbCritical, "Error"
        Visit.Close
        CboVINSCode.SetFocus
        CboVINSCode.SelStart = 0
        CboVINSCode.SelLength = Len(CboVINSCode)
        Set Visit = Nothing
        Exit Sub
      Else
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                    SaveVisit.Open "VisitLimit", MediConnCISDB, adOpenKeyset, adLockOptimistic
                    SaveVisit.AddNew
                Else
                    'strsql = " Select * From AnnualLimit Where PLNCode = '" & Trim(CboPlnCode) & "'" & _
                             " AND HLTCode = '" & Trim(CboHltCode) & "' AND INSCode = '" & Trim(CboInsType) & "'" & _
                             " AND AnnLimitVersion = '" & Trim(txtANNVersion) & "' AND AnnEffDate = '" & Trim(txtALEffDate) & "'"
                    
                    StrSql = " Select * From VisitLimit Where VisitIndex = '" & ChkStr(lstVisitLimit.SelectedItem.SubItems(1)) & "'"
                    'If Trim(temptxtAnnEffDate) <> "" Then
                        'strsql = strsql & " and AnnualEffDate = '" & Format(temptxtAnnEffDate, "mm/dd/yyyy") & "' "
                    'Else
                        'strsql = strsql & " and (AnnualEffDate) IS  NULL "
                    'End If
                    SaveVisit.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                    'If SaveAnnual.RecordCount = 0 Then
 '                       SaveAnnual.AddNew
                    'End If
                End If
            
                With SaveVisit
                    If InStr(1, CboVPLNCode, "-") > 0 Then
                        !PLNCode = Trim(Mid(CboVPLNCode, 1, InStr(1, CboVPLNCode, "-") - 1))
                    Else
                        !PLNCode = Trim(Mid(CboVPLNCode, 1, 4))
                    End If
                    
                    !INSCode = Mid(CboVINSCode, 1, 1)
                    !HLTCode = Mid(CboVHltCode, 1, 1)
                    !VisitLimit = TxtVisitLimit
                    
                    If Mid(CboVSuppStatus, 1, 1) = "A" Then
                        !SuppVisitLimitStatus = "A"
                    ElseIf Mid(CboVSuppStatus, 1, 1) = "B" Then
                        !SuppVisitLimitStatus = "B"
                        !SuppVisitLimit = TxtVSuppVisit
                    ElseIf Mid(CboVSuppStatus, 1, 1) = "C" Then
                        !SuppVisitLimitStatus = "C"
                        !SuppVisitLimit = TxtVSuppVisit
                    End If
                                        
                    !VisitEffDate = IIf(TxtVLimitEffDate = "__/__/____", "", CDate(TxtVLimitEffDate))
                    !VisitLimitVersion = TxtVLimitVer
                    !VisitLastUpdateUser = Logon
                    !VisitLastUpdateDate = Date
                End With
                SaveVisit.update
                Call ClearForm(Me)
                Call VisitListing(VisitCriteria)
                    If EditFlag = False Then
                        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    Else
                        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
 
                SaveVisit.Close
                Set SaveVisit = Nothing
            End If
        EditFlag = False
   End If
  End If
  'End If
  Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub SearchVisit()
    Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(CboVPLNCode)) Then
        strAppend = strAppend + " And PLNCode = '" & Trim(Mid(CboVPLNCode, 1, IIf(InStr(1, CboVPLNCode, "-") = 0, 4, InStr(1, CboVPLNCode, "-") - 1))) & "'"
    End If
    If Len(Trim(CboVHltCode)) Then
        strAppend = strAppend + " And HLTCode = '" & Trim(Mid(CboVHltCode, 1, IIf(InStr(1, CboVHltCode, "-") = 0, 4, InStr(1, CboVHltCode, "-") - 1))) & "'"
    End If
    
    If Len(Trim(CboVINSCode)) Then
        strAppend = strAppend + " And INSCode = '" & Trim(Mid(CboVINSCode, 1, IIf(InStr(1, CboVINSCode, "-") = 0, 4, InStr(1, CboVINSCode, "-") - 1))) & "'"
    End If
    
    If Len(Trim(CboVPAYCode)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboVPAYCode, 1, IIf(InStr(1, CboVPAYCode, "-") = 0, 4, InStr(1, CboVPAYCode, "-") - 1))) & "'"
    End If

    If Len(Trim(CboVGrpCom)) Then
        strAppend = strAppend + " AND GRPCompany like '" & RemStr(Mid(CboVGrpCom, 1, IIf(InStr(1, CboVGrpCom, "-") = 0, 4, InStr(1, CboVGrpCom, "-") - 1))) & "%'"
    End If
    
    VisitCriteria = strAppend
    Call VisitListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub DeleteVisit()
Dim DeleteVisit As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim StrSql, sqlMBM As String
Dim DeleteRecord

EditFlag = False
If CboVPLNCode = "" Then
   MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
ElseIf lstVisitLimit.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBM where PLNCode = '" & RemStr(CboVPLNCode) & "'" 'ChkStr(lstvisitlimit.SelectedItem.Text) & "'"
   
   rsCheckMBM.MaxRecords = 1
   rsCheckMBM.Open sqlMBM, MediConnCISDB, adOpenKeyset, adLockOptimistic
   
     
   If rsCheckMBM.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.RecordCount = 0 Then
           StrSql = " Delete From VisitLimit Where VisitIndex = '" & ChkStr(lstVisitLimit.SelectedItem.SubItems(1)) & "'"
                    
           DeleteVisit.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call VisitListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           CboVPLNCode.SetFocus
         End If
     End If
End If
Exit Sub
End Sub

Private Sub VisitListing(Optional strAppend As String)
'On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim VisitMaster As ListItem
    Dim sql As String
    Dim TotalVisitCount As Integer
    
    
    TotalVisitCount = 0
    lstVisitLimit.ListItems.Clear
    
    sql = " Select PLNCode, INSCode, HLTCode, PAYCode, GRPCompany, VisitIndex, " & _
          " VisitLimit, SuppVisitLimit, VisitEffDate, SuppVisitLimitStatus, VisitLimitVersion " & _
          " From VIEW_VISIT_PAY_PLN where 0=0 "
    
    If Len(Trim(strAppend)) Then
    sql = sql + strAppend
    End If

    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set VisitMaster = lstVisitLimit.ListItems.Add(, , rst2!HLTCode)
                VisitMaster.SubItems(1) = rst2!VisitIndex
                VisitMaster.SubItems(2) = rst2!PLNCode
                VisitMaster.SubItems(3) = rst2!INSCode
                VisitMaster.SubItems(4) = Format(rst2!VisitLimit, "#,##0.00")
                VisitMaster.SubItems(5) = Format(rst2!SuppVisitLimit, "#,##0.00")
                VisitMaster.SubItems(6) = IIf(Len(Trim(rst2!VisitEffDate)) >= 1, rst2!VisitEffDate, "")
                If Len(rst2!PAYCode) Then
                    VisitMaster.SubItems(7) = rst2!PAYCode
                End If
                If Len(rst2!GRPCompany) Then
                   VisitMaster.SubItems(8) = rst2!GRPCompany
                End If
                If Len(rst2!VisitLimitVersion) Then
                   VisitMaster.SubItems(9) = rst2!VisitLimitVersion
                End If
                If Len(rst2!SuppVisitLimitStatus) Then
                   VisitMaster.SubItems(10) = rst2!SuppVisitLimitStatus
                End If
                
                rst2.MoveNext
        TotalVisitCount = TotalVisitCount + 1
        txtTotalRecord = TotalVisitCount
        Loop
        
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub TxtVSuppVisit_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub lstVisitLimit_Click()

    If lstVisitLimit.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboVHltCode = lstVisitLimit.SelectedItem.Text
        CboVPLNCode = lstVisitLimit.SelectedItem.SubItems(2)
        CboVINSCode = lstVisitLimit.SelectedItem.SubItems(3)
        TxtVisitLimit = lstVisitLimit.SelectedItem.SubItems(4)
        TxtVSuppVisit = lstVisitLimit.SelectedItem.SubItems(5)
        TxtVLimitEffDate = lstVisitLimit.SelectedItem.SubItems(6)
        CboVPAYCode = lstVisitLimit.SelectedItem.SubItems(7)
        CboVGrpCom = lstVisitLimit.SelectedItem.SubItems(8)
        TxtVLimitVer = lstVisitLimit.SelectedItem.SubItems(9)
        
        If lstVisitLimit.SelectedItem.SubItems(10) = "A" Then
            CboVSuppStatus.ListIndex = 0
        ElseIf lstVisitLimit.SelectedItem.SubItems(10) = "B" Then
            CboVSuppStatus.ListIndex = 1
        ElseIf lstVisitLimit.SelectedItem.SubItems(10) = "C" Then
            CboVSuppStatus.ListIndex = 2
        End If
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstVisitLimit_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstVisitLimit, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstVisitLimit, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstVisitLimit, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstVisitLimit, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstVisitLimit, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstVisitLimit, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstVisitLimit, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstVisitLimit, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstVisitLimit, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstVisitLimit, ColumnHeader.Index, ldtString, m_blnDirection(10)
    End Select
End Sub

'########################### Start Co-Payment MBM #########################'

Private Sub SaveCoMBM()
Dim COM As New ADODB.Recordset
Dim SaveCOM As New ADODB.Recordset
Dim RecordSaved
Dim StrSql As String
    
    If Len(Trim(cboCOPHLTType)) = "" Then
        MsgBox "Please Enter Health Type", vbOKOnly + vbCritical, DialogBoxTitle
        cboCOPHLTType.SetFocus
    ElseIf Len(Trim(cboCOPPLNCode)) = "" Then
        MsgBox "Please Enter Plan Code ", vbOKOnly + vbCritical, DialogBoxTitle
        cboCOPPLNCode.SetFocus
    ElseIf txtCOPAmt = "" Then
        MsgBox "Please Enter Amount ", vbOKOnly + vbCritical, DialogBoxTitle
        txtCOPAmt.SetFocus
    Else
        StrSql = "Select * From COPaymentMBM where PLNCode = '" & Trim(cboCOPPLNCode) & "'"
        COM.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If COM.RecordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            COM.Close
            cboCOPPLNCode.SetFocus
            cboCOPPLNCode.SelStart = 0
            cboCOPPLNCode.SelLength = Len(cboCOPPLNCode)
            Set COM = Nothing
            Set SaveCOM = Nothing
            Exit Sub
            
        Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                   SaveCOM.Open "COPaymentMBM", MediConnCISDB, adOpenKeyset, adLockOptimistic
                   SaveCOM.AddNew
                Else
                   StrSql = "Select * From COPaymentMBM Where PLNCode = '" & ChkStr(cboCOPPLNCode) & "'"
                   SaveCOM.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                End If
                With SaveCOM
                     !HltType = Trim(Mid(cboCOPHLTType, 1, IIf(InStr(1, cboCOPHLTType, "-") = 0, 1, InStr(1, cboCOPHLTType, "-") - 1)))
                     !PLNCode = Trim(Mid(cboCOPPLNCode, 1, IIf(InStr(1, cboCOPPLNCode, "-") = 0, 4, InStr(1, cboCOPPLNCode, "-") - 1)))
                     !CoAmount = txtCOPAmt
                     !CoLastUpdateUser = Logon
                     !CoLastUpdateDate = Date
                     
                     
                End With
                SaveCOM.update
                Call ClearForm(Me)
                Call CoMBMListing(CoCriteria)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
        
                SaveCOM.Close
            End If
            EditFlag = False
      End If
  End If
End Sub

Private Sub SearchCoMBM()
Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(cboCOPHLTType)) Then
        strAppend = strAppend + " And HLTType = '" & Trim(Mid(cboCOPHLTType, 1, IIf(InStr(1, cboCOPHLTType, "-") = 0, 4, InStr(1, cboCOPHLTType, "-") - 1))) & "'"
    End If
    If Len(Trim(cboCOPPLNCode)) Then
        strAppend = strAppend + " And PLNCode = '" & Trim(Mid(cboCOPPLNCode, 1, IIf(InStr(1, cboCOPPLNCode, "-") = 0, 4, InStr(1, cboCOPPLNCode, "-") - 1))) & "'"
    End If
       
    CoCriteria = strAppend
    Call CoMBMListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub DeleteCoMBM()
Dim DeleteCOM As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim StrSql, sqlMBM As String
Dim DeleteRecord

EditFlag = False
If cboCOPPLNCode = "" Then
   MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
ElseIf lstCoPaymentMBMList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
'Else
'   sqlMBM = "Select MBMNumber from MBM where PLNCode = '" & RemStr(cboCOPPLNCode) & "'" 'ChkStr(lstvisitlimit.SelectedItem.Text) & "'"
   
'   rsCheckMBM.MaxRecords = 1
'   rsCheckMBM.Open sqlMBM, MediConnCISDB, adOpenKeyset, adLockOptimistic
   
     
'   If rsCheckMBM.RecordCount Then
'        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes Then
     'And rsCheckMBM.RecordCount = 0 Then
           StrSql = " Delete From COPaymentMBM Where PLNCode = '" & ChkStr(lstCoPaymentMBMList.SelectedItem.SubItems(1)) & "'"
                    
           DeleteCOM.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call CoMBMListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           cboCOPPLNCode.SetFocus
         End If
     End If
Exit Sub
End Sub

Private Sub CoMBMListing(Optional strAppend As String)
'On Error Resume Next
Dim rst2 As New ADODB.Recordset
Dim CoMaster As ListItem
Dim sql As String
Dim TotalComCount As Integer
    
    
    TotalComCount = 0
    lstCoPaymentMBMList.ListItems.Clear
    
    sql = " Select * from COPaymentMBM where 0 = 0"
          '" VisitLimit, SuppVisitLimit, VisitEffDate, SuppVisitLimitStatus, VisitLimitVersion " & _
          '" From VIEW_VISIT_PAY_PLN where 0=0 "
    
    If Len(Trim(strAppend)) Then
    sql = sql + strAppend
    End If

    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set CoMaster = lstCoPaymentMBMList.ListItems.Add(, , rst2!HltType)
                CoMaster.SubItems(1) = rst2!PLNCode
                CoMaster.SubItems(2) = rst2!CoAmount
                
                rst2.MoveNext
        TotalComCount = TotalComCount + 1
        txtTotalRecord = TotalComCount
        Loop
        
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub lstCoPaymentMBMList_Click()

    If lstCoPaymentMBMList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        cboCOPHLTType = lstCoPaymentMBMList.SelectedItem.Text
        cboCOPPLNCode = lstCoPaymentMBMList.SelectedItem.SubItems(1)
        txtCOPAmt = lstCoPaymentMBMList.SelectedItem.SubItems(2)
       
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstCoPaymentMBMList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstVisitLimit, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstVisitLimit, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstVisitLimit, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    End Select
End Sub

Private Sub CoMBMComboListing()
Dim HLT As New ADODB.Recordset
Dim PLN As New ADODB.Recordset
    
    HLT.Open "Select  HLTCode  , HLTDesc from HLT", MediConnCISDB, adOpenKeyset, adLockOptimistic
    PLN.Open "select PLNCode, PLNDesc  from PLN", MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    cboCOPHLTType.Clear
    cboCOPPLNCode.Clear
       
    Do Until HLT.EOF
       cboCOPHLTType.AddItem HLT!HLTCode & " - " & HLT!HLTDesc
       HLT.MoveNext
    Loop
    
    Do Until PLN.EOF
       cboCOPPLNCode.AddItem PLN!PLNCode & " - " & PLN!PlnDesc
       PLN.MoveNext
    Loop
    
    HLT.Close
    Set HLT = Nothing
    PLN.Close
    Set PLN = Nothing
    
End Sub

Private Sub CboPDentalStatus_Change()
    If InStr(CboPDentalStatus.Text, "null") Then
       CboPDentalStatus.Enabled = False
    End If
End Sub

Private Sub CboPDentalStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPDentalStatus.Text, CboPDentalStatus.SelStart) + Chr(KeyAscii) + Mid(CboPDentalStatus.Text, CboPDentalStatus.SelStart + CboPDentalStatus.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPDentalStatus.ListCount - 1
 
        If NewText = LCase(Left(CboPDentalStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPDentalStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboPDentalStatus.Text = ValidValue
               CboPDentalStatus.SelStart = 0
               CboPDentalStatus.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboPDentalLimit_Change()
    If InStr(CboPDentalLimit.Text, "null") Then
       CboPDentalLimit.Enabled = False
    End If
End Sub

Private Sub CboPDentalLimit_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPDentalLimit.Text, CboPDentalLimit.SelStart) + Chr(KeyAscii) + Mid(CboPDentalLimit.Text, CboPDentalLimit.SelStart + CboPDentalLimit.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPDentalLimit.ListCount - 1
 
        If NewText = LCase(Left(CboPDentalLimit.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPDentalLimit.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboPDentalLimit.Text = ValidValue
               CboPDentalLimit.SelStart = 0
               CboPDentalLimit.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboPSpectStatus_Change()
    If InStr(CboPSpectStatus.Text, "null") Then
       CboPSpectStatus.Enabled = False
    End If
End Sub

Private Sub CboPSpectStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPSpectStatus.Text, CboPSpectStatus.SelStart) + Chr(KeyAscii) + Mid(CboPSpectStatus.Text, CboPSpectStatus.SelStart + CboPSpectStatus.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPSpectStatus.ListCount - 1
 
        If NewText = LCase(Left(CboPSpectStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPSpectStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboPSpectStatus.Text = ValidValue
               CboPSpectStatus.SelStart = 0
               CboPSpectStatus.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboPSpectLimit_Change()
    If InStr(CboPSpectLimit.Text, "null") Then
       CboPSpectLimit.Enabled = False
    End If
End Sub

Private Sub CboPSpectLimit_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPSpectLimit.Text, CboPSpectLimit.SelStart) + Chr(KeyAscii) + Mid(CboPSpectLimit.Text, CboPSpectLimit.SelStart + CboPSpectLimit.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPSpectLimit.ListCount - 1
 
        If NewText = LCase(Left(CboPSpectLimit.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPSpectLimit.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboPSpectLimit.Text = ValidValue
               CboPSpectLimit.SelStart = 0
               CboPSpectLimit.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboPMEPStatus_Change()
    If InStr(CboPMEPStatus.Text, "null") Then
       CboPMEPStatus.Enabled = False
    End If
End Sub

Private Sub CboPMEPStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPMEPStatus.Text, CboPMEPStatus.SelStart) + Chr(KeyAscii) + Mid(CboPMEPStatus.Text, CboPMEPStatus.SelStart + CboPMEPStatus.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPMEPStatus.ListCount - 1
 
        If NewText = LCase(Left(CboPMEPStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPMEPStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboPMEPStatus.Text = ValidValue
               CboPMEPStatus.SelStart = 0
               CboPMEPStatus.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboSDentalStatus_Change()
    If InStr(CboSDentalStatus.Text, "null") Then
       CboSDentalStatus.Enabled = False
    End If
End Sub

Private Sub CboSDentalStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboSDentalStatus.Text, CboSDentalStatus.SelStart) + Chr(KeyAscii) + Mid(CboSDentalStatus.Text, CboSDentalStatus.SelStart + CboSDentalStatus.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboSDentalStatus.ListCount - 1
 
        If NewText = LCase(Left(CboSDentalStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboSDentalStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboSDentalStatus.Text = ValidValue
               CboSDentalStatus.SelStart = 0
               CboSDentalStatus.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboSDentalLimit_Change()
    If InStr(CboSDentalLimit.Text, "null") Then
       CboSDentalLimit.Enabled = False
    End If
End Sub

Private Sub CboSDentalLimit_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboSDentalLimit.Text, CboSDentalLimit.SelStart) + Chr(KeyAscii) + Mid(CboSDentalLimit.Text, CboSDentalLimit.SelStart + CboSDentalLimit.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboSDentalLimit.ListCount - 1
 
        If NewText = LCase(Left(CboSDentalLimit.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboSDentalLimit.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboSDentalLimit.Text = ValidValue
               CboSDentalLimit.SelStart = 0
               CboSDentalLimit.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboSSpecStatus_Change()
    If InStr(CboSSpecStatus.Text, "null") Then
       CboSSpecStatus.Enabled = False
    End If
End Sub

Private Sub CboSSpecStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboSSpecStatus.Text, CboSSpecStatus.SelStart) + Chr(KeyAscii) + Mid(CboSSpecStatus.Text, CboSSpecStatus.SelStart + CboSSpecStatus.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboSSpecStatus.ListCount - 1
 
        If NewText = LCase(Left(CboSSpecStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboSSpecStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboSSpecStatus.Text = ValidValue
               CboSSpecStatus.SelStart = 0
               CboSSpecStatus.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboSSpecLimit_Change()
    If InStr(CboSSpecLimit.Text, "null") Then
       CboSSpecLimit.Enabled = False
    End If
End Sub

Private Sub CboSSpecLimit_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboSSpecLimit.Text, CboSSpecLimit.SelStart) + Chr(KeyAscii) + Mid(CboSSpecLimit.Text, CboSSpecLimit.SelStart + CboSSpecLimit.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboSSpecLimit.ListCount - 1
 
        If NewText = LCase(Left(CboSSpecLimit.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboSSpecLimit.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboSSpecLimit.Text = ValidValue
               CboSSpecLimit.SelStart = 0
               CboSSpecLimit.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboSMEPStatus_Change()
    If InStr(CboSMEPStatus.Text, "null") Then
       CboSMEPStatus.Enabled = False
    End If
End Sub

Private Sub CboSMEPStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboSMEPStatus.Text, CboSMEPStatus.SelStart) + Chr(KeyAscii) + Mid(CboSMEPStatus.Text, CboSMEPStatus.SelStart + CboSMEPStatus.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboSMEPStatus.ListCount - 1
 
        If NewText = LCase(Left(CboSMEPStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboSMEPStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboSMEPStatus.Text = ValidValue
               CboSMEPStatus.SelStart = 0
               CboSMEPStatus.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboSAnnLimit_Change()
    If InStr(CboSAnnLimit.Text, "null") Then
       CboSAnnLimit.Enabled = False
    End If
End Sub

Private Sub CboSAnnLimit_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboSAnnLimit.Text, CboSAnnLimit.SelStart) + Chr(KeyAscii) + Mid(CboSAnnLimit.Text, CboSAnnLimit.SelStart + CboSAnnLimit.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboSAnnLimit.ListCount - 1
 
        If NewText = LCase(Left(CboSAnnLimit.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboSAnnLimit.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboSAnnLimit.Text = ValidValue
               CboSAnnLimit.SelStart = 0
               CboSAnnLimit.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboSVisitLimit_Change()
    If InStr(CboSVisitLimit.Text, "null") Then
       CboSVisitLimit.Enabled = False
    End If
End Sub

Private Sub CboSVisitLimit_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboSVisitLimit.Text, CboSVisitLimit.SelStart) + Chr(KeyAscii) + Mid(CboSVisitLimit.Text, CboSVisitLimit.SelStart + CboSVisitLimit.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboSVisitLimit.ListCount - 1
 
        If NewText = LCase(Left(CboSVisitLimit.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboSVisitLimit.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboSVisitLimit.Text = ValidValue
               CboSVisitLimit.SelStart = 0
               CboSVisitLimit.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboNoofPlan_Change()
    If InStr(CboNoofPlan.Text, "null") Then
       CboNoofPlan.Enabled = False
    End If
End Sub

Private Sub CboNoofPlan_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboNoofPlan.Text, CboNoofPlan.SelStart) + Chr(KeyAscii) + Mid(CboNoofPlan.Text, CboNoofPlan.SelStart + CboNoofPlan.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboNoofPlan.ListCount - 1
 
        If NewText = LCase(Left(CboNoofPlan.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboNoofPlan.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboNoofPlan.Text = ValidValue
               CboNoofPlan.SelStart = 0
               CboNoofPlan.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub SaveMCO1(tmpInsType As String, tmpAdultAmt As Double, tmpChildAmt As Double, tmpSuppAdultAmt As Double, tmpSuppChildAmt As Double)
Dim StrMCOSql As String
Dim StrSql1, StrSql2 As String
Dim MCORec As New ADODB.Recordset
Dim RecFound As Boolean
    
    StrSql1 = " Select MCOIndex From MCO where PLNCode = '" & Trim(CboMCOPln) & "'" & _
              " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
              " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'"
    StrSql2 = ""
    StrSql2 = StrSql1 & " AND INSCode = '" & tmpInsType & "'"
    MCORec.Open StrSql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not MCORec.EOF
        RecFound = True
        MsgBox "Record Not Save for Insured Type '" & tmpInsType & "! Please Change The Corresponding Field", vbCritical, "Error"
        Exit Do
    Loop
    MCORec.Close
    Set MCORec = Nothing
    
    If RecFound = False Then
        StrMCOSql = "Insert Into MCO(PlnCode, HltCode, InsCode, " & _
                "SuppMCOStatus, MCOEffDate," & _
                " MCOAmtAdult, MCOAmtChild, MCOSuppAdultAmt, MCOSuppChildAmt, " & _
                " MCOLastUpdateUser, MCOLastUpdateDate) " & _
                "Values('" & Trim(CboMCOPln) & "','" & Trim(Mid(CboMCOHlt, 1, 1)) & "','" & tmpInsType & "'," & _
                "'" & Mid(cboSuppMCOStatus, 1, 1) & "','" & Format(txtEffDate, "yyyy/mm/dd") & "'," & _
                "" & tmpAdultAmt & "," & tmpChildAmt & "," & tmpSuppAdultAmt & "," & tmpSuppChildAmt & "," & _
                "'" & Logon & "','" & Format(Date, "yyyy/mm/dd") & "')"
        
        MediConnCISDB.Execute StrMCOSql
    End If
End Sub
