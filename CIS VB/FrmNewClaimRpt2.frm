VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmNewClaimRpt2 
   Caption         =   "Claims Report 2 - Jupiter"
   ClientHeight    =   8592
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8592
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Caption         =   "Reports Selection"
      Height          =   4095
      Left            =   7920
      TabIndex        =   135
      Top             =   2880
      Width           =   3855
      Begin VB.OptionButton Option15 
         Caption         =   "By Group Company"
         Height          =   195
         Left            =   120
         TabIndex        =   122
         Top             =   3000
         Width           =   1935
      End
      Begin VB.OptionButton Option6 
         Caption         =   "By Ins. Type && Ageband"
         Height          =   195
         Left            =   120
         TabIndex        =   113
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton Option5 
         Caption         =   "By Ins. Type && Plan"
         Height          =   195
         Left            =   120
         TabIndex        =   112
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton Option4 
         Caption         =   "By Claim No"
         Height          =   195
         Left            =   120
         TabIndex        =   111
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option7 
         Caption         =   "By Ins. Type && Sex"
         Height          =   195
         Left            =   120
         TabIndex        =   114
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton Option8 
         Caption         =   "By Ageband"
         Height          =   195
         Left            =   120
         TabIndex        =   115
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton Option9 
         Caption         =   "By Sex"
         Height          =   195
         Left            =   120
         TabIndex        =   116
         Top             =   1560
         Width           =   975
      End
      Begin VB.OptionButton Option10 
         Caption         =   "By Ageband && Sex"
         Height          =   195
         Left            =   120
         TabIndex        =   117
         Top             =   1800
         Width           =   1695
      End
      Begin VB.OptionButton Option11 
         Caption         =   "By Diagnosis"
         Height          =   195
         Left            =   120
         TabIndex        =   118
         Top             =   2040
         Width           =   1455
      End
      Begin VB.OptionButton Option12 
         Caption         =   "By Clinic"
         Height          =   195
         Left            =   120
         TabIndex        =   119
         Top             =   2280
         Width           =   1215
      End
      Begin VB.OptionButton Option13 
         Caption         =   "By Payor"
         Height          =   195
         Left            =   120
         TabIndex        =   120
         Top             =   2520
         Width           =   975
      End
      Begin VB.OptionButton Option14 
         Caption         =   "By Payor, Ins. Type && Plan"
         Height          =   195
         Left            =   120
         TabIndex        =   121
         Top             =   2760
         Width           =   2295
      End
      Begin VB.OptionButton Option16 
         Caption         =   "By Ins. Type && Rel."
         Height          =   195
         Left            =   2160
         TabIndex        =   123
         Top             =   360
         Width           =   1650
      End
      Begin VB.OptionButton Option17 
         Caption         =   "By Ins. Type"
         Height          =   195
         Left            =   2160
         TabIndex        =   124
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option18 
         Caption         =   "By Bordx No"
         Height          =   195
         Left            =   2160
         TabIndex        =   125
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option19 
         Caption         =   "By Policy Year"
         Height          =   195
         Left            =   2160
         TabIndex        =   126
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton Option20 
         Caption         =   "By Takeover"
         Height          =   195
         Left            =   2160
         TabIndex        =   127
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Frame Frame22 
      Caption         =   "Selection Criteria"
      Height          =   6615
      Left            =   0
      TabIndex        =   142
      Top             =   360
      Width           =   7905
      Begin VB.TextBox txtPreDiag1 
         Height          =   285
         Left            =   4080
         MaxLength       =   3
         TabIndex        =   64
         Top             =   2400
         Width           =   1200
      End
      Begin VB.CheckBox ChkCLNGrp1 
         Height          =   255
         Left            =   7560
         TabIndex        =   52
         Top             =   1040
         Width           =   255
      End
      Begin VB.CheckBox ChkCLN1 
         Height          =   255
         Left            =   7560
         TabIndex        =   49
         Top             =   750
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom1 
         Height          =   255
         Left            =   7560
         TabIndex        =   46
         Top             =   500
         Width           =   255
      End
      Begin VB.CheckBox ChkPayor1 
         Height          =   255
         Left            =   7560
         TabIndex        =   43
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox ChkICNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7320
         TabIndex        =   94
         Top             =   4920
         Width           =   375
      End
      Begin VB.CheckBox ChkMBMIC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7320
         TabIndex        =   97
         Top             =   5160
         Width           =   375
      End
      Begin VB.CheckBox ChkSubmission 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7320
         TabIndex        =   107
         Top             =   6240
         Width           =   375
      End
      Begin VB.CheckBox ChkInvoice 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7320
         TabIndex        =   104
         Top             =   6000
         Width           =   375
      End
      Begin VB.CheckBox ChkBranch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7320
         TabIndex        =   54
         Top             =   1300
         Width           =   375
      End
      Begin VB.CheckBox ChkMCDays 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2640
         TabIndex        =   40
         Top             =   5685
         Width           =   255
      End
      Begin VB.CheckBox ChkBordxDate 
         Caption         =   "Check1"
         Height          =   255
         Left            =   7320
         TabIndex        =   91
         Top             =   4680
         Width           =   255
      End
      Begin VB.ComboBox cboPayorCode 
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
         Left            =   4080
         Sorted          =   -1  'True
         TabIndex        =   41
         Top             =   200
         Width           =   3120
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   200
         Width           =   1320
      End
      Begin VB.CheckBox ChkDOV 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7320
         TabIndex        =   57
         Top             =   1575
         Width           =   255
      End
      Begin VB.CheckBox chkExpDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   7320
         TabIndex        =   75
         Top             =   3120
         Width           =   255
      End
      Begin VB.CheckBox chkDOR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7320
         TabIndex        =   60
         Top             =   1860
         Width           =   255
      End
      Begin VB.CheckBox ChkJoinDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   7320
         TabIndex        =   78
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox ChkDOB 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   7320
         TabIndex        =   81
         Top             =   3600
         Width           =   255
      End
      Begin VB.CheckBox ChkPayor 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7320
         TabIndex        =   42
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7320
         TabIndex        =   45
         Top             =   500
         Width           =   255
      End
      Begin VB.CheckBox ChkPolNo 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   33
         Top             =   4920
         Width           =   255
      End
      Begin VB.CheckBox ChkCLN 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7320
         TabIndex        =   48
         Top             =   750
         Width           =   255
      End
      Begin VB.CheckBox ChkCLNGrp 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7320
         TabIndex        =   51
         Top             =   1040
         Width           =   255
      End
      Begin VB.CheckBox ChkFirst 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7320
         TabIndex        =   66
         Top             =   2400
         Width           =   255
      End
      Begin VB.CheckBox ChkEffDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   7320
         TabIndex        =   72
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox ChkBordx 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7320
         TabIndex        =   88
         Top             =   4440
         Width           =   255
      End
      Begin VB.CheckBox ChkPlan 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7320
         TabIndex        =   69
         Top             =   2640
         Width           =   375
      End
      Begin VB.CheckBox ChkClaimability 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   15
         Top             =   2400
         Width           =   375
      End
      Begin VB.CheckBox ChkRenewal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   27
         Top             =   4110
         Width           =   255
      End
      Begin VB.CheckBox ChkDataStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   13
         Top             =   2145
         Width           =   375
      End
      Begin VB.CheckBox ChkRelation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   11
         Top             =   1845
         Width           =   375
      End
      Begin VB.CheckBox ChkTakeOver 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   29
         Top             =   4395
         Width           =   375
      End
      Begin VB.CheckBox ChkAgentCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   31
         Top             =   4665
         Width           =   255
      End
      Begin VB.CheckBox ChkHLT 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox ChkINS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   500
         Width           =   255
      End
      Begin VB.CheckBox ChkRace 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   1305
         Width           =   255
      End
      Begin VB.CheckBox ChkSex 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox ChkState 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   25
         Top             =   3840
         Width           =   255
      End
      Begin VB.CheckBox ChkPreMode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   23
         Top             =   3555
         Width           =   375
      End
      Begin VB.CheckBox ChkExpired 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   21
         Top             =   3285
         Width           =   375
      End
      Begin VB.CheckBox ChkPolStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   2685
         Width           =   255
      End
      Begin VB.CheckBox ChkCaseStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   2970
         Width           =   255
      End
      Begin VB.CheckBox ChkPolYear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2640
         TabIndex        =   37
         Top             =   5400
         Width           =   255
      End
      Begin VB.CheckBox ChkDOW 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7320
         TabIndex        =   63
         Top             =   2175
         Width           =   255
      End
      Begin VB.ComboBox CboInsType 
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1320
      End
      Begin VB.ComboBox CboMemType 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1320
      End
      Begin VB.ComboBox cboSuppType 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1000
         Width           =   1320
      End
      Begin VB.ComboBox CboRace 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1260
         Width           =   1320
      End
      Begin VB.ComboBox CboSex 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1545
         Width           =   1320
      End
      Begin VB.ComboBox CboRelation 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1830
         Width           =   1320
      End
      Begin VB.ComboBox cboDataStatus 
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   2115
         Width           =   1320
      End
      Begin VB.ComboBox CboClaimability 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   2400
         Width           =   1320
      End
      Begin VB.ComboBox CboStatus 
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   2655
         Width           =   1320
      End
      Begin VB.ComboBox CboCaseStatus 
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   2940
         Width           =   1320
      End
      Begin VB.ComboBox CboExpired 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   3230
         Width           =   1320
      End
      Begin VB.ComboBox CboPreMode 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   3510
         Width           =   1320
      End
      Begin VB.ComboBox CboState 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   3800
         Width           =   1320
      End
      Begin VB.ComboBox cboRenewalInd 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   4080
         Width           =   1320
      End
      Begin VB.ComboBox cboTakeOver 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   4365
         Width           =   1320
      End
      Begin VB.TextBox txtAgentCode 
         Height          =   285
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   30
         Top             =   4650
         Width           =   1320
      End
      Begin VB.TextBox txtPolNo 
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   32
         Top             =   4905
         Width           =   1320
      End
      Begin VB.TextBox TxtPatientName 
         Height          =   285
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   34
         Top             =   5145
         Width           =   1320
      End
      Begin VB.TextBox txtPOLYear1 
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   35
         Top             =   5400
         Width           =   480
      End
      Begin VB.TextBox txtPOLYear2 
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   36
         Top             =   5400
         Width           =   480
      End
      Begin VB.TextBox TxtMCDays1 
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   38
         Top             =   5640
         Width           =   480
      End
      Begin VB.TextBox TxtMCDays2 
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   39
         Top             =   5640
         Width           =   480
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   4080
         Sorted          =   -1  'True
         TabIndex        =   44
         Top             =   480
         Width           =   3120
      End
      Begin VB.ComboBox cboCLNName 
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
         Left            =   4080
         Sorted          =   -1  'True
         TabIndex        =   47
         Top             =   720
         Width           =   3120
      End
      Begin VB.ComboBox cboCLNGrp 
         Height          =   315
         Left            =   4080
         Sorted          =   -1  'True
         TabIndex        =   50
         Top             =   1000
         Width           =   3120
      End
      Begin VB.TextBox TxtBranch 
         Height          =   315
         Left            =   4080
         MaxLength       =   50
         TabIndex        =   53
         Top             =   1290
         Width           =   3120
      End
      Begin VB.TextBox txtPreDiag2 
         Height          =   285
         Left            =   6000
         MaxLength       =   3
         TabIndex        =   65
         Top             =   2400
         Width           =   1200
      End
      Begin VB.TextBox TxtPlan1 
         Height          =   315
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   67
         Top             =   2610
         Width           =   1200
      End
      Begin VB.TextBox TxtPlan2 
         Height          =   315
         Left            =   6000
         MaxLength       =   4
         TabIndex        =   68
         Top             =   2610
         Width           =   1200
      End
      Begin VB.TextBox TxtMBMNo1 
         Height          =   285
         Left            =   4080
         MaxLength       =   15
         TabIndex        =   82
         Top             =   3935
         Width           =   1200
      End
      Begin VB.TextBox TxtCaseNo1 
         Height          =   285
         Left            =   4080
         MaxLength       =   15
         TabIndex        =   84
         Top             =   4170
         Width           =   1200
      End
      Begin VB.TextBox TxtBordxNo1 
         Height          =   285
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   86
         Top             =   4425
         Width           =   1200
      End
      Begin VB.TextBox TxtMBMNo2 
         Height          =   285
         Left            =   6000
         MaxLength       =   15
         TabIndex        =   83
         Top             =   3935
         Width           =   1200
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   285
         Left            =   6000
         MaxLength       =   15
         TabIndex        =   85
         Top             =   4170
         Width           =   1200
      End
      Begin VB.TextBox TxtBordxNo2 
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   87
         Top             =   4425
         Width           =   1200
      End
      Begin VB.TextBox TxtICNo1 
         Height          =   285
         Left            =   4080
         MaxLength       =   20
         TabIndex        =   92
         Top             =   4940
         Width           =   1200
      End
      Begin VB.TextBox TxtMBMIc1 
         Height          =   285
         Left            =   4080
         MaxLength       =   20
         TabIndex        =   95
         Top             =   5190
         Width           =   1200
      End
      Begin VB.TextBox TxtICNo2 
         Height          =   285
         Left            =   6000
         MaxLength       =   20
         TabIndex        =   93
         Top             =   4940
         Width           =   1200
      End
      Begin VB.TextBox TxtMBMIc2 
         Height          =   285
         Left            =   6000
         MaxLength       =   20
         TabIndex        =   96
         Top             =   5190
         Width           =   1200
      End
      Begin VB.TextBox txtActualAmt1 
         Height          =   285
         Left            =   4080
         TabIndex        =   98
         Top             =   5450
         Width           =   1215
      End
      Begin VB.TextBox txtActualAmt2 
         Height          =   285
         Left            =   6000
         TabIndex        =   99
         Top             =   5450
         Width           =   1215
      End
      Begin VB.TextBox txtAppAmt1 
         Height          =   285
         Left            =   4080
         TabIndex        =   100
         Top             =   5700
         Width           =   1215
      End
      Begin VB.TextBox txtAppAmt2 
         Height          =   285
         Left            =   6000
         TabIndex        =   101
         Top             =   5700
         Width           =   1215
      End
      Begin VB.TextBox TxtInvoice1 
         Height          =   285
         Left            =   4080
         MaxLength       =   20
         TabIndex        =   102
         Top             =   5950
         Width           =   1200
      End
      Begin VB.TextBox TxtInvoice2 
         Height          =   285
         Left            =   6000
         MaxLength       =   20
         TabIndex        =   103
         Top             =   5950
         Width           =   1200
      End
      Begin VB.TextBox TxtSubmission1 
         Height          =   285
         Left            =   4080
         MaxLength       =   20
         TabIndex        =   105
         Top             =   6210
         Width           =   1200
      End
      Begin VB.TextBox TxtSubmission2 
         Height          =   285
         Left            =   6000
         MaxLength       =   20
         TabIndex        =   106
         Top             =   6210
         Width           =   1200
      End
      Begin MSMask.MaskEdBox TxtDOV1 
         Height          =   315
         Left            =   4080
         TabIndex        =   55
         Top             =   1575
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   572
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOV2 
         Height          =   315
         Left            =   6000
         TabIndex        =   56
         Top             =   1570
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR1 
         Height          =   315
         Left            =   4080
         TabIndex        =   58
         Top             =   1860
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOW1 
         Height          =   315
         Left            =   4080
         TabIndex        =   61
         Top             =   2150
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR2 
         Height          =   315
         Left            =   6000
         TabIndex        =   59
         Top             =   1860
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOW2 
         Height          =   315
         Left            =   6000
         TabIndex        =   62
         Top             =   2150
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayEffDate1 
         Height          =   315
         Left            =   4080
         TabIndex        =   70
         Top             =   2880
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayExpDate1 
         Height          =   315
         Left            =   4080
         TabIndex        =   73
         Top             =   3120
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateJoin1 
         Height          =   315
         Left            =   4080
         TabIndex        =   76
         Top             =   3360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOB1 
         Height          =   315
         Left            =   4080
         TabIndex        =   79
         Top             =   3645
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayEffDate2 
         Height          =   315
         Left            =   6000
         TabIndex        =   71
         Top             =   2880
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayExpDate2 
         Height          =   315
         Left            =   6000
         TabIndex        =   74
         Top             =   3120
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateJoin2 
         Height          =   315
         Left            =   6000
         TabIndex        =   77
         Top             =   3360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOB2 
         Height          =   315
         Left            =   6000
         TabIndex        =   80
         Top             =   3645
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtBordxDate1 
         Height          =   315
         Left            =   4080
         TabIndex        =   89
         Top             =   4650
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtBordxDate2 
         Height          =   315
         Left            =   6000
         TabIndex        =   90
         Top             =   4650
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label33 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   135
         Left            =   7560
         TabIndex        =   216
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label84 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   211
         Top             =   5040
         Width           =   375
      End
      Begin VB.Label Label85 
         Caption         =   "Patient I/C"
         Height          =   255
         Left            =   3000
         TabIndex        =   210
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label Label77 
         Caption         =   "Act Amt"
         Height          =   255
         Left            =   3000
         TabIndex        =   209
         ToolTipText     =   "Approved Amt"
         Top             =   5565
         Width           =   645
      End
      Begin VB.Label Label78 
         Caption         =   "To"
         Height          =   255
         Left            =   5640
         TabIndex        =   208
         Top             =   5520
         Width           =   405
      End
      Begin VB.Label Label79 
         Caption         =   "App Amt"
         Height          =   255
         Left            =   3000
         TabIndex        =   207
         ToolTipText     =   "Approved Amt"
         Top             =   5820
         Width           =   645
      End
      Begin VB.Label Label80 
         Caption         =   "To"
         Height          =   255
         Left            =   5640
         TabIndex        =   206
         Top             =   5760
         Width           =   405
      End
      Begin VB.Label Label15 
         Caption         =   "MBM IC"
         Height          =   255
         Left            =   3000
         TabIndex        =   205
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "To"
         Height          =   255
         Left            =   5640
         TabIndex        =   204
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "To"
         Height          =   255
         Left            =   5640
         TabIndex        =   203
         Top             =   6240
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Bill No"
         Height          =   225
         Left            =   3000
         TabIndex        =   202
         Top             =   6360
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Invoice No"
         Height          =   255
         Left            =   3000
         TabIndex        =   201
         Top             =   6120
         Width           =   975
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   200
         Top             =   6000
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Branch"
         Height          =   255
         Left            =   3000
         TabIndex        =   199
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "MC Days"
         Height          =   255
         Left            =   120
         TabIndex        =   198
         Top             =   5745
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   1680
         TabIndex        =   197
         Top             =   5745
         Width           =   375
      End
      Begin VB.Label Label38 
         Caption         =   "Payor"
         Height          =   225
         Left            =   3000
         TabIndex        =   196
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   195
         Top             =   4700
         Width           =   375
      End
      Begin VB.Label Label70 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   3000
         TabIndex        =   194
         Top             =   4700
         Width           =   975
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   193
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label Label66 
         Caption         =   "Claim No"
         Height          =   255
         Left            =   3000
         TabIndex        =   192
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label65 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   191
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "Member No"
         Height          =   255
         Left            =   3000
         TabIndex        =   190
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "Supp. Type"
         Height          =   255
         Left            =   120
         TabIndex        =   189
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Policy No."
         Height          =   255
         Left            =   120
         TabIndex        =   188
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "State"
         Height          =   255
         Left            =   120
         TabIndex        =   187
         Top             =   3860
         Width           =   855
      End
      Begin VB.Label Label27 
         Caption         =   "Prem. Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   186
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Expired Status"
         Height          =   255
         Left            =   120
         TabIndex        =   185
         Top             =   3340
         Width           =   1095
      End
      Begin VB.Label Label19 
         Height          =   255
         Left            =   2520
         TabIndex        =   184
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Relationship"
         Height          =   255
         Left            =   120
         TabIndex        =   183
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Race"
         Height          =   255
         Left            =   120
         TabIndex        =   182
         Top             =   1330
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Mem Type"
         Height          =   255
         Left            =   120
         TabIndex        =   181
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Sex"
         Height          =   255
         Left            =   120
         TabIndex        =   180
         Top             =   1650
         Width           =   855
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   179
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label47 
         Caption         =   "DOV"
         Height          =   255
         Left            =   3000
         TabIndex        =   178
         ToolTipText     =   "Date of Admission"
         Top             =   1650
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   177
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   176
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   175
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "DOR"
         Height          =   255
         Left            =   3000
         TabIndex        =   174
         ToolTipText     =   "Member Value Date"
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Exp Date"
         Height          =   255
         Left            =   3000
         TabIndex        =   173
         ToolTipText     =   "Payor Expiry Date"
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   3000
         TabIndex        =   172
         ToolTipText     =   "Payor Effective Date"
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label73 
         Caption         =   "HLT Type"
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
         TabIndex        =   171
         Top             =   225
         Width           =   705
      End
      Begin VB.Label Label26 
         Caption         =   "INS Type"
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
         TabIndex        =   170
         Top             =   480
         Width           =   690
      End
      Begin VB.Label Label72 
         Caption         =   "Data Status"
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
         TabIndex        =   169
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   168
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label51 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   3000
         TabIndex        =   167
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label52 
         Caption         =   "Policy Status"
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
         TabIndex        =   166
         Top             =   2760
         Width           =   1020
      End
      Begin VB.Label Label53 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   120
         TabIndex        =   165
         Top             =   2450
         Width           =   975
      End
      Begin VB.Label Label54 
         Caption         =   "Case Status"
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
         TabIndex        =   164
         Top             =   3010
         Width           =   945
      End
      Begin VB.Label Label6 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   3000
         TabIndex        =   163
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Treat Reason"
         Height          =   255
         Left            =   3000
         TabIndex        =   162
         Top             =   2415
         Width           =   1095
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   161
         Top             =   2415
         Width           =   375
      End
      Begin VB.Label Label36 
         Caption         =   "Clinic"
         Height          =   255
         Left            =   3000
         TabIndex        =   160
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label37 
         Caption         =   "Clinic Grp"
         Height          =   255
         Left            =   3000
         TabIndex        =   159
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label40 
         Caption         =   "Date Join"
         Height          =   255
         Left            =   3000
         TabIndex        =   158
         ToolTipText     =   "Date Join"
         Top             =   3390
         Width           =   855
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   157
         Top             =   3380
         Width           =   375
      End
      Begin VB.Label Label42 
         Caption         =   "DOB"
         Height          =   255
         Left            =   3000
         TabIndex        =   156
         ToolTipText     =   "Date of Birth"
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   155
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label Label44 
         Caption         =   "Renewal "
         Height          =   255
         Left            =   120
         TabIndex        =   154
         Top             =   4180
         Width           =   1095
      End
      Begin VB.Label Label45 
         Caption         =   "Take Over"
         Height          =   255
         Left            =   120
         TabIndex        =   153
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label55 
         Caption         =   "Agent Code"
         Height          =   255
         Left            =   120
         TabIndex        =   152
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   1680
         TabIndex        =   151
         Top             =   5505
         Width           =   375
      End
      Begin VB.Label Label58 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7320
         TabIndex        =   150
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label59 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2640
         TabIndex        =   149
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label60 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   3000
         TabIndex        =   148
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label61 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   147
         Top             =   4440
         Width           =   375
      End
      Begin VB.Label Label62 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   120
         TabIndex        =   146
         Top             =   5190
         Width           =   1095
      End
      Begin VB.Label Label56 
         Caption         =   "Policy Year"
         Height          =   255
         Left            =   120
         TabIndex        =   145
         Top             =   5505
         Width           =   855
      End
      Begin VB.Label Label76 
         Caption         =   "DOW"
         Height          =   255
         Left            =   3000
         TabIndex        =   144
         ToolTipText     =   "Worksheet Date"
         Top             =   2190
         Width           =   855
      End
      Begin VB.Label Label75 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   143
         Top             =   2190
         Width           =   375
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   140
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "  Claims Reports 2 - Jupiter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   0
         TabIndex        =   141
         Top             =   0
         Width           =   3285
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Reports Type"
      Height          =   2655
      Left            =   7920
      TabIndex        =   139
      Top             =   360
      Width           =   3855
      Begin VB.OptionButton Option1 
         Caption         =   "Cost Details - Actual"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   108
         Top             =   350
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cost Details - Approved"
         Height          =   195
         Left            =   360
         TabIndex        =   109
         Top             =   650
         Width           =   2295
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Cost Details - Both"
         Height          =   195
         Left            =   360
         TabIndex        =   110
         Top             =   950
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   136
      Top             =   7320
      Width           =   11775
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   5400
         TabIndex        =   128
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9720
         TabIndex        =   133
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10560
         TabIndex        =   134
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         Height          =   375
         Left            =   8760
         TabIndex        =   132
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   375
         Left            =   7080
         TabIndex        =   130
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   375
         Left            =   7920
         TabIndex        =   131
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   6240
         TabIndex        =   129
         Top             =   160
         Width           =   855
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   240
         TabIndex        =   138
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label64 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1320
         TabIndex        =   137
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   612
      Left            =   0
      TabIndex        =   212
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1080
      ButtonWidth     =   614
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   375
      Left            =   0
      TabIndex        =   213
      Top             =   6960
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   656
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   33
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   26
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   27
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   28
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   29
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   30
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   31
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   32
         Object.Width           =   1764
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Fgrid 
      Height          =   375
      Left            =   0
      TabIndex        =   214
      Top             =   6960
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   656
      _Version        =   393216
      Rows            =   30
      Cols            =   5
      BackColorFixed  =   -2147483638
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label22 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   215
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "FrmNewClaimRpt2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private AA As Integer
Private FinalDiag, Description As String
Private GPFeeBasicMed, GPFeeMedInj, GPFeeLTMed, SPFeeBasicMed, SPFeeMedInj As String
Private SPFeeLTMed, XRay, UltraSound, ECG, Dressing, Nebulizer As String
Private ID, TSSTO, MinorSurgery, POthers, BloodDrawing As String
Private BloodGlucose, Urine, PapSmear, LOthers, DTPPolio As String
Private DTPPolioHIB, HepatitisB, MMR, IOthers, Dental, Maternity, PRVCheckup As String
Private Total, SubTotal As String
Private strAppend As String
Private strAppendCase As String
Private intExit As Integer
Private Current As String
Private GrpCompany As String
Private Clinics As String
Private m_blnDirection(1 To 33) As Boolean

Private Sub ListViewHeader1()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Claim No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "DOV"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "201(XRay)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "203(ECG)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "204(Dressing)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "206(ID)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "209(POthers)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "303(Urine)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "305(LOthers)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "404(MMR)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "405(IOthers)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "501(Dental)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "502(Maternity)"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(31).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(32).Text = "Total"
    ListView2.ColumnHeaders(32).Alignment = lvwColumnRight
    For AA = 33 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Case No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "DOA"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "DOD"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "101(R&B)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "102(ICU)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "103(HSS)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "104(OT)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "202(SURG)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "303(POST)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "401(OP)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "405(KID)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "406(CAN)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "801(MED REP)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "802(TEL)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(31).Text = "Sub-Total"
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(32).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(32).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(33).Text = "Total"
    ListView2.ColumnHeaders(33).Alignment = lvwColumnRight
End If
End Sub

Private Sub ChkCLN1_Click()
Dim strsql As String
Dim CLN As New ADODB.Recordset

cboCLNName.Clear
If ChkCLN1.Value = 1 Then
    'MediNeptuneConnCISWS.Open "File Name=" & BOSPath & "Conn\MediNeptuneConnCISWS.UDL"
    strsql = "EXEC SearchCLN"
    CLN.Open strsql, MediNeptuneConnCISWS, adOpenKeyset, adLockOptimistic
    
    Do Until CLN.EOF
        cboCLNName.AddItem CLN!CLNCode & " - " & Trim(CLN!CLNName)
        CLN.MoveNext
    Loop
    CLN.Close
    Set CLN = Nothing
    'MediNeptuneConnCISWS.Close
    'Set MediNeptuneConnCISWS = Nothing
End If
End Sub

Private Sub ChkCLNGrp1_Click()
Dim strsql As String
Dim ClinicGrp As New ADODB.Recordset

    cboCLNGrp.Clear
    If ChkCLNGrp1.Value = 1 Then
      '  MediNeptuneConnCISWS.Open "File Name=" & BOSPath & "Conn\MediNeptuneConnCISWS.UDL"
        strsql = "EXEC SearchCLNGrp"
        ClinicGrp.Open strsql, MediNeptuneConnCISWS, adOpenKeyset, adLockOptimistic
        
        Do Until ClinicGrp.EOF
            cboCLNGrp.AddItem Trim(ClinicGrp!ClinicGrpName)
            ClinicGrp.MoveNext
        Loop
        ClinicGrp.Close
        Set ClinicGrp = Nothing
    '    MediNeptuneConnCISWS.Close
    '    Set MediNeptuneConnCISWS = Nothing
    End If
    
End Sub

Private Sub ChkGrpCom1_Click()
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset

CboGrpCom.Clear
If ChkGrpCom1.Value = 1 Then
  '  MediNeptuneConnCISWS.Open "File Name=" & BOSPath & "Conn\MediNeptuneConnCISWS.UDL"
    
    strsql = "EXEC SearchPayGrp '" & Mid(cboPayorCode, 1, 2) & "'"
    PAYGRP.Open strsql, MediNeptuneConnCISWS, adOpenKeyset, adLockOptimistic
    
    Do Until PAYGRP.EOF
        CboGrpCom.AddItem Trim(PAYGRP!GrpCoName)
        PAYGRP.MoveNext
    Loop
    
    PAYGRP.Close
    Set PAYGRP = Nothing
  '  MediNeptuneConnCISWS.Close
  '  Set MediNeptuneConnCISWS = Nothing
End If

End Sub

Private Sub ChkPayor1_Click()
Dim strsql As String
Dim PAY As New ADODB.Recordset

    cboPayorCode.Clear
    If ChkPayor1.Value = 1 Then
       ' MediNeptuneConnCISWS.Open "File Name=" & BOSPath & "Conn\MediNeptuneConnCISWS.UDL"
            LoadPayor (cboHLTCode)
       ' MediNeptuneConnCISWS.Close
       ' Set MediNeptuneConnCISWS = Nothing
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

Private Sub CboHLTCode_Change()
    cboPayorCode.Clear
    If InStr(cboHLTCode.Text, "null") Then
       cboHLTCode.Enabled = False
    End If
End Sub

Private Sub LoadPayor(Optional HLTCode As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim PAY As New ADODB.Recordset
Dim strsql As String

    If Mid(ChkStr(HLTCode), 1, 1) = "N" Then
        strAppendCase = "41"
        'sqlPAY = "select PAYCode, PAYCompanyName from PAY "
        'sqlPAY = sqlPAY + " Where HLTCODE = 'N' "
    ElseIf Mid(ChkStr(HLTCode), 1, 1) = "S" Then
        strAppendCase = "42"
        'sqlPAY = "select PAYCode, PAYCompanyName from PAY "
        'sqlPAY = sqlPAY + " Where HLTCODE = 'S' "
    Else
        strsql = "EXEC SearchPayor"
        PAY.Open strsql, MediNeptuneConnCISWS, adOpenKeyset, adLockOptimistic
    
        Do Until PAY.EOF
            cboPayorCode.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
            PAY.MoveNext
        Loop
        PAY.Close
        Set PAY = Nothing
        Exit Sub
    End If

    strAppend = ""
    
    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
    
    'PAY.Open sqlPAY, MediNeptuneConnCISDB, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
    Do Until PAY.EOF
        cboPayorCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
End Sub


Private Sub CmdBack_Click()
    
    
    If Option3.Value = True Then
        Fgrid.Height = 375
        Fgrid.Left = 0
        Fgrid.Top = 6960
        Fgrid.Width = 11775
        Frame22.Visible = True
        Frame4.Visible = True
        Frame6.Visible = True
    Else
        ListView2.Height = 375
        ListView2.Left = 0
        ListView2.Top = 6960
        ListView2.Width = 11775
        Frame22.Visible = True
        Frame4.Visible = True
        Frame6.Visible = True
    End If
    
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    cmdExit.Enabled = True
    CmdClear.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub CmdView_Click()
    
    If Option3.Value = True Then
        Fgrid.Height = 6720
        Fgrid.Left = 0
        Fgrid.Top = 580
        Fgrid.Width = 11775
        ListView2.Visible = False
        Frame22.Visible = False
        Frame4.Visible = False
        Frame6.Visible = False
    Else
        ListView2.Height = 6720
        ListView2.Left = 0
        ListView2.Top = 580
        ListView2.Width = 11775
        Frame22.Visible = False
        Frame4.Visible = False
        Frame6.Visible = False
        Fgrid.Visible = False
    End If
End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "FrmNewClaimRpt2")
    intExit = 0
    'MediNeptuneConnCISWS.Open "File Name=" & BOSPath & "Conn\MediNeptuneConnCISWS.UDL"
        Call ComboListing
        Call ListViewHeader1
        Call DoInitialSettings
        ListView2.Visible = True
  '  MediNeptuneConnCISWS.Close
   ' Set MediNeptuneConnCISWS = Nothing
    
End Sub

' =================== Command Button Click =============================

Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
    
        If Option1.Value = True Or Option2.Value = True Then
            Call ExportToExcelAnalysis1(ListView2)
        ElseIf Option3.Value = True Then
            Call ExportToExcelWorksheet
        Else
        
        End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
Dim strAppend As String
    strAppend = ""
    lblCurrent = ""
    lblCurrent = ""
    
    CmdPrint.Enabled = False
    CmdClear.Enabled = False
    cmdExit.Enabled = False
    
    If Option3.Value = True Then
        Fgrid.Height = 6735
        Fgrid.Left = 0
        Fgrid.Top = 580
        Fgrid.Width = 11775
        ListView2.Visible = False
        Frame22.Visible = False
        Frame4.Visible = False
        Frame6.Visible = False
    Else
        ListView2.Height = 6720
        ListView2.Left = 0
        ListView2.Top = 580
        ListView2.Width = 11775
        Frame22.Visible = False
        Frame4.Visible = False
        Frame6.Visible = False
        Fgrid.Visible = False
    End If
    
    Screen.MousePointer = vbHourglass
   ' MediNeptuneConnCISWS.Open "File Name=" & BOSPath & "Conn\MediNeptuneConnCISWS.UDL"
        Call SearchRecords
   ' MediNeptuneConnCISWS.Close
   ' Set MediNeptuneConnCISWS = Nothing
    Screen.MousePointer = vbDefault
End Sub

Public Function ClearForm(frmCall)
On Error Resume Next
Dim ctl As Control

For Each ctl In frmCall.Controls
    If TypeOf ctl Is TextBox Then
        ctl.Text = ""
    ElseIf TypeOf ctl Is ListBox Then
            ctl.Clear
    ElseIf TypeOf ctl Is ComboBox Then
            ctl.Text = ""
    ElseIf TypeOf ctl Is checkbox Then
            ctl.Value = False
    ElseIf TypeOf ctl Is MaskEdBox Then
        ctl.Text = "__/__/____"
    
    End If
Next ctl
End Function

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lblCurrent = ""
    lblCurrent = ""
    ListView2.ListItems.Clear
End Sub
' =================== Command Button Click / Change ==========================


'============== ComboBox Listing =================================
Private Sub ComboListing()
Dim PolStatus As New ADODB.Recordset
Dim state As New ADODB.Recordset
Dim strsql As String
Dim strsql2 As String
        
    cboHLTCode.AddItem "S" & " - " & "Sihat"
    cboHLTCode.AddItem "N" & " - " & "Non-Sihat"
    
    CboInsType.AddItem "I" & " - " & "Individual"
    CboInsType.AddItem "F" & " - " & "Family"
    CboInsType.AddItem "G" & " - " & "Group Individual"
    CboInsType.AddItem "H" & " - " & "Group Family"
    
    CboMemType.AddItem "Both"
    CboMemType.AddItem "P" & " - " & "Principal"
    CboMemType.AddItem "S" & " - " & "Supplementary"
    CboMemType.Text = "Both"
    
    CboRace.AddItem "C - Chinese"
    CboRace.AddItem "I - Indian"
    CboRace.AddItem "M - Malay"
    CboRace.AddItem "O - Other"
            
    CboSex.AddItem "M" & " - " & "Male"
    CboSex.AddItem "F" & " - " & "Female"
    
    CboRelation.AddItem "D - Daughter"
    CboRelation.AddItem "H - Husband"
    CboRelation.AddItem "S - Son"
    CboRelation.AddItem "W - Wife"
    
    cboDataStatus.AddItem "A - Active"
    cboDataStatus.AddItem "D - Delete"
    
    With CboClaimability
        .AddItem "A - Accepted"
        .AddItem "R - Rejected"
    End With
    
    strsql = "EXEC SearchPolStatus"
    PolStatus.Open strsql, MediNeptuneConnCISWS, adOpenKeyset, adLockOptimistic
    
    Do Until PolStatus.EOF
        CboStatus.AddItem (PolStatus!pstatuscode) & " - " & (PolStatus!pstatusDesc)
        PolStatus.MoveNext
    Loop
    PolStatus.Close
    Set PolStatus = Nothing
    
    With CboCaseStatus
        .AddItem "O - Open"
        .AddItem "D - Delete"
        .AddItem "A - All"
    End With
    
    CboExpired.AddItem "N - No"
    CboExpired.AddItem "Y - Yes"
            
    CboPreMode.AddItem "A - Annually"
    CboPreMode.AddItem "Q - Quarterly"
    CboPreMode.AddItem "S - Semi Annual"
    CboPreMode.AddItem "M - Monthly"
        
    strsql2 = " EXEC SearchState"
    state.Open strsql2, MediNeptuneConnCISWS, adOpenKeyset, adLockOptimistic
    
    Do Until state.EOF
        CboState.AddItem ChkStr(state!STADescription)
        state.MoveNext
    Loop
    state.Close
    Set state = Nothing
        
    cboRenewalInd.AddItem "N - New"
    cboRenewalInd.AddItem "P - Partial"
    cboRenewalInd.AddItem "R - ReNew"
    cboRenewalInd.AddItem "T - Timely"
    
    cboTakeOver.AddItem "N - None"
    cboTakeOver.AddItem "Y - Take Over"
    cboTakeOver.AddItem "C - Conversion"
               
    cboSuppType.AddItem "Both"
    cboSuppType.AddItem "C" & " - " & "Combined Limit"
    cboSuppType.AddItem "S" & " - " & "Separated Limit"
    cboSuppType.Text = "Both"
    
End Sub
'============== ComboBox Listing =================================

'=========== Option Button, Combo Box Change/KeyPress ===================
Private Sub cboPayorCode_Change()

    If cboPayorCode.Text = "" Then
        If InStr(cboPayorCode.Text, "null") Then
           cboPayorCode.Enabled = False
        End If
    End If
End Sub

Private Sub cboPayorCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPayorCode.Text, cboPayorCode.SelStart) + Chr(KeyAscii) + Mid(cboPayorCode.Text, cboPayorCode.SelStart + cboPayorCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPayorCode.ListCount - 1
 
        If NewText = LCase(Left(cboPayorCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPayorCode.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPayorCode.Text = ValidValue
                cboPayorCode.SelStart = 0
                cboPayorCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub



Private Sub CboHLTCode_KeyPress(KeyAscii As Integer)
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

Private Sub CboInsType_Change()

    If CboInsType.Text = "" Then
        If InStr(CboInsType.Text, "null") Then
           CboInsType.Enabled = False
        End If
    End If
End Sub

Private Sub CboInsType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

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

Private Sub CboMemType_Change()

    If CboMemType.Text = "" Then
        If InStr(CboMemType.Text, "null") Then
           CboMemType.Enabled = False
        End If
    End If
End Sub

Private Sub cboSuppType_Change()

    If cboSuppType.Text = "" Then
        If InStr(cboSuppType.Text, "null") Then
           cboSuppType.Enabled = False
        End If
    End If
End Sub







Private Sub CboRace_Change()

    If CboRace.Text = "" Then
        If InStr(CboRace.Text, "null") Then
           CboRace.Enabled = False
        End If
    End If
End Sub

Private Sub CboRace_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRace.Text, CboRace.SelStart) + Chr(KeyAscii) + Mid(CboRace.Text, CboRace.SelStart + CboRace.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboRace.ListCount - 1
 
        If NewText = LCase(Left(CboRace.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboRace.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboRace.Text = ValidValue
                CboRace.SelStart = 0
                CboRace.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboSex_Change()

    If CboSex.Text = "" Then
        If InStr(CboSex.Text, "null") Then
           CboSex.Enabled = False
        End If
    End If
End Sub

Private Sub CboSex_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboSex.Text, CboSex.SelStart) + Chr(KeyAscii) + Mid(CboSex.Text, CboSex.SelStart + CboSex.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboSex.ListCount - 1
 
        If NewText = LCase(Left(CboSex.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboSex.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboSex.Text = ValidValue
                CboSex.SelStart = 0
                CboSex.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboRelation_Change()

    If CboRelation.Text = "" Then
        If InStr(CboRelation.Text, "null") Then
           CboRelation.Enabled = False
        End If
    End If
End Sub

Private Sub CboRelation_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRelation.Text, CboRelation.SelStart) + Chr(KeyAscii) + Mid(CboRelation.Text, CboRelation.SelStart + CboRelation.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboRelation.ListCount - 1
 
        If NewText = LCase(Left(CboRelation.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboRelation.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboRelation.Text = ValidValue
                CboRelation.SelStart = 0
                CboRelation.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboDataStatus_Change()

    If cboDataStatus.Text = "" Then
        If InStr(cboDataStatus.Text, "null") Then
           cboDataStatus.Enabled = False
        End If
    End If
End Sub

Private Sub cboDataStatus_KeyPress(KeyAscii As Integer)
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
End Sub

Private Sub CboClaimability_Change()

    If CboClaimability.Text = "" Then
        If InStr(CboClaimability.Text, "null") Then
           CboClaimability.Enabled = False
        End If
    End If
End Sub

Private Sub CboClaimability_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClaimability.Text, CboClaimability.SelStart) + Chr(KeyAscii) + Mid(CboClaimability.Text, CboClaimability.SelStart + CboClaimability.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClaimability.ListCount - 1
 
        If NewText = LCase(Left(CboClaimability.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClaimability.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClaimability.Text = ValidValue
                CboClaimability.SelStart = 0
                CboClaimability.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboStatus_Change()

    If CboStatus.Text = "" Then
        If InStr(CboStatus.Text, "null") Then
           CboStatus.Enabled = False
        End If
    End If
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboStatus.Text, CboStatus.SelStart) + Chr(KeyAscii) + Mid(CboStatus.Text, CboStatus.SelStart + CboStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboStatus.ListCount - 1
 
        If NewText = LCase(Left(CboStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboStatus.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboStatus.Text = ValidValue
                CboStatus.SelStart = 0
                CboStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboCaseStatus_Change()

    If CboCaseStatus.Text = "" Then
        If InStr(CboCaseStatus.Text, "null") Then
           CboCaseStatus.Enabled = False
        End If
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

Private Sub CboExpired_Change()

    If CboExpired.Text = "" Then
        If InStr(CboExpired.Text, "null") Then
           CboExpired.Enabled = False
        End If
    End If
End Sub

Private Sub CboExpired_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboExpired.Text, CboExpired.SelStart) + Chr(KeyAscii) + Mid(CboExpired.Text, CboExpired.SelStart + CboExpired.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboExpired.ListCount - 1
 
        If NewText = LCase(Left(CboExpired.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboExpired.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboExpired.Text = ValidValue
                CboExpired.SelStart = 0
                CboExpired.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboPreMode_Change()

    If CboPreMode.Text = "" Then
        If InStr(CboPreMode.Text, "null") Then
           CboPreMode.Enabled = False
        End If
    End If
End Sub

Private Sub CboPreMode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPreMode.Text, CboPreMode.SelStart) + Chr(KeyAscii) + Mid(CboPreMode.Text, CboPreMode.SelStart + CboPreMode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPreMode.ListCount - 1
 
        If NewText = LCase(Left(CboPreMode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPreMode.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPreMode.Text = ValidValue
                CboPreMode.SelStart = 0
                CboPreMode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboState_Change()

    If CboState.Text = "" Then
        If InStr(CboState.Text, "null") Then
           CboState.Enabled = False
        End If
    End If
End Sub

Private Sub CboState_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboState.Text, CboState.SelStart) + Chr(KeyAscii) + Mid(CboState.Text, CboState.SelStart + CboState.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboState.ListCount - 1
 
        If NewText = LCase(Left(CboState.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboState.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboState.Text = ValidValue
                CboState.SelStart = 0
                CboState.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboRenewalInd_Change()

    If cboRenewalInd.Text = "" Then
        If InStr(cboRenewalInd.Text, "null") Then
           cboRenewalInd.Enabled = False
        End If
    End If
End Sub

Private Sub cboRenewalInd_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboRenewalInd.Text, cboRenewalInd.SelStart) + Chr(KeyAscii) + Mid(cboRenewalInd.Text, cboRenewalInd.SelStart + cboRenewalInd.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboRenewalInd.ListCount - 1
 
        If NewText = LCase(Left(cboRenewalInd.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboRenewalInd.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboRenewalInd.Text = ValidValue
                cboRenewalInd.SelStart = 0
                cboRenewalInd.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboTakeOver_Change()

    If cboTakeOver.Text = "" Then
        If InStr(cboTakeOver.Text, "null") Then
           cboTakeOver.Enabled = False
        End If
    End If
End Sub

Private Sub cboTakeOver_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboTakeOver.Text, cboTakeOver.SelStart) + Chr(KeyAscii) + Mid(cboTakeOver.Text, cboTakeOver.SelStart + cboTakeOver.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboTakeOver.ListCount - 1
 
        If NewText = LCase(Left(cboTakeOver.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboTakeOver.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboTakeOver.Text = ValidValue
                cboTakeOver.SelStart = 0
                cboTakeOver.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboGrpCom_Change()

    If CboGrpCom.Text = "" Then
        If InStr(CboGrpCom.Text, "null") Then
           CboGrpCom.Enabled = False
        End If
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

Private Sub cboCLNName_Change()

    If cboCLNName.Text = "" Then
        If InStr(cboCLNName.Text, "null") Then
           cboCLNName.Enabled = False
        End If
    End If
End Sub

Private Sub cboCLNName_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNName.Text, cboCLNName.SelStart) + Chr(KeyAscii) + Mid(cboCLNName.Text, cboCLNName.SelStart + cboCLNName.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCLNName.ListCount - 1
 
        If NewText = LCase(Left(cboCLNName.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNName.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCLNName.Text = ValidValue
                cboCLNName.SelStart = 0
                cboCLNName.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCLNGrp_Change()

    If cboCLNGrp.Text = "" Then
        If InStr(cboCLNGrp.Text, "null") Then
           cboCLNGrp.Enabled = False
        End If
    End If
End Sub

Private Sub cboCLNGrp_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNGrp.Text, cboCLNGrp.SelStart) + Chr(KeyAscii) + Mid(cboCLNGrp.Text, cboCLNGrp.SelStart + cboCLNGrp.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCLNGrp.ListCount - 1
 
        If NewText = LCase(Left(cboCLNGrp.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNGrp.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCLNGrp.Text = ValidValue
                cboCLNGrp.SelStart = 0
                cboCLNGrp.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)

    If Option4.Value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(30)
        Case 31
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(31)
        Case 32
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(32)
        End Select
    
    ElseIf Option5.Value = True Or Option7.Value = True Or _
           Option10.Value = True Or Option11.Value = True Or _
           Option12.Value = True Or Option16.Value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(30)
        End Select
        
    ElseIf Option6.Value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(30)
        End Select
        
    ElseIf Option8.Value = True Or Option19.Value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(29)
        End Select
        
    ElseIf Option9.Value = True Or Option13.Value = True Or _
           Option15.Value = True Or Option17.Value = True Or _
           Option18.Value = True Or Option20.Value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(29)
        End Select
        
    ElseIf Option14.Value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(30)
        Case 31
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(31)
        End Select
    End If

End Sub


Private Sub Option1_Click()
        
    Fgrid.Visible = False
    ListView2.Visible = True
    ListView2.Height = 375
    ListView2.Left = 0
    ListView2.Top = 6960
    ListView2.Width = 11775
    Frame22.Visible = True
    Frame4.Visible = True
    Frame6.Visible = True
        
    Option3.Visible = True
    Option4.Visible = True
    Option5.Visible = True
    Option6.Visible = True
    Option7.Visible = True
    Option8.Visible = True
    Option9.Visible = True
    Option10.Visible = True
    Option11.Visible = True
    Option12.Visible = True
    Option13.Visible = True
    Option14.Visible = True
    Option15.Visible = True
    Option13.Visible = True
    Option14.Visible = True
    Option16.Visible = True
    Option17.Visible = True
    Option18.Visible = True
    Option19.Visible = True
    Option20.Visible = True

ListView2.ColumnHeaders(1).Text = "Case No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "DOV"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "101(R&B)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "102(ICU)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "103(HSS)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "104(OT)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "202(SURG)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "303(POST)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "401(OP)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "405(KID)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "406(CAN)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "801(MED REP)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "802(TEL)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Sub-Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(31).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(32).Text = "Total"
    ListView2.ColumnHeaders(32).Alignment = lvwColumnRight
    
    For AA = 33 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub Option10_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "AgeBand"
    ListView2.ColumnHeaders(2).Text = "Sex"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "201(XRay)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "203(ECG)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "204(Dressing)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "206(ID)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "209(POthers)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "303(Urine)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "305(LOthers)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "404(MMR)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "405(IOthers)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "501(Dental)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "502(Maternity)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Ageband"
    ListView2.ColumnHeaders(2).Text = "Sex"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(R&B)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(ICU)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(HSS)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(OT)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(SURG)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "303(POST)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "401(OP)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "405(KID)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "406(CAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "801(MED REP)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "802(TEL)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Sub-Total"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub Option11_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Diagnosis"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "201(XRay)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "203(ECG)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "204(Dressing)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "206(ID)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "209(POthers)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "303(Urine)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "305(LOthers)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "404(MMR)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "405(IOthers)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "501(Dental)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "502(Maternity)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Final Diag"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "101(R&B)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(ICU)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(HSS)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(OT)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(SURG)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "303(POST)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "401(OP)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "405(KID)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "406(CAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "801(MED REP)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "802(TEL)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Sub-Total"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub Option12_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Clinic Code"
    ListView2.ColumnHeaders(2).Text = "Clinic Name"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "201(XRay)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "203(ECG)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "204(Dressing)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "206(ID)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "209(POthers)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "303(Urine)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "305(LOthers)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "404(MMR)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "405(IOthers)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "501(Dental)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "502(Maternity)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "101(R&B)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(ICU)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(HSS)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(OT)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "202(SURG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "303(POST)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "401(OP)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "405(KID)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "406(CAN)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(MED REP)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(TEL)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Sub-Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub Option13_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(XRay)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ECG)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Dressing)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "206(ID)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "209(POthers)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "303(Urine)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "305(LOthers)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "404(MMR)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "405(IOthers)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "501(Dental)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "502(Maternity)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Claim Type"
    ListView2.ColumnHeaders(2).Text = "101(R&B)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(ICU)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(HSS)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(OT)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "202(SURG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "303(POST)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "401(OP)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "405(KID)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "406(CAN)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(MED REP)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(TEL)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Sub-Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub Option14_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Plan Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "201(XRay)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "203(ECG)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "204(Dressing)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "206(ID)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "209(POthers)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "303(Urine)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "305(LOthers)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "404(MMR)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "405(IOthers)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "501(Dental)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "502(Maternity)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(31).Text = "Total"
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    
    For AA = 32 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Claim Nature"
    ListView2.ColumnHeaders(2).Text = "101(R&B)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(ICU)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(HSS)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(OT)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "202(SURG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "303(POST)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "401(OP)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "405(KID)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "406(CAN)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(MED REP)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(TEL)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Sub-Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub Option15_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Group Company"
    ListView2.ColumnHeaders(2).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(XRay)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ECG)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Dressing)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "206(ID)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "209(POthers)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "303(Urine)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "305(LOthers)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "404(MMR)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "405(IOthers)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "501(Dental)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "502(Maternity)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Claim Nature"
    ListView2.ColumnHeaders(2).Text = "Sex"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(R&B)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(ICU)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(HSS)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(OT)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(SURG)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "303(POST)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "401(OP)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "405(KID)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "406(CAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "801(MED REP)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "802(TEL)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Sub-Total"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub





Private Sub Option16_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Relation"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "201(XRay)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "203(ECG)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "204(Dressing)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "206(ID)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "209(POthers)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "303(Urine)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "305(LOthers)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "404(MMR)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "405(IOthers)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "501(Dental)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "502(Maternity)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Relation"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(R&B)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(ICU)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(HSS)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(OT)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(SURG)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "303(POST)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "401(OP)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "405(KID)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "406(CAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "801(MED REP)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "802(TEL)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Sub-Total"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub Option17_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(XRay)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ECG)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Dressing)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "206(ID)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "209(POthers)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "303(Urine)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "305(LOthers)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "404(MMR)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "405(IOthers)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "501(Dental)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "502(Maternity)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "101(R&B)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(ICU)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(HSS)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(OT)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "202(SURG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "303(POST)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "401(OP)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "405(KID)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "406(CAN)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(MED REP)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(TEL)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Sub-Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub Option2_Click()
    
    Fgrid.Visible = False
    ListView2.Visible = True
    ListView2.Height = 375
    ListView2.Left = 0
    ListView2.Top = 6960
    ListView2.Width = 11775
    Frame22.Visible = True
    Frame4.Visible = True
    Frame6.Visible = True

    Option3.Visible = True
    Option4.Visible = True
    Option5.Visible = True
    Option6.Visible = True
    Option7.Visible = True
    Option8.Visible = True
    Option9.Visible = True
    Option10.Visible = True
    Option11.Visible = True
    Option12.Visible = True
    Option13.Visible = True
    Option14.Visible = True
    Option15.Visible = True
    Option13.Visible = True
    Option14.Visible = True
    Option16.Visible = True
    Option17.Visible = True
    Option18.Visible = True
    Option19.Visible = True
    Option20.Visible = True
    
ListView2.ColumnHeaders(1).Text = "Case No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "DOV"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "101(R&B)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "102(ICU)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "103(HSS)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "104(OT)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "202(SURG)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "303(POST)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "401(OP)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "405(KID)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "406(CAN)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "801(MED REP)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "802(TEL)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Sub-Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(31).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(32).Text = "Total"
    ListView2.ColumnHeaders(32).Alignment = lvwColumnRight
    
    For AA = 33 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End Sub

Private Sub Option18_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Bordx No"
    ListView2.ColumnHeaders(2).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(XRay)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ECG)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Dressing)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "206(ID)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "209(POthers)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "303(Urine)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "305(LOthers)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "404(MMR)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "405(IOthers)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "501(Dental)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "502(Maternity)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Bordx No"
    ListView2.ColumnHeaders(2).Text = "101(R&B)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(ICU)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(HSS)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(OT)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "202(SURG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "303(POST)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "401(OP)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "405(KID)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "406(CAN)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(MED REP)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(TEL)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Sub-Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub Option19_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Policy Year"
    ListView2.ColumnHeaders(2).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(XRay)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ECG)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Dressing)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "206(ID)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "209(POthers)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "303(Urine)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "305(LOthers)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "404(MMR)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "405(IOthers)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "501(Dental)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "502(Maternity)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Policy Year"
    ListView2.ColumnHeaders(2).Text = "101(R&B)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(ICU)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(HSS)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(OT)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "202(SURG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "303(POST)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "401(OP)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "405(KID)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "406(CAN)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(MED REP)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(TEL)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Sub-Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub Option20_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Take Over"
    ListView2.ColumnHeaders(2).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(XRay)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ECG)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Dressing)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "206(ID)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "209(POthers)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "303(Urine)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "305(LOthers)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "404(MMR)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "405(IOthers)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "501(Dental)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "502(Maternity)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Takeover"
    ListView2.ColumnHeaders(2).Text = "101(R&B)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(ICU)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(HSS)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(OT)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "202(SURG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "303(POST)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "401(OP)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "405(KID)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "406(CAN)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(MED REP)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(TEL)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Sub-Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub Option3_Click()
    
    ListView2.Visible = False
    Fgrid.Visible = True
    Fgrid.Height = 375
    Fgrid.Left = 0
    Fgrid.Top = 6960
    Fgrid.Width = 11775
    Frame22.Visible = True
    Frame4.Visible = True
    Frame6.Visible = True
        
    
    Option4.Visible = False
    Option5.Visible = False
    Option6.Visible = False
    Option7.Visible = False
    Option8.Visible = False
    Option9.Visible = False
    Option10.Visible = False
    Option11.Visible = False
    Option12.Visible = False
    Option13.Visible = False
    Option14.Visible = False
    Option15.Visible = False
    Option13.Visible = False
    Option14.Visible = False
    Option16.Visible = False
    Option17.Visible = False
    Option18.Visible = False
    Option19.Visible = False
    Option20.Visible = False
End Sub



Private Sub Option4_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Claim No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "DOV"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "101(R&B)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "102(ICU)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "103(HSS)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "104(OT)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "202(SURG)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "303(POST)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "401(OP)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "405(KID)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "406(CAN)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "801(MED REP)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "802(TEL)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Sub-Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(31).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(32).Text = "Total"
    ListView2.ColumnHeaders(32).Alignment = lvwColumnRight
    
    For AA = 33 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Claim No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "DOV"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "101(R&B)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "102(ICU)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "103(HSS)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "104(OT)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "202(SURG)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "303(POST)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "401(OP)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "405(KID)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "406(CAN)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "801(MED REP)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "802(TEL)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Sub-Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(31).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(32).Text = "Total"
    ListView2.ColumnHeaders(32).Alignment = lvwColumnRight
    
    For AA = 33 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End If
End Sub

Private Sub Option5_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "201(XRay)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "203(ECG)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "204(Dressing)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "206(ID)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "209(POthers)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "303(Urine)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "305(LOthers)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "404(MMR)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "405(IOthers)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "501(Dental)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "502(Maternity)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "101(R&B)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(ICU)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(HSS)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(OT)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(SURG)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "303(POST)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "401(OP)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "405(KID)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "406(CAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "801(MED REP)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "802(TEL)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Sub-Total"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub Option6_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "AgeBand"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "201(XRay)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "203(ECG)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "204(Dressing)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "206(ID)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "209(POthers)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "303(Urine)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "305(LOthers)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "404(MMR)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "405(IOthers)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "501(Dental)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "502(Maternity)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Ageband"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(R&B)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(ICU)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(HSS)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(OT)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(SURG)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "303(POST)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "401(OP)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "405(KID)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "406(CAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "801(MED REP)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "802(TEL)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Sub-Total"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub Option7_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Sex"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "201(XRay)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "203(ECG)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "204(Dressing)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "206(ID)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "209(POthers)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "303(Urine)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "305(LOthers)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "404(MMR)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "405(IOthers)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "501(Dental)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "502(Maternity)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Sex"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(R&B)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(ICU)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(HSS)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(OT)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(SURG)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "303(POST)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "401(OP)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "405(KID)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "406(CAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "801(MED REP)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "802(TEL)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Sub-Total"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub Option8_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Ageband"
    ListView2.ColumnHeaders(2).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(XRay)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ECG)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Dressing)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "206(ID)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "209(POthers)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "303(Urine)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "305(LOthers)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "404(MMR)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "405(IOthers)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "501(Dental)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "502(Maternity)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Ageband"
    ListView2.ColumnHeaders(2).Text = "101(R&B)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(ICU)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(HSS)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(OT)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "202(SURG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "303(POST)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "401(OP)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "405(KID)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "406(CAN)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(MED REP)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(TEL)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Sub-Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub Option9_Click()
If Option1.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Sex"
    ListView2.ColumnHeaders(2).Text = "101(GPFeeBasicMed)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(GPFeeMedInj)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(GPFeeLTMed)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(SPFeeBasicMed)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(SPFeeMedInj)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(XRay)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(UltraSound)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(ECG)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Dressing)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "205(Bebulizer)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "206(ID)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "207(TSSTO)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "208(MinorSurgery)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "209(POthers)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "301(BloodDrawing)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "302(BloodGlucose)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "303(Urine)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "304(PapSmear)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "305(LOthers)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "401(DTPPolio)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "402(DTPPolioHIB)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "403(HepatitisB)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "404(MMR)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "405(IOthers)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "501(Dental)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "502(Maternity)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "503(PRVCheckup)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf Option2.Value = True Then
    ListView2.ColumnHeaders(1).Text = "Sex"
    ListView2.ColumnHeaders(2).Text = "101(R&B)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(ICU)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(HSS)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(OT)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "201(PRE-DIAG)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "202(SURG)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "203(ANAEST)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "301(PRE-SPEC)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "302(PHYS VISIT)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "303(POST)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "401(OP)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "402(AMBUL)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "403(PHYSIO)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "404(KID&CAN)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "405(KID)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "406(CAN)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(ORGAN)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(GOVT.INC)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(GUARDIAN)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(MED REP)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(TEL)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(ADMIN)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "902(MATERNITY)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "903(OTHERS)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Sub-Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Limit Exceeded"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    For AA = 30 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

'=========== Check Date Format ===========================

Private Sub TxtDOV1_LostFocus()
    If IsDate(TxtDOV1) Then
    Else
        If TxtDOV1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOV1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOV2_LostFocus()
    If IsDate(TxtDOV2) Then
    Else
        If TxtDOV2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOV2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOR1_LostFocus()
    If IsDate(TxtDOR1) Then
    Else
        If TxtDOR1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOR1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOR2_LostFocus()
    If IsDate(TxtDOR2) Then
    Else
        If TxtDOR2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOR2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOW1_LostFocus()
    If IsDate(txtDOW1) Then
    Else
        If txtDOW1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOW1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOW2_LostFocus()
    If IsDate(txtDOW2) Then
    Else
        If txtDOW2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOW2.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayEffDate1_LostFocus()
    If IsDate(TxtPayEffDate1) Then
    Else
        If TxtPayEffDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayEffDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayEffDate2_LostFocus()
    If IsDate(TxtPayEffDate2) Then
    Else
        If TxtPayEffDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayEffDate2.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayExpDate1_LostFocus()
    If IsDate(TxtPayExpDate1) Then
    Else
        If TxtPayExpDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayExpDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayExpDate2_LostFocus()
    If IsDate(TxtPayExpDate2) Then
    Else
        If TxtPayExpDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayExpDate2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDateJoin1_LostFocus()
    If IsDate(txtDateJoin1) Then
    Else
        If txtDateJoin1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateJoin1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDateJoin2_LostFocus()
    If IsDate(txtDateJoin2) Then
    Else
        If txtDateJoin2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateJoin2.SetFocus
        End If
    End If
End Sub

Private Sub txtDOB1_LostFocus()
    If IsDate(txtDOB1) Then
    Else
        If txtDOB1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOB1.SetFocus
        End If
    End If
End Sub

Private Sub txtDOB2_LostFocus()
    If IsDate(txtDOB2) Then
    Else
        If txtDOB2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOB2.SetFocus
        End If
    End If
End Sub

Private Sub TxtBordxDate1_LostFocus()
    If IsDate(TxtBordxDate1) Then
    Else
        If TxtBordxDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtBordxDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtBordxDate2_LostFocus()
    If IsDate(TxtBordxDate2) Then
    Else
        If TxtBordxDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtBordxDate2.SetFocus
        End If
    End If
End Sub
'=========== Check Date Format ===========================




Private Function SearchRecords()
Dim sql As String
Dim strAppend As String

    strAppend = ""
    GrpCompany = ""
    Clinics = ""

    If Len(Mid(cboHLTCode, 1, 1)) Then
        strAppend = strAppend + " AND HLTCode ='" & Trim(Mid(cboHLTCode, 1, 1)) & "'"
    End If
    
    If ChkHLT.Value = 1 Then
        strAppend = strAppend & " AND (HLTCode Is null OR HLTCode = '')"
    End If
    
    If Len(Mid(CboInsType, 1, 1)) Then
        strAppend = strAppend & " AND INSCode = '" & Trim(Mid(CboInsType, 1, 1)) & "'"
    End If
    
    If ChkINS.Value = 1 Then
        strAppend = strAppend & " AND (INSCode Is null OR INSCode = '')"
    End If
       
    If Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(cboSuppType) = "Both" Then
        strAppend = strAppend & " AND (MBMMemberType = 'P' Or MBMMemberType = 'Q') And MBMPatientCovID = '00' "
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
        strAppend = strAppend & " AND MBMMemberType = 'P'  And MBMPatientCovID = '00'"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
        strAppend = strAppend & " AND MBMMemberType = 'Q' And MBMPatientCovID = '00'"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(cboSuppType) = "Both" Then
        strAppend = strAppend & " AND (MBMCMemberType = 'S' OR MBMCMemberType = 'C')"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
        strAppend = strAppend & " AND MBMCMemberType = 'C'"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
        strAppend = strAppend & " AND MBMCMemberType = 'S'"
    ElseIf Trim(CboMemType) = "Both" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
        strAppend = strAppend & " AND (MBMMemberType = 'P' OR MBMCMemberType = 'C')"
    ElseIf Trim(CboMemType) = "Both" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
        strAppend = strAppend & " AND (MBMMemberType = 'Q' OR MBMCMemberType = 'S')"
    End If
    
    If Len(Mid(CboRace, 1, 1)) Then
        strAppend = strAppend & " AND RACCode = '" & Trim(Mid(CboRace, 1, 1)) & "'"
    End If
    
    If ChkRace.Value = 1 Then
        strAppend = strAppend & " AND (RACCode Is null OR RACCode = '')"
    End If
    
    If Len(Mid(CboSex, 1, 1)) Then
        strAppend = strAppend & " AND Sex = '" & Trim(Mid(CboSex, 1, 1)) & "'"
    End If
    
    If ChkSex.Value = 1 Then
        strAppend = strAppend & " AND (Sex Is null OR Sex = '')"
    End If
    
    If Len(Mid(CboRelation, 1, 1)) Then
        strAppend = strAppend & " AND Relation = '" & Trim(Mid(CboRelation, 1, 1)) & "'"
    End If
    
    If ChkRelation.Value = 1 Then
        strAppend = strAppend & " AND (Relation Is null OR Relation = '')"
    End If
    
    If Len(Mid(cboDataStatus, 1, 1)) Then
        strAppend = strAppend & " AND MBMDataStatus = '" & Trim(Mid(cboDataStatus, 1, 1)) & "'"
    End If
    
    If ChkDataStatus.Value = 1 Then
        strAppend = strAppend & " AND (MBMDataStatus Is null OR MBMDataStatus = '')"
    End If
    
    If Len(Mid(CboClaimability, 1, 1)) Then
        strAppend = strAppend & " AND Claimability = '" & Trim(Mid(CboClaimability, 1, 1)) & "'"
    End If
    
    If ChkClaimability.Value = 1 Then
        strAppend = strAppend & " AND (Claimability Is null OR Claimability = '')"
    End If
    
    If Len(Mid(CboStatus, 1, 1)) Then
        strAppend = strAppend & " AND MBMStatus = '" & Trim(Mid(CboStatus, 1, 1)) & "'"
    End If
    
    If ChkPolStatus.Value = 1 Then
        strAppend = strAppend & " AND (MBMStatus Is null OR MBMStatus = '')"
    End If

    If Len(Mid(CboCaseStatus, 1, 1)) Then
        strAppend = strAppend & " AND CASStatus = '" & Trim(Mid(CboCaseStatus, 1, 1)) & "'"
    End If
    
    If ChkCaseStatus.Value = 1 Then
        strAppend = strAppend & " AND (CASStatus Is null OR CASStatus = '')"
    End If
    
    If Len(Trim(Mid(CboExpired, 1, 1))) Then
          strAppend = strAppend & "AND MBMExpiredStatus = '" & Trim(Mid(CboExpired, 1, 1)) & "' "
    End If
    
    If ChkExpired.Value = 1 Then
        strAppend = strAppend & " AND (MBMExpiredStatus Is null OR MBMExpiredStatus = '')"
    End If
    
    If Len(Mid(CboPreMode, 1, 1)) Then
        strAppend = strAppend & " AND MBMInstallment = '" & Trim(Mid(CboPreMode, 1, 1)) & "'"
    End If
    
    If ChkPreMode.Value = 1 Then
        strAppend = strAppend & " AND (MBMInstallment Is null OR MBMInstallment = '')"
    End If
    
    If Len(Trim(CboState)) Then
          strAppend = strAppend & "AND STACode= '" & Trim(CboState) & "' "
    End If
    
    If ChkState.Value = 1 Then
        strAppend = strAppend & " AND (STACode Is null OR STACode = '')"
    End If
    
    If Mid(cboRenewalInd, 1, 1) = "P" Then
        strAppend = strAppend & " AND (MBMRenewFlag = 'P')"
    ElseIf Mid(cboRenewalInd, 1, 1) = "N" Then
        strAppend = strAppend & " AND (MBMRenewFlag Is null OR MBMRenewFlag = '' OR MBMRenewFlag = 'N') "
    ElseIf Mid(cboRenewalInd, 1, 1) = "R" Then
        strAppend = strAppend & " AND (MBMRenewFlag = 'R')"
    ElseIf Mid(cboRenewalInd, 1, 1) = "T" Then
        strAppend = strAppend & " AND (MBMRenewFlag = 'T')"
    End If
    
    If ChkRenewal.Value = 1 Then
        strAppend = strAppend & " AND (MBMRenewFlag Is null OR MBMRenewFlag = '')"
    End If
    
    If Len(Mid(cboTakeOver, 1, 1)) Then
        strAppend = strAppend & " AND MBMTakeOver = '" & Trim(Mid(cboTakeOver, 1, 1)) & "'"
    End If
    
    If ChkTakeOver.Value = 1 Then
        strAppend = strAppend & " AND (MBMTakeOver Is null OR MBMTakeOver = '')"
    End If
    
    If Len(Trim(txtAgentCode)) Then
       strAppend = strAppend + " And  MBMAgentCode = '" & ChkStr(txtAgentCode) & "'"
    End If
    
    If ChkAgentCode.Value = 1 Then
        strAppend = strAppend & " AND (MBMAgentCode Is null OR MBMAgentCode = '')"
    End If
    
    If Len(Trim(txtPolNo)) Then
       strAppend = strAppend + " And  MBMPolicyNo = '" & ChkStr(txtPolNo) & "'"
    End If
    
    If ChkPolNo.Value = 1 Then
        strAppend = strAppend & " AND (MBMPolicyNo Is null OR MBMPolicyNo = '')"
    End If
    
    If Len(Trim(TxtPatientName)) Then
       strAppend = strAppend + " And  PatientName = '" & ChkStr(TxtPatientName) & "'"
    End If
    
    If Len(Trim(txtPOLYear1)) Then
        strAppend = strAppend & " AND MBMPolicyYear >= '" & Trim(txtPOLYear1) & "'"
    End If
    
    If Len(Trim(txtPOLYear2)) Then
        strAppend = strAppend & " AND MBMPolicyYear <= '" & Trim(txtPOLYear2) & "'"
    End If
    
    If ChkPolYear.Value = 1 Then
        strAppend = strAppend & " AND (MBMPolicyYear Is null OR MBMPolicyYear = '')"
    End If
    
    If Len(Trim(TxtMCDays1)) Then
        strAppend = strAppend & " AND MCDays >= '" & Trim(TxtMCDays1) & "'"
    End If
    
    If Len(Trim(TxtMCDays2)) Then
        strAppend = strAppend & " AND MCDays <= '" & Trim(TxtMCDays2) & "'"
    End If
    
    If ChkMCDays.Value = 1 Then
        strAppend = strAppend & " AND (MCDays Is null OR MCDays = '')"
    End If
    
    If Len(Mid(cboPayorCode, 1, 2)) Then
        strAppend = strAppend & " AND PAYCode = '" & Trim(Mid(cboPayorCode, 1, 2)) & "'"
    End If
    
    If ChkPayor.Value = 1 Then
        strAppend = strAppend & " AND (PAYCode Is null OR PAYCode = '')"
    End If
           
    If Len(Trim(CboGrpCom)) Then
         GrpCompany = Replace(CboGrpCom, "'", "''")
         strAppend = strAppend + " AND MBMGroupCompany = '" & Trim(GrpCompany) & "'"
    End If
    
    If ChkGrpCom.Value = 1 Then
        strAppend = strAppend & " AND (MBMGroupCompany Is null OR MBMGroupCompany = '')"
    End If
    
    If Len(Trim(cboCLNName)) Then
        Clinics = Replace(cboCLNName, "'", "''")
        strAppend = strAppend + " AND HOSName = '" & Trim(Clinics) & "'"
    End If
    
    If ChkCLN.Value = 1 Then
        strAppend = strAppend & " AND (HOSName Is null OR HOSName = '')"
    End If
    
    If Len(Trim(cboCLNGrp)) Then
        strAppend = strAppend + " And HOSGRPName = '" & Trim(cboCLNGrp) & "'"
    End If
    
    If ChkCLNGrp.Value = 1 Then
        strAppend = strAppend & " AND (HOSGRPName Is null OR HOSGRPName = '')"
    End If
    
    If Len(Trim(TxtBranch)) Then
        strAppend = strAppend + " And BRNCode = '" & Trim(TxtBranch) & "'"
    End If
    
    If ChkBranch.Value = 1 Then
        strAppend = strAppend & " AND (BRNCode Is null OR BRNCode = '')"
    End If
    
    If Len(Trim(TxtDOV1)) > 0 And IsDate(TxtDOV1) = True Then
        strAppend = strAppend + " And DateVisit >= '" & Format(TxtDOV1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOV2)) > 0 And IsDate(TxtDOV2) = True Then
        strAppend = strAppend + " And DateVisit <= '" & Format(TxtDOV2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkDOV.Value = 1 Then
        strAppend = strAppend & " AND (DateVisit Is null OR DateVisit = '')"
    End If
    
    If Len(Trim(TxtDOR1)) > 0 And IsDate(TxtDOR1) = True Then
        strAppend = strAppend + " And RegisterDate >= '" & Format(TxtDOR1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOR2)) > 0 And IsDate(TxtDOR2) = True Then
        strAppend = strAppend + " And RegisterDate <= '" & Format(TxtDOR2, "mm/dd/yyyy") & "'"
    End If
    
    If chkDOR.Value = 1 Then
        strAppend = strAppend & " AND (RegisterDate Is null OR RegisterDate = '')"
    End If
    
    If Len(Trim(txtDOW1)) > 0 And IsDate(txtDOW1) = True Then
        strAppend = strAppend + " And ClaimOpenDate >= '" & Format(txtDOW1, "mm/dd/yyyy") & "'"
    End If

    If Len(Trim(txtDOW2)) > 0 And IsDate(txtDOW2) = True Then
        strAppend = strAppend + " And ClaimOpenDate <= '" & Format(txtDOW2, "mm/dd/yyyy") & "'"
    End If

    If ChkDOW.Value = 1 Then
        strAppend = strAppend & " AND (ClaimOpenDate Is null OR ClaimOpenDate = '')"
    End If
    
    If Len(Trim(txtPreDiag1)) Then
        strAppend = strAppend + " And CASTreatReason1 >= '" & Trim(txtPreDiag1) & "'"
    End If
    
    If Len(Trim(txtPreDiag2)) Then
        strAppend = strAppend + " And CASTreatReason1 <= '" & Trim(txtPreDiag2) & "'"
    End If
    
    If ChkFirst.Value = 1 Then
        strAppend = strAppend & " AND (CASTreatReason1 Is null OR CASTreatReason1 = '')"
    End If
    
    If Len(Trim(TxtPlan1)) Then
        strAppend = strAppend & " AND PLNCode >= '" & Trim(TxtPlan1) & "'"
    End If
    
    If Len(Trim(TxtPlan2)) Then
        strAppend = strAppend & " AND PLNCode <= '" & Trim(TxtPlan2) & "'"
    End If
    
    If ChkPlan.Value = 1 Then
        strAppend = strAppend & " AND (PLNCode Is null OR PLNCode = '')"
    End If
    
    If Len(Trim(TxtPayEffDate1)) > 0 And IsDate(TxtPayEffDate1) = True Then
        strAppend = strAppend + " And MBMPayorEffDate >= '" & Format(TxtPayEffDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtPayEffDate2)) > 0 And IsDate(TxtPayEffDate2) = True Then
        strAppend = strAppend + " And MBMPayorEffDate <= '" & Format(TxtPayEffDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkEffDate.Value = 1 Then
        strAppend = strAppend & " AND (MBMPayorEffDate Is null OR MBMPayorEffDate = '')"
    End If
    
    If Len(Trim(TxtPayExpDate1)) And IsDate(TxtPayExpDate1) = True Then
        strAppend = strAppend & " AND MBMPayorExpDate >= '" & Format(TxtPayExpDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtPayExpDate2)) And IsDate(TxtPayExpDate2) = True Then
        strAppend = strAppend & " AND MBMPayorExpDate <= '" & Format(TxtPayExpDate2, "mm/dd/yyyy") & "'"
    End If
    
    If chkExpDate.Value = 1 Then
        strAppend = strAppend & " AND (MBMPayorExpDate Is null OR MBMPayorExpDate = '')"
    End If
    
    If Len(Trim(txtDateJoin1)) And IsDate(txtDateJoin1) = True Then
        strAppend = strAppend & " AND MBMFirstJoinedDate >= '" & Format(txtDateJoin1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDateJoin2)) And IsDate(txtDateJoin2) = True Then
        strAppend = strAppend & " AND MBMFirstJoinedDate <= '" & Format(txtDateJoin2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkJoinDate.Value = 1 Then
        strAppend = strAppend & " AND (MBMFirstJoinedDate Is null OR MBMFirstJoinedDate = '')"
    End If
    
    If Len(Trim(txtDOB1)) And IsDate(txtDOB1) = True Then
        strAppend = strAppend & " AND DOB >= '" & Format(txtDOB1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOB2)) And IsDate(txtDOB2) = True Then
        strAppend = strAppend & " AND DOB <= '" & Format(txtDOB2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkDOB.Value = 1 Then
        strAppend = strAppend & " AND (DOB Is null OR DOB = '')"
    End If
    
    If Len(Trim(TxtMBMNo1)) Then
       strAppend = strAppend + " AND MBMNumber >= '" & RemStr(Trim(TxtMBMNo1)) & "'"
    End If
    
    If Len(Trim(TxtMBMNo2)) Then
       strAppend = strAppend + " AND MBMNumber <= '" & RemStr(Trim(TxtMBMNo2)) & "'"
    End If
    
    If Len(Trim(TxtCaseNo1)) Then
       strAppend = strAppend + " AND CASNumber >= '" & RemStr(Trim(TxtCaseNo1)) & "'"
    End If
    
    If Len(Trim(TxtCaseNo2)) Then
       strAppend = strAppend + " AND CASNumber <= '" & RemStr(Trim(TxtCaseNo2)) & "'"
    End If
        
    If Len(Trim(TxtBordxNo1)) Then
        strAppend = strAppend & " AND BordxRefNo >= '" & Trim(TxtBordxNo1) & "'"
    End If
    
    If Len(Trim(TxtBordxNo2)) Then
        strAppend = strAppend & " AND BordxRefNo <= '" & Trim(TxtBordxNo2) & "'"
    End If
    
    If ChkBordx.Value = 1 Then
        strAppend = strAppend & " AND (BordxRefNo Is null OR BordxRefNo = '')"
    End If
    
    If Len(Trim(TxtBordxDate1)) And IsDate(TxtBordxDate1) = True Then
        strAppend = strAppend & " AND BordxDate >= '" & Format(TxtBordxDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtBordxDate2)) And IsDate(TxtBordxDate2) = True Then
        strAppend = strAppend & " AND BordxDate <= '" & Format(TxtBordxDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkBordxDate.Value = 1 Then
        strAppend = strAppend & " AND (BordxDate Is null OR BordxDate = '')"
    End If
    
    If Len(Trim(TxtICNo1)) Then
       strAppend = strAppend + " AND ICNo >= '" & RemStr(Trim(TxtICNo1)) & "'"
    End If
    
    If Len(Trim(TxtICNo2)) Then
       strAppend = strAppend + " AND ICNo <= '" & RemStr(Trim(TxtICNo2)) & "'"
    End If
    
    If ChkICNo.Value = 1 Then
        strAppend = strAppend & " AND (ICNo Is null OR ICNo = '')"
    End If
    
    If Len(Trim(TxtMBMIc1)) Then
       strAppend = strAppend + " AND MBMIcBcPp >= '" & RemStr(Trim(TxtMBMIc1)) & "'"
    End If
    
    If Len(Trim(TxtMBMIc2)) Then
       strAppend = strAppend + " AND MBMIcBcPp <= '" & RemStr(Trim(TxtMBMIc2)) & "'"
    End If
    
    If ChkMBMIC.Value = 1 Then
        strAppend = strAppend & " AND (MBMIcBcPp Is null OR MBMIcBcPp = '')"
    End If
    
    If Len(txtActualAmt1) Then
        strAppend = strAppend + " AND CLMTotalActual >= " & Trim(txtActualAmt1) & ""
    End If
        
    If Len(txtActualAmt2) Then
        strAppend = strAppend + " AND CLMTotalActual <= " & Trim(txtActualAmt2) & ""
    End If
    
    If Len(txtAppAmt1) Then
        strAppend = strAppend + " AND CLMTotalApproved >= " & Trim(txtAppAmt1) & ""
    End If
        
    If Len(txtAppAmt2) Then
        strAppend = strAppend + " AND CLMTotalApproved <= " & Trim(txtAppAmt2) & ""
    End If
    
    If Len(Trim(TxtInvoice1)) Then
        strAppend = strAppend + " AND InvoiceNo >= '" & Trim(TxtInvoice1) & "'"
    End If
    
    If Len(Trim(TxtInvoice2)) Then
        strAppend = strAppend + " AND InvoiceNo <= '" & Trim(TxtInvoice2) & "'"
    End If
    
    If ChkInvoice.Value = 1 Then
        strAppend = strAppend + " AND (InvoiceNo Is null OR InvoiceNo = '')"
    End If
    
    If Len(Trim(TxtSubmission1)) Then
        strAppend = strAppend + " AND BillNo >= '" & Trim(TxtSubmission1) & "'"
    End If
    
    If Len(Trim(TxtSubmission2)) Then
        strAppend = strAppend + " AND BillNo <= '" & Trim(TxtSubmission2) & "'"
    End If
    
    If ChkSubmission.Value = 1 Then
        strAppend = strAppend + " AND (BillNo Is null OR BillNo = '')"
    End If
    
    
    If Option1.Value = True And Option4.Value = True Then
        Call ActualClaim(strAppend)
    ElseIf Option1.Value = True And Option5.Value = True Then
        Call ActualInsPlan(strAppend)
    ElseIf Option1.Value = True And Option6.Value = True Then
        Call ActualInsAge(strAppend)
    ElseIf Option1.Value = True And Option7.Value = True Then
        Call ActualInsSex(strAppend)
    ElseIf Option1.Value = True And Option8.Value = True Then
        Call ActualAgeband(strAppend)
    ElseIf Option1.Value = True And Option9.Value = True Then
        Call ActualSex(strAppend)
    ElseIf Option1.Value = True And Option10.Value = True Then
        Call ActualAgebandSex(strAppend)
    ElseIf Option1.Value = True And Option11.Value = True Then
        Call ActualFinalDiag(strAppend)
    ElseIf Option1.Value = True And Option12.Value = True Then
        Call ActualClinic(strAppend)
    ElseIf Option1.Value = True And Option13.Value = True Then
        Call ActualPayor(strAppend)
    ElseIf Option1.Value = True And Option14.Value = True Then
        Call ActualPayorInsPlan(strAppend)
    ElseIf Option1.Value = True And Option15.Value = True Then
        Call ActualGrpCompany(strAppend)
    ElseIf Option1.Value = True And Option16.Value = True Then
        Call ActualInsRelation(strAppend)
    ElseIf Option1.Value = True And Option17.Value = True Then
        Call ActualIns(strAppend)
    ElseIf Option1.Value = True And Option18.Value = True Then
        Call ActualBordxNo(strAppend)
    ElseIf Option1.Value = True And Option19.Value = True Then
        Call ActualPolYear(strAppend)
    ElseIf Option1.Value = True And Option20.Value = True Then
        Call ActualTakeover(strAppend)
    ElseIf Option2.Value = True And Option4.Value = True Then
        Call ApprovedClaim(strAppend)
    ElseIf Option2.Value = True And Option5.Value = True Then
        Call ApprovedInsPlan(strAppend)
    ElseIf Option2.Value = True And Option6.Value = True Then
        Call ApprovedInsAge(strAppend)
    ElseIf Option2.Value = True And Option7.Value = True Then
        Call ApprovedInsSex(strAppend)
    ElseIf Option2.Value = True And Option8.Value = True Then
        Call ApprovedAgeband(strAppend)
    ElseIf Option2.Value = True And Option9.Value = True Then
        Call ApprovedSex(strAppend)
    ElseIf Option2.Value = True And Option10.Value = True Then
        Call ApprovedAgebandSex(strAppend)
    ElseIf Option2.Value = True And Option11.Value = True Then
        Call ApprovedFinalDiag(strAppend)
    ElseIf Option2.Value = True And Option12.Value = True Then
        Call ApprovedClinic(strAppend)
    ElseIf Option2.Value = True And Option13.Value = True Then
        Call ApprovedPayor(strAppend)
    ElseIf Option2.Value = True And Option14.Value = True Then
        Call ApprovedPayorInsPlan(strAppend)
    ElseIf Option2.Value = True And Option15.Value = True Then
        Call ApprovedGrpCompany(strAppend)
    ElseIf Option2.Value = True And Option16.Value = True Then
        Call ApprovedInsRelation(strAppend)
    ElseIf Option2.Value = True And Option17.Value = True Then
        Call ApprovedIns(strAppend)
    ElseIf Option2.Value = True And Option18.Value = True Then
        Call ApprovedBordxNo(strAppend)
    ElseIf Option2.Value = True And Option19.Value = True Then
        Call ApprovedPolYear(strAppend)
    ElseIf Option2.Value = True And Option20.Value = True Then
        Call ApprovedTakeover(strAppend)
    ElseIf Option3.Value = True Then
        Call CostDetailActual(strAppend)
        Call CostDetailApproved(strAppend)
        Call CostDetailCLN(strAppend)
        Call CostDetailMBM(strAppend)
    End If
End Function

Private Sub ActualClaim(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    '====================
    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    '==================
    Total = 0
    Current = 0
        
    ListView2.ListItems.Clear
            
    'sql = " SELECT CASNumber, MBMNumber, PatientName, DateVisit, " & _
          " CLMGPFeeBasicMed, CLMGPFeeMedInj, CLMGPFeeLTMed, CLMSPFeeBasicMed, " & _
          " CLMSPFeeMedInj, CLMSPFeeLTMed, CLMXRay, CLMUltraSound, CLMECG, CLMDressing, " & _
          " CLMNebulizer, CLMID, CLMTSSTO, CLMMinorSurgery, CLMPOthers, " & _
          " CLMBloodDrawing, CLMBloodGlucose, CLMUrine, CLMPapSmear, CLMLOthers, " & _
          " CLMDTPPolio, CLMDTPPolioHIB, CLMHepatitisB, CLMMMR, CLMIOthers, " & _
          " CLMDental, CLMMaternity, CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
        
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "5"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            'lblCurrent = rstCount!TotalRecord
            With rst2
                If Len(rst2!CasNumber) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!CasNumber)
                End If
                                              
                If Len(rst2!MBMNumber) Then
                   Record.SubItems(1) = rst2!MBMNumber
                End If
                
                If Len(rst2!PatientName) Then
                   Record.SubItems(2) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!DateVisit) Then
                    Record.SubItems(3) = Format(rst2!DateVisit, "dd/mm/yyyy")
                End If
                                
                If Len(rst2!CLMGPFeeBasicMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(7) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(7) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(8) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(8) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(9) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(9) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(10) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(10) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(11) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(11) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(12) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(12) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(13) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(13) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(14) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(14) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(15) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(15) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(16) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(16) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(17) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(17) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(18) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(18) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(19) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(19) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(20) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(20) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(21) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(21) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(22) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(22) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(23) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(23) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(24) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(24) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(25) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(25) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(26) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(26) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(27) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(27) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(28) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(28) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(29) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(29) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(30) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(30) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(31) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualInsPlan(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT INSCode, PLNCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by INSCode, PLNCode Order by INSCode, PLNCode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "12"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , !INSCode)
                End If
                                              
                If Len(!PLNCode) Then
                    Record.SubItems(1) = Trim(!PLNCode)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(2) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(2) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(3) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(7) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(7) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(8) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(8) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(9) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(9) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(10) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(10) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(11) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(11) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(12) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(12) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(13) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(13) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(14) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(14) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(15) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(15) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(16) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(17) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(18) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(18) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(19) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(19) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(20) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(20) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(21) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(22) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(23) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(23) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(24) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(24) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(25) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(25) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(26) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(26) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(27) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(27) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(28) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(28) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(29) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
   
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualInsAge(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT INSCode, AGEBand, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by INSCode, AGEBand Order by INSCode, Ageband "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "11"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , !INSCode)
                End If
                                              
                If Len(!AgeBand) Then
                    Record.SubItems(1) = Trim(!AgeBand)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(2) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(2) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(3) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(7) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(7) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(8) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(8) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(9) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(9) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(10) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(10) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(11) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(11) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(12) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(12) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(13) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(13) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(14) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(14) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(15) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(15) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(16) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(17) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(18) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(18) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(19) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(19) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(20) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(20) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(21) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(22) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(23) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(23) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(24) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(24) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(25) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(25) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(26) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(26) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(27) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(27) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(28) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(28) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(29) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualAgeband(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT AGEBand, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by AGEBand "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "1"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
     Do
        Do Until rst2.EOF
            
            With rst2
                If Len(!AgeBand) Then
                    Set Record = ListView2.ListItems.Add(, , !AgeBand)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualInsSex(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT INSCode, Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by INSCode, Sex Order by INSCode, Sex "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "15"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , !INSCode)
                End If
                                              
                If Len(!Sex) Then
                    Record.SubItems(1) = Trim(!Sex)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(2) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(2) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(3) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(7) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(7) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(8) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(8) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(9) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(9) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(10) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(10) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(11) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(11) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(12) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(12) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(13) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(13) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(14) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(14) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(15) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(15) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(16) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(17) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(18) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(18) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(19) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(19) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(20) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(20) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(21) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(22) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(23) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(23) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(24) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(24) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(25) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(25) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(26) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(26) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(27) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(27) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(28) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(28) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(29) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualSex(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by Sex Order by Sex "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "18"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!Sex) Then
                    Set Record = ListView2.ListItems.Add(, , !Sex)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualAgebandSex(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT AGEBand, Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by AGEBand, Sex "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
        
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "3"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!AgeBand) Then
                    Set Record = ListView2.ListItems.Add(, , !AgeBand)
                End If
                                              
                If Len(!Sex) Then
                    Record.SubItems(1) = Trim(!Sex)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(2) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(2) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(3) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(7) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(7) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(8) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(8) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(9) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(9) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(10) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(10) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(11) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(11) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(12) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(12) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(13) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(13) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(14) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(14) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(15) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(15) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(16) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(17) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(18) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(18) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(19) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(19) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(20) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(20) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(21) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(22) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(23) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(23) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(24) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(24) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(25) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(25) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(26) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(26) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(27) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(27) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(28) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(28) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(29) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualFinalDiag(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT CASTreatReason1, ReasonDesc, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by CASTreatReason1, ReasonDesc Order by CASTreatReason1, ReasonDesc "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "8"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!CASTreatReason1) Then
                    Set Record = ListView2.ListItems.Add(, , !CASTreatReason1)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                              
                If Len(!ReasonDesc) Then
                    Record.SubItems(1) = Trim(!ReasonDesc)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(2) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(2) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(3) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(7) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(7) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(8) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(8) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(9) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(9) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(10) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(10) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(11) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(11) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(12) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(12) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(13) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(13) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(14) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(14) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(15) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(15) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(16) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(17) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(18) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(18) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(19) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(19) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(20) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(20) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(21) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(22) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(23) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(23) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(24) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(24) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(25) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(25) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(26) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(26) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(27) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(27) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(28) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(28) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(29) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualClinic(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT ClinicCode, CLNName, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by ClinicCode, CLNName Order by ClinicCode, CLNName "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "7"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!ClinicCode) Then
                    Set Record = ListView2.ListItems.Add(, , !ClinicCode)
                End If
                                              
                If Len(!CLNName) Then
                    Record.SubItems(1) = Trim(!CLNName)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(2) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(2) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(3) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(7) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(7) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(8) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(8) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(9) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(9) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(10) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(10) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(11) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(11) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(12) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(12) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(13) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(13) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(14) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(14) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(15) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(15) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(16) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(17) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(18) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(18) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(19) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(19) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(20) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(20) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(21) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(22) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(23) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(23) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(24) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(24) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(25) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(25) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(26) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(26) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(27) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(27) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(28) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(28) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(29) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualClaimNature(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    Others = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT ClaimNatureCode, SUM(CONVERT(Money,CLMBoardAmt)) AS CLMBoardAmt, SUM(CONVERT(Money,CLMICUAmt)) AS CLMICUAmt, " & _
          " SUM(CONVERT(Money,CLMServices)) AS CLMServices, SUM(CONVERT(Money,CLMTheatre)) AS CLMTheatre, SUM(CONVERT(Money,CLMPreHSAmt)) AS CLMPreHSAmt, " & _
          " SUM(CONVERT(Money,CLMSurgicalAmt)) AS CLMSurgicalAmt, SUM(CONVERT(Money,CLMAna)) AS CLMAna, SUM(CONVERT(Money,CLMNSPreHos)) AS CLMNSPreHos, " & _
          " SUM(CONVERT(Money,CLMPhysician)) AS CLMPhysician, SUM(CONVERT(Money,CLMPost)) AS CLMPost, SUM(CONVERT(Money,CLMAmbulance)) AS CLMAmbulance, " & _
          " SUM(CONVERT(Money,CLMOutPatient)) AS CLMOutPatient, SUM(CONVERT(Money,CLMMonthly)) AS CLMMonthly, SUM(CONVERT(Money,CLMOrgan)) AS CLMOrgan, " & _
          " SUM(CONVERT(Money,CLMMedical)) AS CLMMedical, SUM(CONVERT(Money,CLMPhone)) AS CLMPhone, SUM(CONVERT(Money,CLMGuardianAmt)) AS CLMGuardianAmt, " & _
          " SUM(CONVERT(Money,CLMAdmin)) AS CLMAdmin, SUM(CONVERT(Money,CLMTax)) AS CLMTax, SUM(CONVERT(Money,CLMOthers2)) AS CLMOthers2, SUM(CONVERT(Money,CLMCashAmt)) AS CLMCashAmt " & _
          " From VIEW_CLMReport_Act Where 0 = 0 "
    
    'sql2 = " Group by ClaimNatureCode Order by ClaimNatureCode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Act where 0 = 0 "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "6"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!ClaimNatureCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!ClaimNatureCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(5) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREDiag = 0
                    Else
                        PREDiag = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(7) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(8) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        PreSpec = 0
                    Else
                        PreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(9) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(10) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(11) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(12) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                'If Len(rst2!CLMBoardAmt) Then
                   'Record.SubItems(13) = Format(rst2!CLMBoardAmt, "#,##0.00")
                    'If rst2!CLMAmbulance = "-" Then
                        'Physio = 0
                    'Else
                        'Physio = rst2!CLMAmbulance
                    'End If
                'End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(14) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        KidCan = 0
                    Else
                        KidCan = rst2!CLMMonthly
                    End If
                End If
                
                'If Len(rst2!CLMBoardAmt) Then
                   'Record.SubItems(15) = Format(rst2!CLMBoardAmt, "#,##0.00")
                    'if rst2!CLMMonthly = "-" then
                        'Kid = 0
                    'Else
                        'Kid = rst2!CLMMonthly
                    'End If
                'End If
                
                'If Len(rst2!CLMICUAmt) Then
                   'Record.SubItems(16) = Format(rst2!CLMICUAmt, "#,##0.00")
                    'if rst2!CLMMonthly = "-" then
                        'Can = 0
                    'else
                        'Can = rst2!CLMMonthly
                    'End If
                'End If
                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                
                If Len(rst2!CLMCashAmt) Then
                   Record.SubItems(18) = Format(rst2!CLMCashAmt, "#,##0.00")
                    If rst2!CLMCashAmt = "-" Then
                        Govt = 0
                    Else
                        Govt = rst2!CLMCashAmt
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers2
                    End If
                End If
                
                'If Len(rst2!CLMOthers2) Then
                   'Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    'If rst2!CLMOthers2 = "-" Then
                        'Maternity = 0
                    'Else
                        'Maternity =
                    'End If
                'End If
                
                'If Len(rst2!CLMOthers2) Then
                   'Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    'If rst2!CLMOthers2 = "-" Then
                        'Others2 = 0
                    'Else
                        'Others2 =
                    'End If
                'End If
                
                Total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(PREDiag) + CDbl(Surg) + CDbl(Anaest) _
                        + CDbl(PreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) + CDbl(Ambul) + CDbl(Physio) _
                        + CDbl(KidCan) + CDbl(Kid) + CDbl(Can) + CDbl(Organ) + CDbl(Govt) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Maternity) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
   
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualPayor(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT PAYCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by PAYCode Order by PAYCode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "16"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualPayorInsPlan(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT PAYCode, INSCode, PLNCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by PAYCode, INSCode, PLNCode Order by PAYCode, INSCode, PLNCode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "17"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                End If
                                              
                If Len(!INSCode) Then
                    Record.SubItems(1) = Trim(!INSCode)
                End If
                
                If Len(!PLNCode) Then
                    Record.SubItems(2) = Trim(!PLNCode)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(3) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(6) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(7) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(7) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(8) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(8) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(9) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(9) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(10) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(10) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(11) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(11) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(12) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(12) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(13) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(13) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(14) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(14) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(15) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(15) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(16) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(16) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(17) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(18) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(18) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(19) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(19) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(20) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(20) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(21) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(21) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(22) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(23) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(23) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(24) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(24) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(25) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(25) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(26) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(26) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(27) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(27) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(28) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(28) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(29) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(29) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(30) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
 
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualInsRelation(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT INSCode, Relation, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by INSCode, Relation Order by INSCode, Relation "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "13"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , !INSCode)
                End If
                                              
                If Len(!Relation) Then
                    Record.SubItems(1) = Trim(!Relation)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(2) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(2) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(3) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(7) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(7) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(8) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(8) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(9) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(9) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(10) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(10) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(11) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(11) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(12) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(12) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(13) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(13) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(14) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(14) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(15) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(15) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(16) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(17) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(18) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(18) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(19) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(19) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(20) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(20) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(21) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(22) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(23) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(23) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(24) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(24) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(25) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(25) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(26) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(26) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(27) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(27) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(28) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(28) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(29) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualIns(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT INSCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by INSCode Order by INSCode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "10"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , !INSCode)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
   
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualBordxNo(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT BordxRefNo, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by BordxRefNo Order by BordxRefNo "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "

    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "4"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!Bordxrefno) Then
                    Set Record = ListView2.ListItems.Add(, , !Bordxrefno)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
   
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualPolYear(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT MBMPolicyYear, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by MBMPolicyYear Order by MBMPolicyYear "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "14"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!MBMPolicyYear) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMPolicyYear)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
   
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualTakeover(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT MBMTakeOver, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by MBMTakeOver Order by MBMTakeOver "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "19"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!MBMTakeOver) Then
                    Set Record = ListView2.ListItems.Add(, , Trim(!MBMTakeOver))
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Sub DoInitialSettings()
    Fgrid.Row = 0
    Fgrid.Col = 0
    Fgrid = "Benefits"
    Fgrid.Col = 1
    Fgrid = "Actual"
    Fgrid.Col = 2
    Fgrid = "Approved"
    Fgrid.Col = 3
    Fgrid = "Pay2Clinic"
    Fgrid.Col = 4
    Fgrid = "Pay2/fr Mem"

    Fgrid.Col = 0
    
    Fgrid.ColAlignment(0) = 1
    Fgrid.ColAlignment(1) = 6
    Fgrid.ColAlignment(2) = 6
    Fgrid.ColAlignment(3) = 6
    Fgrid.ColAlignment(4) = 6
    Fgrid.ColWidth(0) = 3000
    Fgrid.ColWidth(1) = 1500
    Fgrid.ColWidth(2) = 1500
    Fgrid.ColWidth(3) = 1500
    Fgrid.ColWidth(4) = 1500
    Fgrid.Row = 1
    Fgrid = "101 - GP Fee & Basic Med"
    
    Fgrid.Row = 2
    Fgrid = "102 - GP Fee & Med & Inj"
    
    Fgrid.Row = 3
    Fgrid = "103 - GP Fee & LT Med"
    
    Fgrid.Row = 4
    Fgrid = "104 - SP Fee & Basic Med"
    
    Fgrid.Row = 5
    Fgrid = "105 - SP Fee & Med & Inj"
    
    Fgrid.Row = 6
    Fgrid = "106 - SP Fee & LT Med"
    
    Fgrid.Row = 7
    Fgrid = "201 - X-Ray"
    
    Fgrid.Row = 8
    Fgrid = "202 - Ultra Sound"
    
    Fgrid.Row = 9
    Fgrid = "203 - ECG"
    
    Fgrid.Row = 10
    Fgrid = "204 - Dressing"
    
    Fgrid.Row = 11
    Fgrid = "205 - Nebulizer"
    
    Fgrid.Row = 12
    Fgrid = "206 - I&D/T&S"
    
    Fgrid.Row = 13
    Fgrid = "207 - STO"
    
    Fgrid.Row = 14
    Fgrid = "208 - Minor Surgery"
    
    Fgrid.Row = 15
    Fgrid = "209 - Procedure Others"
    
    Fgrid.Row = 16
    Fgrid = "301 - Blooding Drawing"
    
    Fgrid.Row = 17
    Fgrid = "302 - Blood Glucose"
    
    Fgrid.Row = 18
    Fgrid = "303 - Urine"
    
    Fgrid.Row = 19
    Fgrid = "304 - Pap Smear"
    
    Fgrid.Row = 20
    Fgrid = "305 - Lab Others"
    
    Fgrid.Row = 21
    Fgrid = "401 - DTP & Polio"
    
    Fgrid.Row = 22
    Fgrid = "402 - DTP & Polio & HIB"
    
    Fgrid.Row = 23
    Fgrid = "403 - Hepatitis B"
    
    Fgrid.Row = 24
    Fgrid = "404 - MMR"
    
    Fgrid.Row = 25
    Fgrid = "405 - Immunization Others"
    
    Fgrid.Row = 26
    Fgrid = "501 - Dental"
    
    Fgrid.Row = 27
    Fgrid = "502 - Maternity"
    
    Fgrid.Row = 28
    Fgrid = "503 - Preventive Checkup"
    
    Fgrid.Row = 29
    Fgrid.CellFontBold = True
    Fgrid.CellFontSize = 8
    Fgrid = "Total"
 
End Sub

Private Sub CostDetailActual(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    Total = 0
    
    'sql = " SELECT SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
        
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual WHERE SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "37"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            With rst2
                Fgrid.Col = 1
                Fgrid.Row = 1
                Fgrid.ColAlignment(1) = 6
                
                If Len(!CLMGPFeeBasicMed) Then
                    Fgrid = !CLMGPFeeBasicMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 2
                If Len(!CLMGPFeeMedInj) Then
                    Fgrid = !CLMGPFeeMedInj
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 3
                If Len(!CLMGPFeeLTMed) Then
                    Fgrid = !CLMGPFeeLTMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 4
                If Len(!CLMSPFeeBasicMed) Then
                    Fgrid = !CLMSPFeeBasicMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 5
                If Len(!CLMSPFeeMedInj) Then
                    Fgrid = !CLMSPFeeMedInj
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 6
                If Len(!CLMSPFeeLTMed) Then
                    Fgrid = !CLMSPFeeLTMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 7
                If Len(!CLMXRay) Then
                    Fgrid = !CLMXRay
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 8
                If Len(!CLMUltraSound) Then
                    Fgrid = !CLMUltraSound
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 9
                If Len(!CLMECG) Then
                    Fgrid = !CLMECG
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 10
                If Len(!CLMDressing) Then
                    Fgrid = !CLMDressing
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 11
                If Len(!CLMNebulizer) Then
                    Fgrid = !CLMNebulizer
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 12
                If Len(!CLMID) Then
                    Fgrid = !CLMID
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 13
                If Len(!CLMTSSTO) Then
                    Fgrid = !CLMTSSTO
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 14
                If Len(!CLMMinorSurgery) Then
                    Fgrid = !CLMMinorSurgery
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 15
                If Len(!CLMPOthers) Then
                    Fgrid = !CLMPOthers
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 16
                If Len(!CLMBloodDrawing) Then
                    Fgrid = !CLMBloodDrawing
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 17
                If Len(!CLMBloodGlucose) Then
                    Fgrid = !CLMBloodGlucose
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 18
                If Len(!CLMUrine) Then
                    Fgrid = !CLMUrine
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 19
                If Len(!CLMPapSmear) Then
                    Fgrid = !CLMPapSmear
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 20
                If Len(!CLMLOthers) Then
                    Fgrid = !CLMLOthers
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 21
                If Len(!CLMDTPPolio) Then
                    Fgrid = !CLMDTPPolio
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 22
                If Len(!CLMDTPPolioHIB) Then
                    Fgrid = !CLMDTPPolioHIB
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 23
                If Len(!CLMHepatitisB) Then
                    Fgrid = !CLMHepatitisB
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 24
                If Len(!CLMMMR) Then
                    Fgrid = !CLMMMR
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 25
                If Len(!CLMIOthers) Then
                    Fgrid = !CLMIOthers
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 26
                If Len(!CLMDental) Then
                    Fgrid = !CLMDental
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 27
                If Len(!CLMMaternity) Then
                    Fgrid = !CLMMaternity
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 28
                If Len(!CLMPRVCheckup) Then
                    Fgrid = !CLMPRVCheckup
                Else
                    Fgrid = 0
                End If
                    
                                
                Total = CDbl(IIf(IsNumeric(!CLMGPFeeBasicMed) = True, !CLMGPFeeBasicMed, 0)) + CDbl(IIf(IsNumeric(!CLMGPFeeMedInj) = True, !CLMGPFeeMedInj, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMGPFeeLTMed) = True, !CLMGPFeeLTMed, 0)) + CDbl(IIf(IsNumeric(!CLMSPFeeBasicMed) = True, !CLMSPFeeBasicMed, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMSPFeeMedInj) = True, !CLMSPFeeMedInj, 0)) + CDbl(IIf(IsNumeric(!CLMSPFeeLTMed) = True, !CLMSPFeeLTMed, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMXRay) = True, !CLMXRay, 0)) + CDbl(IIf(IsNumeric(!CLMUltraSound) = True, !CLMUltraSound, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMECG) = True, !CLMECG, 0)) + CDbl(IIf(IsNumeric(!CLMDressing) = True, !CLMDressing, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMNebulizer) = True, !CLMNebulizer, 0)) + CDbl(IIf(IsNumeric(!CLMID) = True, !CLMID, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMTSSTO) = True, !CLMTSSTO, 0)) + CDbl(IIf(IsNumeric(!CLMMinorSurgery) = True, !CLMMinorSurgery, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPOthers) = True, !CLMPOthers, 0)) + CDbl(IIf(IsNumeric(!CLMBloodDrawing) = True, !CLMBloodDrawing, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMBloodGlucose) = True, !CLMBloodGlucose, 0)) + CDbl(IIf(IsNumeric(!CLMUrine) = True, !CLMUrine, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPapSmear) = True, !CLMPapSmear, 0)) + CDbl(IIf(IsNumeric(!CLMLOthers) = True, !CLMLOthers, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMDTPPolio) = True, !CLMDTPPolio, 0)) + CDbl(IIf(IsNumeric(!CLMDTPPolioHIB) = True, !CLMDTPPolioHIB, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMHepatitisB) = True, !CLMHepatitisB, 0)) + CDbl(IIf(IsNumeric(!CLMMMR) = True, !CLMMMR, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMIOthers) = True, !CLMIOthers, 0)) + CDbl(IIf(IsNumeric(!CLMDental) = True, !CLMDental, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMaternity) = True, !CLMMaternity, 0)) + CDbl(IIf(IsNumeric(!CLMPRVCheckup) = True, !CLMPRVCheckup, 0))
                        
                
                Fgrid.Row = 29
                Fgrid = Format(Total, "#,##0.00")
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
  
End Sub

Private Sub CostDetailApproved(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    Total = 0
    
    'sql = " SELECT SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
        
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved WHERE SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "38"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
     
        Do Until rst2.EOF
            With rst2
                Fgrid.Col = 2
                Fgrid.Row = 1
                Fgrid.ColAlignment(1) = 6
                
                If Len(!CLMGPFeeBasicMed) Then
                    Fgrid = !CLMGPFeeBasicMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 2
                If Len(!CLMGPFeeMedInj) Then
                    Fgrid = !CLMGPFeeMedInj
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 3
                If Len(!CLMGPFeeLTMed) Then
                    Fgrid = !CLMGPFeeLTMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 4
                If Len(!CLMSPFeeBasicMed) Then
                    Fgrid = !CLMSPFeeBasicMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 5
                If Len(!CLMSPFeeMedInj) Then
                    Fgrid = !CLMSPFeeMedInj
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 6
                If Len(!CLMSPFeeLTMed) Then
                    Fgrid = !CLMSPFeeLTMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 7
                If Len(!CLMXRay) Then
                    Fgrid = !CLMXRay
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 8
                If Len(!CLMUltraSound) Then
                    Fgrid = !CLMUltraSound
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 9
                If Len(!CLMECG) Then
                    Fgrid = !CLMECG
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 10
                If Len(!CLMDressing) Then
                    Fgrid = !CLMDressing
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 11
                If Len(!CLMNebulizer) Then
                    Fgrid = !CLMNebulizer
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 12
                If Len(!CLMID) Then
                    Fgrid = !CLMID
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 13
                If Len(!CLMTSSTO) Then
                    Fgrid = !CLMTSSTO
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 14
                If Len(!CLMMinorSurgery) Then
                    Fgrid = !CLMMinorSurgery
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 15
                If Len(!CLMPOthers) Then
                    Fgrid = !CLMPOthers
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 16
                If Len(!CLMBloodDrawing) Then
                    Fgrid = !CLMBloodDrawing
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 17
                If Len(!CLMBloodGlucose) Then
                    Fgrid = !CLMBloodGlucose
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 18
                If Len(!CLMUrine) Then
                    Fgrid = !CLMUrine
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 19
                If Len(!CLMPapSmear) Then
                    Fgrid = !CLMPapSmear
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 20
                If Len(!CLMLOthers) Then
                    Fgrid = !CLMLOthers
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 21
                If Len(!CLMDTPPolio) Then
                    Fgrid = !CLMDTPPolio
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 22
                If Len(!CLMDTPPolioHIB) Then
                    Fgrid = !CLMDTPPolioHIB
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 23
                If Len(!CLMHepatitisB) Then
                    Fgrid = !CLMHepatitisB
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 24
                If Len(!CLMMMR) Then
                    Fgrid = !CLMMMR
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 25
                If Len(!CLMIOthers) Then
                    Fgrid = !CLMIOthers
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 26
                If Len(!CLMDental) Then
                    Fgrid = !CLMDental
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 27
                If Len(!CLMMaternity) Then
                    Fgrid = !CLMMaternity
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 28
                If Len(!CLMPRVCheckup) Then
                    Fgrid = !CLMPRVCheckup
                Else
                    Fgrid = 0
                End If
                    
                                
                Total = CDbl(IIf(IsNumeric(!CLMGPFeeBasicMed) = True, !CLMGPFeeBasicMed, 0)) + CDbl(IIf(IsNumeric(!CLMGPFeeMedInj) = True, !CLMGPFeeMedInj, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMGPFeeLTMed) = True, !CLMGPFeeLTMed, 0)) + CDbl(IIf(IsNumeric(!CLMSPFeeBasicMed) = True, !CLMSPFeeBasicMed, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMSPFeeMedInj) = True, !CLMSPFeeMedInj, 0)) + CDbl(IIf(IsNumeric(!CLMSPFeeLTMed) = True, !CLMSPFeeLTMed, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMXRay) = True, !CLMXRay, 0)) + CDbl(IIf(IsNumeric(!CLMUltraSound) = True, !CLMUltraSound, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMECG) = True, !CLMECG, 0)) + CDbl(IIf(IsNumeric(!CLMDressing) = True, !CLMDressing, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMNebulizer) = True, !CLMNebulizer, 0)) + CDbl(IIf(IsNumeric(!CLMID) = True, !CLMID, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMTSSTO) = True, !CLMTSSTO, 0)) + CDbl(IIf(IsNumeric(!CLMMinorSurgery) = True, !CLMMinorSurgery, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPOthers) = True, !CLMPOthers, 0)) + CDbl(IIf(IsNumeric(!CLMBloodDrawing) = True, !CLMBloodDrawing, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMBloodGlucose) = True, !CLMBloodGlucose, 0)) + CDbl(IIf(IsNumeric(!CLMUrine) = True, !CLMUrine, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPapSmear) = True, !CLMPapSmear, 0)) + CDbl(IIf(IsNumeric(!CLMLOthers) = True, !CLMLOthers, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMDTPPolio) = True, !CLMDTPPolio, 0)) + CDbl(IIf(IsNumeric(!CLMDTPPolioHIB) = True, !CLMDTPPolioHIB, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMHepatitisB) = True, !CLMHepatitisB, 0)) + CDbl(IIf(IsNumeric(!CLMMMR) = True, !CLMMMR, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMIOthers) = True, !CLMIOthers, 0)) + CDbl(IIf(IsNumeric(!CLMDental) = True, !CLMDental, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMaternity) = True, !CLMMaternity, 0)) + CDbl(IIf(IsNumeric(!CLMPRVCheckup) = True, !CLMPRVCheckup, 0))
                        
                Fgrid.Row = 29
                Fgrid = Format(Total, "#,##0.00")
                              
            End With
            rst2.MoveNext
            Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CostDetailCLN(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    Total = 0
            
    'sql = " SELECT SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Clinic Where SUBSTRING(PAYCode,1,1) <> 'Z' "
        
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Clinic WHERE SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "39"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
            With rst2
                Fgrid.Col = 3
                Fgrid.Row = 1
                Fgrid.ColAlignment(1) = 6
                
                If Len(!CLMGPFeeBasicMed) Then
                    Fgrid = !CLMGPFeeBasicMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 2
                If Len(!CLMGPFeeMedInj) Then
                    Fgrid = !CLMGPFeeMedInj
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 3
                If Len(!CLMGPFeeLTMed) Then
                    Fgrid = !CLMGPFeeLTMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 4
                If Len(!CLMSPFeeBasicMed) Then
                    Fgrid = !CLMSPFeeBasicMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 5
                If Len(!CLMSPFeeMedInj) Then
                    Fgrid = !CLMSPFeeMedInj
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 6
                If Len(!CLMSPFeeLTMed) Then
                    Fgrid = !CLMSPFeeLTMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 7
                If Len(!CLMXRay) Then
                    Fgrid = !CLMXRay
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 8
                If Len(!CLMUltraSound) Then
                    Fgrid = !CLMUltraSound
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 9
                If Len(!CLMECG) Then
                    Fgrid = !CLMECG
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 10
                If Len(!CLMDressing) Then
                    Fgrid = !CLMDressing
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 11
                If Len(!CLMNebulizer) Then
                    Fgrid = !CLMNebulizer
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 12
                If Len(!CLMID) Then
                    Fgrid = !CLMID
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 13
                If Len(!CLMTSSTO) Then
                    Fgrid = !CLMTSSTO
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 14
                If Len(!CLMMinorSurgery) Then
                    Fgrid = !CLMMinorSurgery
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 15
                If Len(!CLMPOthers) Then
                    Fgrid = !CLMPOthers
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 16
                If Len(!CLMBloodDrawing) Then
                    Fgrid = !CLMBloodDrawing
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 17
                If Len(!CLMBloodGlucose) Then
                    Fgrid = !CLMBloodGlucose
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 18
                If Len(!CLMUrine) Then
                    Fgrid = !CLMUrine
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 19
                If Len(!CLMPapSmear) Then
                    Fgrid = !CLMPapSmear
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 20
                If Len(!CLMLOthers) Then
                    Fgrid = !CLMLOthers
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 21
                If Len(!CLMDTPPolio) Then
                    Fgrid = !CLMDTPPolio
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 22
                If Len(!CLMDTPPolioHIB) Then
                    Fgrid = !CLMDTPPolioHIB
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 23
                If Len(!CLMHepatitisB) Then
                    Fgrid = !CLMHepatitisB
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 24
                If Len(!CLMMMR) Then
                    Fgrid = !CLMMMR
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 25
                If Len(!CLMIOthers) Then
                    Fgrid = !CLMIOthers
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 26
                If Len(!CLMDental) Then
                    Fgrid = !CLMDental
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 27
                If Len(!CLMMaternity) Then
                    Fgrid = !CLMMaternity
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 28
                If Len(!CLMPRVCheckup) Then
                    Fgrid = !CLMPRVCheckup
                Else
                    Fgrid = 0
                End If
                    
                                
                Total = CDbl(IIf(IsNumeric(!CLMGPFeeBasicMed) = True, !CLMGPFeeBasicMed, 0)) + CDbl(IIf(IsNumeric(!CLMGPFeeMedInj) = True, !CLMGPFeeMedInj, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMGPFeeLTMed) = True, !CLMGPFeeLTMed, 0)) + CDbl(IIf(IsNumeric(!CLMSPFeeBasicMed) = True, !CLMSPFeeBasicMed, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMSPFeeMedInj) = True, !CLMSPFeeMedInj, 0)) + CDbl(IIf(IsNumeric(!CLMSPFeeLTMed) = True, !CLMSPFeeLTMed, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMXRay) = True, !CLMXRay, 0)) + CDbl(IIf(IsNumeric(!CLMUltraSound) = True, !CLMUltraSound, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMECG) = True, !CLMECG, 0)) + CDbl(IIf(IsNumeric(!CLMDressing) = True, !CLMDressing, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMNebulizer) = True, !CLMNebulizer, 0)) + CDbl(IIf(IsNumeric(!CLMID) = True, !CLMID, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMTSSTO) = True, !CLMTSSTO, 0)) + CDbl(IIf(IsNumeric(!CLMMinorSurgery) = True, !CLMMinorSurgery, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPOthers) = True, !CLMPOthers, 0)) + CDbl(IIf(IsNumeric(!CLMBloodDrawing) = True, !CLMBloodDrawing, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMBloodGlucose) = True, !CLMBloodGlucose, 0)) + CDbl(IIf(IsNumeric(!CLMUrine) = True, !CLMUrine, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPapSmear) = True, !CLMPapSmear, 0)) + CDbl(IIf(IsNumeric(!CLMLOthers) = True, !CLMLOthers, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMDTPPolio) = True, !CLMDTPPolio, 0)) + CDbl(IIf(IsNumeric(!CLMDTPPolioHIB) = True, !CLMDTPPolioHIB, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMHepatitisB) = True, !CLMHepatitisB, 0)) + CDbl(IIf(IsNumeric(!CLMMMR) = True, !CLMMMR, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMIOthers) = True, !CLMIOthers, 0)) + CDbl(IIf(IsNumeric(!CLMDental) = True, !CLMDental, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMaternity) = True, !CLMMaternity, 0)) + CDbl(IIf(IsNumeric(!CLMPRVCheckup) = True, !CLMPRVCheckup, 0))
                        
                
                Fgrid.Row = 29
                Fgrid = Format(Total, "#,##0.00")
                

            End With
            rst2.MoveNext
            Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CostDetailMBM(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    Total = 0
            
    'sql = " SELECT SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_MBM Where SUBSTRING(PAYCode,1,1) <> 'Z' "
        
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_MBM WHERE SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "40"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
            With rst2
                Fgrid.Col = 4
                Fgrid.Row = 1
                Fgrid.ColAlignment(1) = 6
                
                If Len(!CLMGPFeeBasicMed) Then
                    Fgrid = !CLMGPFeeBasicMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 2
                If Len(!CLMGPFeeMedInj) Then
                    Fgrid = !CLMGPFeeMedInj
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 3
                If Len(!CLMGPFeeLTMed) Then
                    Fgrid = !CLMGPFeeLTMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 4
                If Len(!CLMSPFeeBasicMed) Then
                    Fgrid = !CLMSPFeeBasicMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 5
                If Len(!CLMSPFeeMedInj) Then
                    Fgrid = !CLMSPFeeMedInj
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 6
                If Len(!CLMSPFeeLTMed) Then
                    Fgrid = !CLMSPFeeLTMed
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 7
                If Len(!CLMXRay) Then
                    Fgrid = !CLMXRay
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 8
                If Len(!CLMUltraSound) Then
                    Fgrid = !CLMUltraSound
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 9
                If Len(!CLMECG) Then
                    Fgrid = !CLMECG
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 10
                If Len(!CLMDressing) Then
                    Fgrid = !CLMDressing
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 11
                If Len(!CLMNebulizer) Then
                    Fgrid = !CLMNebulizer
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 12
                If Len(!CLMID) Then
                    Fgrid = !CLMID
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 13
                If Len(!CLMTSSTO) Then
                    Fgrid = !CLMTSSTO
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 14
                If Len(!CLMMinorSurgery) Then
                    Fgrid = !CLMMinorSurgery
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 15
                If Len(!CLMPOthers) Then
                    Fgrid = !CLMPOthers
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 16
                If Len(!CLMBloodDrawing) Then
                    Fgrid = !CLMBloodDrawing
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 17
                If Len(!CLMBloodGlucose) Then
                    Fgrid = !CLMBloodGlucose
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 18
                If Len(!CLMUrine) Then
                    Fgrid = !CLMUrine
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 19
                If Len(!CLMPapSmear) Then
                    Fgrid = !CLMPapSmear
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 20
                If Len(!CLMLOthers) Then
                    Fgrid = !CLMLOthers
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 21
                If Len(!CLMDTPPolio) Then
                    Fgrid = !CLMDTPPolio
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 22
                If Len(!CLMDTPPolioHIB) Then
                    Fgrid = !CLMDTPPolioHIB
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 23
                If Len(!CLMHepatitisB) Then
                    Fgrid = !CLMHepatitisB
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 24
                If Len(!CLMMMR) Then
                    Fgrid = !CLMMMR
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 25
                If Len(!CLMIOthers) Then
                    Fgrid = !CLMIOthers
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 26
                If Len(!CLMDental) Then
                    Fgrid = !CLMDental
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 27
                If Len(!CLMMaternity) Then
                    Fgrid = !CLMMaternity
                Else
                    Fgrid = 0
                End If
                
                Fgrid.Row = 28
                If Len(!CLMPRVCheckup) Then
                    Fgrid = !CLMPRVCheckup
                Else
                    Fgrid = 0
                End If
                    
                                
                Total = CDbl(IIf(IsNumeric(!CLMGPFeeBasicMed) = True, !CLMGPFeeBasicMed, 0)) + CDbl(IIf(IsNumeric(!CLMGPFeeMedInj) = True, !CLMGPFeeMedInj, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMGPFeeLTMed) = True, !CLMGPFeeLTMed, 0)) + CDbl(IIf(IsNumeric(!CLMSPFeeBasicMed) = True, !CLMSPFeeBasicMed, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMSPFeeMedInj) = True, !CLMSPFeeMedInj, 0)) + CDbl(IIf(IsNumeric(!CLMSPFeeLTMed) = True, !CLMSPFeeLTMed, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMXRay) = True, !CLMXRay, 0)) + CDbl(IIf(IsNumeric(!CLMUltraSound) = True, !CLMUltraSound, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMECG) = True, !CLMECG, 0)) + CDbl(IIf(IsNumeric(!CLMDressing) = True, !CLMDressing, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMNebulizer) = True, !CLMNebulizer, 0)) + CDbl(IIf(IsNumeric(!CLMID) = True, !CLMID, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMTSSTO) = True, !CLMTSSTO, 0)) + CDbl(IIf(IsNumeric(!CLMMinorSurgery) = True, !CLMMinorSurgery, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPOthers) = True, !CLMPOthers, 0)) + CDbl(IIf(IsNumeric(!CLMBloodDrawing) = True, !CLMBloodDrawing, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMBloodGlucose) = True, !CLMBloodGlucose, 0)) + CDbl(IIf(IsNumeric(!CLMUrine) = True, !CLMUrine, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPapSmear) = True, !CLMPapSmear, 0)) + CDbl(IIf(IsNumeric(!CLMLOthers) = True, !CLMLOthers, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMDTPPolio) = True, !CLMDTPPolio, 0)) + CDbl(IIf(IsNumeric(!CLMDTPPolioHIB) = True, !CLMDTPPolioHIB, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMHepatitisB) = True, !CLMHepatitisB, 0)) + CDbl(IIf(IsNumeric(!CLMMMR) = True, !CLMMMR, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMIOthers) = True, !CLMIOthers, 0)) + CDbl(IIf(IsNumeric(!CLMDental) = True, !CLMDental, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMaternity) = True, !CLMMaternity, 0)) + CDbl(IIf(IsNumeric(!CLMPRVCheckup) = True, !CLMPRVCheckup, 0))
                        
                
                Fgrid.Row = 29
                Fgrid = Format(Total, "#,##0.00")
                                                
            End With
            rst2.MoveNext
            Loop
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ActualGrpCompany(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT MBMGroupCompany, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by MBMGroupCompany Order by MBMGroupCompany "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> 'Z' "

    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "9"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!MBMGroupCompany) Then
                    Set Record = ListView2.ListItems.Add(, , Trim(!MBMGroupCompany))
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
   
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedClaim(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    '====================
    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    '==================
    Total = 0
    Current = 0
        
    ListView2.ListItems.Clear
            
    'sql = " SELECT CASNumber, MBMNumber, PatientName, DateVisit, " & _
          " CLMGPFeeBasicMed, CLMGPFeeMedInj, CLMGPFeeLTMed, CLMSPFeeBasicMed, " & _
          " CLMSPFeeMedInj, CLMSPFeeLTMed, CLMXRay, CLMUltraSound, CLMECG, CLMDressing, " & _
          " CLMNebulizer, CLMID, CLMTSSTO, CLMMinorSurgery, CLMPOthers, " & _
          " CLMBloodDrawing, CLMBloodGlucose, CLMUrine, CLMPapSmear, CLMLOthers, " & _
          " CLMDTPPolio, CLMDTPPolioHIB, CLMHepatitisB, CLMMMR, CLMIOthers, " & _
          " CLMDental, CLMMaternity, CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
        
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "23"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!CasNumber) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!CasNumber)
                End If
                                              
                If Len(rst2!MBMNumber) Then
                   Record.SubItems(1) = rst2!MBMNumber
                End If
                
                If Len(rst2!PatientName) Then
                   Record.SubItems(2) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!DateVisit) Then
                    Record.SubItems(3) = Format(rst2!DateVisit, "dd/mm/yyyy")
                End If
                                
                If Len(rst2!CLMGPFeeBasicMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(7) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(7) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(8) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(8) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(9) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(9) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(10) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(10) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(11) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(11) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(12) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(12) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(13) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(13) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(14) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(14) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(15) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(15) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(16) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(16) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(17) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(17) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(18) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(18) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(19) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(19) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(20) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(20) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(21) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(21) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(22) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(22) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(23) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(23) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(24) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(24) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(25) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(25) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(26) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(26) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(27) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(27) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(28) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(28) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(29) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(29) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(30) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(30) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(31) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedInsPlan(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT INSCode, PLNCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by INSCode, PLNCode Order by INSCode, PLNCode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "29"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , !INSCode)
                End If
                                              
                If Len(!PLNCode) Then
                    Record.SubItems(1) = Trim(!PLNCode)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(2) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(2) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(3) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(7) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(7) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(8) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(8) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(9) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(9) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(10) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(10) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(11) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(11) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(12) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(12) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(13) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(13) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(14) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(14) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(15) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(15) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(16) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(17) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(18) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(18) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(19) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(19) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(20) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(20) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(21) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(22) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(23) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(23) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(24) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(24) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(25) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(25) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(26) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(26) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(27) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(27) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(28) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(28) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(29) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedInsAge(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT INSCode, AGEBand, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by INSCode, AGEBand Order by INSCode, Ageband "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "28"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , !INSCode)
                End If
                                              
                If Len(!AgeBand) Then
                    Record.SubItems(1) = Trim(!AgeBand)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(2) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(2) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(3) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(7) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(7) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(8) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(8) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(9) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(9) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(10) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(10) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(11) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(11) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(12) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(12) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(13) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(13) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(14) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(14) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(15) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(15) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(16) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(17) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(18) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(18) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(19) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(19) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(20) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(20) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(21) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(22) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(23) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(23) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(24) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(24) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(25) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(25) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(26) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(26) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(27) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(27) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(28) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(28) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(29) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedInsSex(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT INSCode, Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by INSCode, Sex Order by INSCode, Sex "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "31"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , !INSCode)
                End If
                                              
                If Len(!Sex) Then
                    Record.SubItems(1) = Trim(!Sex)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(2) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(2) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(3) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(7) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(7) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(8) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(8) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(9) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(9) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(10) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(10) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(11) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(11) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(12) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(12) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(13) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(13) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(14) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(14) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(15) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(15) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(16) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(17) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(18) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(18) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(19) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(19) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(20) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(20) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(21) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(22) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(23) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(23) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(24) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(24) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(25) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(25) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(26) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(26) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(27) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(27) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(28) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(28) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(29) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedAgeband(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT AGEBand, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by AGEBand Order by AGEBand "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "20"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!AgeBand) Then
                    Set Record = ListView2.ListItems.Add(, , !AgeBand)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedSex(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by Sex Order by Sex "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "35"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!Sex) Then
                    Set Record = ListView2.ListItems.Add(, , !Sex)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedAgebandSex(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT AGEBand, Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by AGEBand, Sex Order by Ageband, Sex "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "21"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!AgeBand) Then
                    Set Record = ListView2.ListItems.Add(, , !AgeBand)
                End If
                                              
                If Len(!Sex) Then
                    Record.SubItems(1) = Trim(!Sex)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(2) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(2) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(3) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(7) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(7) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(8) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(8) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(9) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(9) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(10) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(10) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(11) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(11) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(12) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(12) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(13) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(13) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(14) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(14) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(15) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(15) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(16) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(17) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(18) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(18) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(19) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(19) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(20) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(20) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(21) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(22) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(23) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(23) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(24) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(24) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(25) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(25) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(26) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(26) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(27) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(27) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(28) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(28) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(29) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
   
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedFinalDiag(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT CASTreatReason1, ReasonDesc, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by CASTreatReason1, ReasonDesc Order by CASTreatReason1, ReasonDesc "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "25"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!CASTreatReason1) Then
                    Set Record = ListView2.ListItems.Add(, , !CASTreatReason1)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                              
                If Len(!ReasonDesc) Then
                    Record.SubItems(1) = Trim(!ReasonDesc)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(2) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(2) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(3) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(7) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(7) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(8) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(8) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(9) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(9) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(10) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(10) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(11) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(11) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(12) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(12) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(13) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(13) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(14) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(14) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(15) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(15) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(16) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(17) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(18) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(18) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(19) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(19) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(20) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(20) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(21) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(22) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(23) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(23) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(24) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(24) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(25) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(25) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(26) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(26) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(27) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(27) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(28) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(28) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(29) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedClinic(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT ClinicCode, CLNName, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by ClinicCode, CLNName Order by ClinicCode, CLNName "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "24"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!ClinicCode) Then
                    Set Record = ListView2.ListItems.Add(, , !ClinicCode)
                End If
                                              
                If Len(!CLNName) Then
                    Record.SubItems(1) = Trim(!CLNName)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(2) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(2) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(3) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(7) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(7) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(8) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(8) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(9) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(9) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(10) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(10) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(11) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(11) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(12) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(12) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(13) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(13) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(14) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(14) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(15) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(15) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(16) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(17) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(18) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(18) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(19) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(19) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(20) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(20) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(21) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(22) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(23) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(23) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(24) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(24) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(25) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(25) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(26) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(26) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(27) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(27) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(28) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(28) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(29) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
   
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedPayor(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT PAYCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by PAYCode Order by PAYCode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "32"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0

    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedPayorInsPlan(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT PAYCode, INSCode, PLNCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by PAYCode, INSCode, PLNCode Order by PAYCode, INSCode, PLNCode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "33"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                End If
                                              
                If Len(!INSCode) Then
                    Record.SubItems(1) = Trim(!INSCode)
                End If
                
                If Len(!PLNCode) Then
                    Record.SubItems(2) = Trim(!PLNCode)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(3) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(6) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(7) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(7) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(8) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(8) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(9) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(9) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(10) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(10) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(11) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(11) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(12) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(12) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(13) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(13) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(14) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(14) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(15) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(15) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(16) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(16) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(17) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(18) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(18) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(19) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(19) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(20) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(20) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(21) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(21) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(22) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(23) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(23) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(24) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(24) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(25) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(25) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(26) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(26) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(27) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(27) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(28) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(28) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(29) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(29) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(30) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0

    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedGrpCompany(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT MBMGroupCompany, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by MBMGroupCompany Order by MBMGroupCompany "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "26"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!MBMGroupCompany) Then
                    Set Record = ListView2.ListItems.Add(, , Trim(!MBMGroupCompany))
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedInsRelation(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT INSCode, Relation, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by INSCode, Relation Order by INSCode, Relation "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "30"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , !INSCode)
                End If
                                              
                If Len(!Relation) Then
                    Record.SubItems(1) = Trim(!Relation)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(2) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(2) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(3) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(4) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(4) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(5) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(6) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(6) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(7) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(7) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(8) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(8) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(9) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(9) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(10) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(10) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(11) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(11) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(12) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(12) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(13) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(13) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(14) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(14) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(15) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(15) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(16) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(17) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(17) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(18) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(18) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(19) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(19) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(20) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(20) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(21) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(22) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(22) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(23) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(23) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(24) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(24) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(25) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(25) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(26) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(26) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(27) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(27) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(28) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(28) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(29) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0

    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedIns(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT INSCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by INSCode Order by INSCode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "27"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , !INSCode)
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
   
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedBordxNo(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT BordxRefNo, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by BordxRefNo Order by BordxRefNo "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "22"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!Bordxrefno) Then
                    Set Record = ListView2.ListItems.Add(, , !Bordxrefno)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedPolYear(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT MBMPolicyYear, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by MBMPolicyYear Order by MBMPolicyYear "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "34"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!MBMPolicyYear) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMPolicyYear)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0

    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ApprovedTakeover(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    GPFeeBasicMed = 0
    GPFeeMedInj = 0
    GPFeeLTMed = 0
    SPFeeBasicMed = 0
    SPFeeMedInj = 0
    SPFeeLTMed = 0
    XRay = 0
    UltraSound = 0
    ECG = 0
    Dressing = 0
    Nebulizer = 0
    ID = 0
    TSSTO = 0
    MinorSurgery = 0
    POthers = 0
    BloodDrawing = 0
    BloodGlucose = 0
    Urine = 0
    PapSmear = 0
    LOthers = 0
    DTPPolio = 0
    DTPPolioHIB = 0
    HepatitisB = 0
    MMR = 0
    IOthers = 0
    Dental = 0
    Maternity = 0
    PRVCheckup = 0
    
    Total = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT MBMTakeOver, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj, " & _
          " SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, " & _
          " SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, " & _
          " SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, " & _
          " SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, " & _
          " SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, " & _
          " SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB, " & _
          " SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, " & _
          " SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup " & _
          " From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'sql2 = " Group by MBMTakeOver Order by MBMTakeOver "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> 'Z' "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    '    strCount = strCount + strAppend
    'Else
    '    sql = sql + sql2
    '    strCount = strCount + strAppend
    'End If
    
    'Set rstCount = MediNeptuneConnCISWS.Execute(strCount)
    'rst2.Open sql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "36"

    Set cmd.ActiveConnection = MediNeptuneConnCISWS
    cmd.CommandText = "CISClaimTracking2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            
            With rst2
                If Len(!MBMTakeOver) Then
                    Set Record = ListView2.ListItems.Add(, , Trim(!MBMTakeOver))
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(!CLMGPFeeBasicMed) Then
                    Record.SubItems(1) = Format(!CLMGPFeeBasicMed, "#,##0.00")
                    GPFeeBasicMed = rst2!CLMGPFeeBasicMed
                Else
                    Record.SubItems(1) = "0"
                    GPFeeBasicMed = 0
                End If
                
                If Len(rst2!CLMGPFeeMedInj) Then
                    Record.SubItems(2) = Format(rst2!CLMGPFeeMedInj, "#,##0.00")
                    GPFeeMedInj = rst2!CLMGPFeeMedInj
                Else
                    Record.SubItems(2) = "0"
                    GPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMGPFeeLTMed) Then
                    Record.SubItems(3) = Format(rst2!CLMGPFeeLTMed, "#,##0.00")
                    GPFeeLTMed = rst2!CLMGPFeeLTMed
                Else
                    Record.SubItems(3) = "0"
                    GPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMSPFeeMedInj) Then
                    Record.SubItems(4) = Format(rst2!CLMSPFeeMedInj, "#,##0.00")
                    SPFeeMedInj = rst2!CLMSPFeeMedInj
                Else
                    Record.SubItems(4) = "0"
                    SPFeeMedInj = 0
                End If
                
                If Len(rst2!CLMSPFeeLTMed) Then
                    Record.SubItems(5) = Format(rst2!CLMSPFeeLTMed, "#,##0.00")
                    SPFeeLTMed = rst2!CLMSPFeeLTMed
                Else
                    Record.SubItems(5) = "0"
                    SPFeeLTMed = 0
                End If
                
                If Len(rst2!CLMXRay) Then
                    Record.SubItems(6) = Format(rst2!CLMXRay, "#,##0.00")
                    XRay = rst2!CLMXRay
                Else
                    Record.SubItems(6) = "0"
                    XRay = 0
                End If
                
                If Len(rst2!CLMUltraSound) Then
                    Record.SubItems(7) = Format(rst2!CLMUltraSound, "#,##0.00")
                    UltraSound = rst2!CLMUltraSound
                Else
                    Record.SubItems(7) = "0"
                    UltraSound = 0
                End If
                
                If Len(rst2!CLMECG) Then
                    Record.SubItems(8) = Format(rst2!CLMECG, "#,##0.00")
                    ECG = rst2!CLMECG
                Else
                    Record.SubItems(8) = "0"
                    ECG = 0
                End If
                
                If Len(rst2!CLMDressing) Then
                    Record.SubItems(9) = Format(rst2!CLMDressing, "#,##0.00")
                    Dressing = rst2!CLMDressing
                Else
                    Record.SubItems(9) = "0"
                    Dressing = 0
                End If
                
                If Len(rst2!CLMNebulizer) Then
                    Record.SubItems(10) = Format(rst2!CLMNebulizer, "#,##0.00")
                    Nebulizer = rst2!CLMNebulizer
                Else
                    Record.SubItems(10) = "0"
                    Nebulizer = 0
                End If
                
                If Len(rst2!CLMID) Then
                    Record.SubItems(11) = Format(rst2!CLMID, "#,##0.00")
                    ID = rst2!CLMID
                Else
                    Record.SubItems(11) = "0"
                    ID = 0
                End If
                
                If Len(rst2!CLMTSSTO) Then
                    Record.SubItems(12) = Format(rst2!CLMTSSTO, "#,##0.00")
                    TSSTO = rst2!CLMTSSTO
                Else
                    Record.SubItems(12) = "0"
                    TSSTO = 0
                End If
                
                If Len(rst2!CLMMinorSurgery) Then
                    Record.SubItems(13) = Format(rst2!CLMMinorSurgery, "#,##0.00")
                    MinorSurgery = rst2!CLMMinorSurgery
                Else
                    Record.SubItems(13) = "0"
                    MinorSurgery = 0
                End If
                
                If Len(rst2!CLMPOthers) Then
                    Record.SubItems(14) = Format(rst2!CLMPOthers, "#,##0.00")
                    POthers = rst2!CLMPOthers
                Else
                    Record.SubItems(14) = "0"
                    POthers = 0
                End If
                
                If Len(rst2!CLMBloodDrawing) Then
                    Record.SubItems(15) = Format(rst2!CLMBloodDrawing, "#,##0.00")
                    BloodDrawing = rst2!CLMBloodDrawing
                Else
                    Record.SubItems(15) = "0"
                    BloodDrawing = 0
                End If
                
                If Len(rst2!CLMBloodGlucose) Then
                    Record.SubItems(16) = Format(rst2!CLMBloodGlucose, "#,##0.00")
                    BloodGlucose = rst2!CLMBloodGlucose
                Else
                    Record.SubItems(16) = "0"
                    BloodGlucose = 0
                End If
                
                If Len(rst2!CLMUrine) Then
                    Record.SubItems(17) = Format(rst2!CLMUrine, "#,##0.00")
                    Urine = rst2!CLMUrine
                Else
                    Record.SubItems(17) = "0"
                    Urine = 0
                End If
                
                If Len(rst2!CLMPapSmear) Then
                    Record.SubItems(18) = Format(rst2!CLMPapSmear, "#,##0.00")
                    PapSmear = rst2!CLMPapSmear
                Else
                    Record.SubItems(18) = "0"
                    PapSmear = 0
                End If
                
                If Len(rst2!CLMLOthers) Then
                    Record.SubItems(19) = Format(rst2!CLMLOthers, "#,##0.00")
                    LOthers = rst2!CLMLOthers
                Else
                    Record.SubItems(19) = "0"
                    LOthers = 0
                End If
                
                If Len(rst2!CLMDTPPolio) Then
                    Record.SubItems(20) = Format(rst2!CLMDTPPolio, "#,##0.00")
                    DTPPolio = rst2!CLMDTPPolio
                Else
                    Record.SubItems(20) = "0"
                    DTPPolio = rst2!CLMDTPPolio
                End If
                
                If Len(rst2!CLMDTPPolioHIB) Then
                    Record.SubItems(21) = Format(rst2!CLMDTPPolioHIB, "#,##0.00")
                    DTPPolioHIB = rst2!CLMDTPPolioHIB
                Else
                    Record.SubItems(21) = "0"
                    DTPPolioHIB = 0
                End If
                
                If Len(rst2!CLMHepatitisB) Then
                    Record.SubItems(22) = Format(rst2!CLMHepatitisB, "#,##0.00")
                    HepatitisB = rst2!CLMHepatitisB
                Else
                    Record.SubItems(22) = "0"
                    HepatitisB = 0
                End If
                
                If Len(rst2!CLMMMR) Then
                    Record.SubItems(23) = Format(rst2!CLMMMR, "#,##0.00")
                    MMR = rst2!CLMMMR
                Else
                    Record.SubItems(23) = "0"
                    MMR = 0
                End If
                
                If Len(rst2!CLMIOthers) Then
                    Record.SubItems(24) = Format(rst2!CLMIOthers, "#,##0.00")
                    IOthers = rst2!CLMIOthers
                Else
                    Record.SubItems(24) = "0"
                    IOthers = 0
                End If
                
                If Len(rst2!CLMDental) Then
                    Record.SubItems(25) = Format(rst2!CLMDental, "#,##0.00")
                    Dental = rst2!CLMDental
                Else
                    Record.SubItems(25) = "0"
                    Dental = 0
                End If
                
                If Len(rst2!CLMMaternity) Then
                    Record.SubItems(26) = Format(rst2!CLMMaternity, "#,##0.00")
                    Maternity = rst2!CLMMaternity
                Else
                    Record.SubItems(26) = "0"
                    Maternity = 0
                End If
                
                If Len(rst2!CLMPRVCheckup) Then
                    Record.SubItems(27) = Format(rst2!CLMPRVCheckup, "#,##0.00")
                    PRVCheckup = rst2!CLMPRVCheckup
                Else
                    Record.SubItems(27) = "0"
                    PRVCheckup = 0
                End If
                
                Total = CDbl(GPFeeBasicMed) + CDbl(GPFeeMedInj) + CDbl(GPFeeLTMed) + CDbl(SPFeeBasicMed) + CDbl(SPFeeMedInj) + CDbl(SPFeeLTMed) + CDbl(XRay) _
                    + CDbl(UltraSound) + CDbl(ECG) + CDbl(Dressing) + CDbl(Nebulizer) + CDbl(ID) + CDbl(TSSTO) _
                    + CDbl(MinorSurgery) + CDbl(POthers) + CDbl(BloodDrawing) + CDbl(BloodGlucose) + CDbl(Urine) + CDbl(PapSmear) _
                    + CDbl(LOthers) + CDbl(DTPPolio) + CDbl(DTPPolioHIB) + CDbl(HepatitisB) + CDbl(MMR) + CDbl(IOthers) _
                    + CDbl(Dental) + CDbl(Maternity) + CDbl(PRVCheckup)
                        
                Record.SubItems(28) = Format(Total, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Public Function ExportToExcelWorksheet()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet

    Screen.MousePointer = vbHourglass
    
    'Set objExcel = New excel.Application
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
    
    objExcel.WindowState = xlMaximized

    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\ClaimReport.xlt")
    Set objWorksheet = objWorkbook.Sheets(1)
    'Turn off the alerts, otherwise user will have to confirm my actions.
    objExcel.DisplayAlerts = False
    
    Set objWorksheet = objWorkbook.Worksheets.Item(1)
    objWorksheet.Cells(3, "A") = Trim(cboPayorCode)
        
    '============== BENEFIT =======================
    Fgrid.Col = 0
    Fgrid.Row = 0
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "A") = Fgrid
    End If
    Fgrid.Row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "A") = Fgrid
    End If
    Fgrid.Row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "A") = Fgrid
    End If
    
    Fgrid.Row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "A") = Fgrid
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "A") = Fgrid
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "A") = Fgrid
    End If
    Fgrid.Row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "A") = Fgrid
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "A") = Fgrid
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "A") = Fgrid
    End If
    Fgrid.Row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "A") = Fgrid
    End If
    Fgrid.Row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "A") = Fgrid
    End If
    Fgrid.Row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "A") = Fgrid
    End If
    Fgrid.Row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "A") = Fgrid
    End If
    Fgrid.Row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "A") = Fgrid
    End If
    Fgrid.Row = 14
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(19, "A") = Fgrid
    End If
    Fgrid.Row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "A") = Fgrid
    End If
    Fgrid.Row = 16
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(21, "A") = Fgrid
    End If
    Fgrid.Row = 17
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(22, "A") = Fgrid
    End If
    Fgrid.Row = 18
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(23, "A") = Fgrid
    End If
    Fgrid.Row = 19
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(24, "A") = Fgrid
    End If
    Fgrid.Row = 20
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(25, "A") = Fgrid
    End If
    Fgrid.Row = 21
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(26, "A") = Fgrid
    End If
    Fgrid.Row = 22
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(27, "A") = Fgrid
    End If
    Fgrid.Row = 23
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(28, "A") = Fgrid
    End If
    Fgrid.Row = 24
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(29, "A") = Fgrid
    End If
    Fgrid.Row = 25
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(30, "A") = Fgrid
    End If
    Fgrid.Row = 26
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(31, "A") = Fgrid
    End If
    Fgrid.Row = 27
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(32, "A") = Fgrid
    End If
    Fgrid.Row = 28
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(33, "A") = Fgrid
    End If
    Fgrid.Row = 29
    Set objRange = objWorksheet.Rows(34)
    objRange.Font.Size = 12
    objRange.Font.Bold = True
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(34, "A") = Fgrid
    End If
    '============== BENEFIT =======================
    
    '============== ACTUAL =======================
    Fgrid.Col = 1
    Fgrid.Row = 0
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "B") = Fgrid
    End If
    Fgrid.Row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "B") = Fgrid
    End If
    Fgrid.Row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "B") = Fgrid
    End If
    
    Fgrid.Row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "B") = Fgrid
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "B") = Fgrid
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "B") = Fgrid
    End If
    Fgrid.Row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "B") = Fgrid
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "B") = Fgrid
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "B") = Fgrid
    End If
    Fgrid.Row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "B") = Fgrid
    End If
    Fgrid.Row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "B") = Fgrid
    End If
    Fgrid.Row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "B") = Fgrid
    End If
    Fgrid.Row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "B") = Fgrid
    End If
    Fgrid.Row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "B") = Fgrid
    End If
    Fgrid.Row = 14
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(19, "B") = Fgrid
    End If
    Fgrid.Row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "B") = Fgrid
    End If
    Fgrid.Row = 16
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(21, "B") = Fgrid
    End If
    Fgrid.Row = 17
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(22, "B") = Fgrid
    End If
    Fgrid.Row = 18
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(23, "B") = Fgrid
    End If
    Fgrid.Row = 19
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(24, "B") = Fgrid
    End If
    Fgrid.Row = 20
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(25, "B") = Fgrid
    End If
    Fgrid.Row = 21
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(26, "B") = Fgrid
    End If
    Fgrid.Row = 22
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(27, "B") = Fgrid
    End If
    Fgrid.Row = 23
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(28, "B") = Fgrid
    End If
    Fgrid.Row = 24
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(29, "B") = Fgrid
    End If
    Fgrid.Row = 25
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(30, "B") = Fgrid
    End If
    Fgrid.Row = 26
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(31, "B") = Fgrid
    End If
    Fgrid.Row = 27
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(32, "B") = Fgrid
    End If
    Fgrid.Row = 28
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(33, "B") = Fgrid
    End If
    Fgrid.Row = 29
    Set objRange = objWorksheet.Rows(34)
    objRange.Font.Size = 12
    objRange.Font.Bold = True
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(34, "B") = Fgrid
    End If
    '============== ACTUAL =======================
    
    '============== APPROVED =======================
    Fgrid.Col = 2
    Fgrid.Row = 0
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "C") = Fgrid
    End If
    Fgrid.Row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "C") = Fgrid
    End If
    Fgrid.Row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "C") = Fgrid
    End If
    
    Fgrid.Row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "C") = Fgrid
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "C") = Fgrid
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "C") = Fgrid
    End If
    Fgrid.Row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "C") = Fgrid
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "C") = Fgrid
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "C") = Fgrid
    End If
    Fgrid.Row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "C") = Fgrid
    End If
    Fgrid.Row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "C") = Fgrid
    End If
    Fgrid.Row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "C") = Fgrid
    End If
    Fgrid.Row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "C") = Fgrid
    End If
    Fgrid.Row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "C") = Fgrid
    End If
    Fgrid.Row = 14
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(19, "C") = Fgrid
    End If
    Fgrid.Row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "C") = Fgrid
    End If
    Fgrid.Row = 16
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(21, "C") = Fgrid
    End If
    Fgrid.Row = 17
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(22, "C") = Fgrid
    End If
    Fgrid.Row = 18
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(23, "C") = Fgrid
    End If
    Fgrid.Row = 19
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(24, "C") = Fgrid
    End If
    Fgrid.Row = 20
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(25, "C") = Fgrid
    End If
    Fgrid.Row = 21
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(26, "C") = Fgrid
    End If
    Fgrid.Row = 22
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(27, "C") = Fgrid
    End If
    Fgrid.Row = 23
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(28, "C") = Fgrid
    End If
    Fgrid.Row = 24
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(29, "C") = Fgrid
    End If
    Fgrid.Row = 25
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(30, "C") = Fgrid
    End If
    Fgrid.Row = 26
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(31, "C") = Fgrid
    End If
    Fgrid.Row = 27
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(32, "C") = Fgrid
    End If
    Fgrid.Row = 28
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(33, "C") = Fgrid
    End If
    Fgrid.Row = 29
    Set objRange = objWorksheet.Rows(34)
    objRange.Font.Size = 12
    objRange.Font.Bold = True
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(34, "C") = Fgrid
    End If
    '============== APPROVED =======================
    
    '============== PAY2CLINIC =======================
    Fgrid.Col = 3
    Fgrid.Row = 0
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "D") = Fgrid
    End If
    Fgrid.Row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "D") = Fgrid
    End If
    Fgrid.Row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "D") = Fgrid
    End If
    
    Fgrid.Row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "D") = Fgrid
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "D") = Fgrid
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "D") = Fgrid
    End If
    Fgrid.Row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "D") = Fgrid
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "D") = Fgrid
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "D") = Fgrid
    End If
    Fgrid.Row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "D") = Fgrid
    End If
    Fgrid.Row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "D") = Fgrid
    End If
    Fgrid.Row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "D") = Fgrid
    End If
    Fgrid.Row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "D") = Fgrid
    End If
    Fgrid.Row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "D") = Fgrid
    End If
    Fgrid.Row = 14
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(19, "D") = Fgrid
    End If
    Fgrid.Row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "D") = Fgrid
    End If
    Fgrid.Row = 16
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(21, "D") = Fgrid
    End If
    Fgrid.Row = 17
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(22, "D") = Fgrid
    End If
    Fgrid.Row = 18
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(23, "D") = Fgrid
    End If
    Fgrid.Row = 19
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(24, "D") = Fgrid
    End If
    Fgrid.Row = 20
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(25, "D") = Fgrid
    End If
    Fgrid.Row = 21
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(26, "D") = Fgrid
    End If
    Fgrid.Row = 22
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(27, "D") = Fgrid
    End If
    Fgrid.Row = 23
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(28, "D") = Fgrid
    End If
    Fgrid.Row = 24
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(29, "D") = Fgrid
    End If
    Fgrid.Row = 25
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(30, "D") = Fgrid
    End If
    Fgrid.Row = 26
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(31, "D") = Fgrid
    End If
    Fgrid.Row = 27
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(32, "D") = Fgrid
    End If
    Fgrid.Row = 28
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(33, "D") = Fgrid
    End If
    Fgrid.Row = 29
    Set objRange = objWorksheet.Rows(34)
    objRange.Font.Size = 12
    objRange.Font.Bold = True
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(34, "D") = Fgrid
    End If
    '============== PAY2CLINIC =======================
    
    '============== PAY2/FR MEM =======================
    Fgrid.Col = 4
    Fgrid.Row = 0
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "E") = Fgrid
    End If
    Fgrid.Row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "E") = Fgrid
    End If
    Fgrid.Row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "E") = Fgrid
    End If
    
    Fgrid.Row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "E") = Fgrid
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "E") = Fgrid
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "E") = Fgrid
    End If
    Fgrid.Row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "E") = Fgrid
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "E") = Fgrid
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "E") = Fgrid
    End If
    Fgrid.Row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "E") = Fgrid
    End If
    Fgrid.Row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "E") = Fgrid
    End If
    Fgrid.Row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "E") = Fgrid
    End If
    Fgrid.Row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "E") = Fgrid
    End If
    Fgrid.Row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "E") = Fgrid
    End If
    Fgrid.Row = 14
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(19, "E") = Fgrid
    End If
    Fgrid.Row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "E") = Fgrid
    End If
    Fgrid.Row = 16
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(21, "E") = Fgrid
    End If
    Fgrid.Row = 17
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(22, "E") = Fgrid
    End If
    Fgrid.Row = 18
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(23, "E") = Fgrid
    End If
    Fgrid.Row = 19
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(24, "E") = Fgrid
    End If
    Fgrid.Row = 20
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(25, "E") = Fgrid
    End If
    Fgrid.Row = 21
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(26, "E") = Fgrid
    End If
    Fgrid.Row = 22
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(27, "E") = Fgrid
    End If
    Fgrid.Row = 23
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(28, "E") = Fgrid
    End If
    Fgrid.Row = 24
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(29, "E") = Fgrid
    End If
    Fgrid.Row = 25
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(30, "E") = Fgrid
    End If
    Fgrid.Row = 26
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(31, "E") = Fgrid
    End If
    Fgrid.Row = 27
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(32, "E") = Fgrid
    End If
    Fgrid.Row = 28
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(33, "E") = Fgrid
    End If
    Fgrid.Row = 29
    Set objRange = objWorksheet.Rows(34)
    objRange.Font.Size = 12
    objRange.Font.Bold = True
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(34, "E") = Fgrid
    End If
    '============== PAY2/FR MEM =======================
        
    Screen.MousePointer = vbDefault
    
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    
    Exit Function
ExitFunction:
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    Screen.MousePointer = vbDefault
    objExcel.Quit
    Exit Function
HANDLE_ERROR:
    MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
    Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    GoTo ExitFunction
    
End Function
