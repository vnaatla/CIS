VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMBMEnquiry 
   Caption         =   "Member Enquiry"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12015
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   15
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Member Enquiry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   6615
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11655
      Begin VB.CommandButton cmdSearchMBM 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   3
         Top             =   225
         Width           =   675
      End
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Go"
         Height          =   285
         Left            =   3240
         TabIndex        =   1
         Top             =   225
         Width           =   765
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Pre"
         Height          =   285
         Left            =   3960
         TabIndex        =   4
         Top             =   225
         Width           =   765
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
         Height          =   285
         Left            =   4680
         TabIndex        =   5
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Label25 
         Caption         =   "Pol Status"
         Height          =   330
         Left            =   6240
         TabIndex        =   12
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label76 
         Caption         =   "Mem Num."
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label Label90 
         Caption         =   "Data Status"
         Height          =   330
         Left            =   9660
         TabIndex        =   10
         Top             =   240
         Width           =   915
      End
      Begin VB.Label txtDataStatus 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10580
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label103 
         Caption         =   "Exp. Status"
         Height          =   330
         Left            =   8360
         TabIndex        =   8
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblEXPStatus 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9260
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.Label txtMBMStatus 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7040
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5445
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9604
      _Version        =   393216
      Tabs            =   11
      TabsPerRow      =   11
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
      TabCaption(0)   =   "Principal &Detail"
      TabPicture(0)   =   "frmMBMEnquiry.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Other Details"
      TabPicture(1)   =   "frmMBMEnquiry.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Supp&lementary"
      TabPicture(2)   =   "frmMBMEnquiry.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstCovAdd"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&Adjustment History"
      TabPicture(3)   =   "frmMBMEnquiry.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstAdjustmentHis"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "HIS &Benefits"
      TabPicture(4)   =   "frmMBMEnquiry.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label105"
      Tab(4).Control(1)=   "Label37"
      Tab(4).Control(2)=   "lstNewBenefitDetails"
      Tab(4).Control(3)=   "lstBenefitDetails"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "&Case"
      TabPicture(5)   =   "frmMBMEnquiry.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lstCaseList"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Mem &History"
      TabPicture(6)   =   "frmMBMEnquiry.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label26"
      Tab(6).Control(1)=   "lstMBMHis"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Account"
      TabPicture(7)   =   "frmMBMEnquiry.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lstMBMAcc"
      Tab(7).Control(1)=   "Frame11"
      Tab(7).Control(2)=   "Frame12"
      Tab(7).Control(3)=   "Frame13"
      Tab(7).Control(4)=   "cmdTotal"
      Tab(7).Control(5)=   "Frame9"
      Tab(7).ControlCount=   6
      TabCaption(8)   =   "Limitation"
      TabPicture(8)   =   "frmMBMEnquiry.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lstLimitation"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Surgical Proc"
      TabPicture(9)   =   "frmMBMEnquiry.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "lstSurgicalProc"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Fees"
      TabPicture(10)  =   "frmMBMEnquiry.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame2"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).ControlCount=   1
      Begin VB.Frame Frame9 
         Caption         =   "Payable To Clinic"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   202
         Top             =   1800
         Width           =   5055
         Begin VB.TextBox txtPayable2Hos 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   203
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtPaid2Hos 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   204
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtHosOS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   205
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label68 
            Caption         =   "Paid Amount"
            Height          =   255
            Left            =   240
            TabIndex        =   208
            Top             =   495
            Width           =   945
         End
         Begin VB.Label Label67 
            Caption         =   "O/S "
            Height          =   255
            Left            =   240
            TabIndex        =   207
            Top             =   735
            Width           =   825
         End
         Begin VB.Label Label59 
            Caption         =   "Bill Amt"
            Height          =   255
            Left            =   240
            TabIndex        =   206
            Top             =   255
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmdTotal 
         Caption         =   "Total"
         Height          =   375
         Left            =   -64440
         TabIndex        =   195
         Top             =   2280
         Width           =   855
      End
      Begin VB.Frame Frame13 
         Caption         =   "Worksheet Summary"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   186
         Top             =   480
         Width           =   5055
         Begin VB.TextBox txtAnnualLmt 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   187
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtAvailableAnnualLmt 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   188
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtApprovedAmt 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   189
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtBordxAmt 
            Height          =   285
            Left            =   1320
            TabIndex        =   190
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label88 
            Caption         =   "Balance Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   194
            Top             =   495
            Width           =   945
         End
         Begin VB.Label Label87 
            Caption         =   "Approved Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   193
            Top             =   735
            Width           =   1065
         End
         Begin VB.Label Label86 
            Caption         =   "Annual Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   192
            Top             =   255
            Width           =   1065
         End
         Begin VB.Label Label89 
            Caption         =   "Bordx Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   191
            Top             =   960
            Width           =   1065
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Payable to Member"
         Height          =   1335
         Left            =   -69360
         TabIndex        =   177
         Top             =   480
         Width           =   4815
         Begin VB.TextBox txtPayableMem 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   178
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtPaid2MBM 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   179
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtPaidByMem 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   180
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtMemOS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   181
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label txtMBMPayable 
            Caption         =   "Payable to/(by) Mem"
            Height          =   255
            Left            =   120
            TabIndex        =   185
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label Label84 
            Caption         =   "Received From Mem"
            Height          =   255
            Left            =   120
            TabIndex        =   184
            Top             =   720
            Width           =   1785
         End
         Begin VB.Label Label82 
            Caption         =   "O/S"
            Height          =   255
            Left            =   120
            TabIndex        =   183
            Top             =   960
            Width           =   1425
         End
         Begin VB.Label Label81 
            Caption         =   "Paid To Mem"
            Height          =   255
            Left            =   120
            TabIndex        =   182
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Others"
         Height          =   1215
         Left            =   -69360
         TabIndex        =   170
         Top             =   1800
         Width           =   4815
         Begin VB.TextBox txtReimbursement 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   171
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtExcess 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   172
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtExcessPaid 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   173
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label80 
            Caption         =   "Total Reimburse"
            Height          =   255
            Left            =   120
            TabIndex        =   176
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label78 
            Caption         =   "Total Excess"
            Height          =   255
            Left            =   120
            TabIndex        =   175
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label92 
            Caption         =   "Excess Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   174
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5010
         Left            =   120
         TabIndex        =   107
         Top             =   360
         Width           =   11445
         Begin VB.TextBox cboPLNCode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   9960
            TabIndex        =   110
            Top             =   210
            Width           =   1200
         End
         Begin VB.TextBox cboINSCode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            TabIndex        =   109
            Top             =   210
            Width           =   1875
         End
         Begin VB.TextBox txtMBMAdd1 
            Height          =   285
            Left            =   1080
            MaxLength       =   40
            TabIndex        =   108
            Top             =   195
            Width           =   4065
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
            Height          =   315
            Left            =   6840
            TabIndex        =   111
            Top             =   480
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMBMBordxDate 
            Height          =   300
            Left            =   6840
            TabIndex        =   112
            Top             =   760
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
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
            Left            =   9960
            TabIndex        =   113
            Top             =   480
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMAdd2 
            Height          =   285
            Left            =   1080
            MaxLength       =   40
            TabIndex        =   114
            Top             =   450
            Width           =   4065
         End
         Begin VB.TextBox txtMBMAdd3 
            Height          =   285
            Left            =   1080
            MaxLength       =   40
            TabIndex        =   115
            Top             =   710
            Width           =   4065
         End
         Begin VB.TextBox txtMBMPostCode 
            Height          =   285
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   116
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox cboCTYCode 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   117
            Top             =   960
            Width           =   1545
         End
         Begin VB.ComboBox cboSTACode 
            Height          =   315
            Left            =   1080
            TabIndex        =   118
            Top             =   1220
            Width           =   4065
         End
         Begin VB.ComboBox cboMBMSex 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   119
            Top             =   1505
            Width           =   1335
         End
         Begin VB.ComboBox cboMBMTakeOver 
            Height          =   315
            Left            =   3600
            TabIndex        =   120
            Top             =   1505
            Width           =   1545
         End
         Begin VB.ComboBox cboMBMMaritalStatus 
            Height          =   315
            Left            =   3600
            TabIndex        =   122
            Top             =   1770
            Width           =   1545
         End
         Begin MSMask.MaskEdBox txtMBMBirthDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   121
            Top             =   1770
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboRACCode 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   123
            Top             =   2030
            Width           =   1335
         End
         Begin VB.ComboBox cboNATCode 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   124
            Top             =   2030
            Width           =   1545
         End
         Begin VB.TextBox cboBRNCode 
            Height          =   315
            Left            =   1080
            TabIndex        =   125
            Top             =   2300
            Width           =   4065
         End
         Begin VB.ComboBox cboMBMRenewal 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   126
            Top             =   2580
            Width           =   1290
         End
         Begin VB.TextBox txtMBMAgentCode 
            Height          =   315
            Left            =   3600
            MaxLength       =   15
            TabIndex        =   127
            Top             =   2580
            Width           =   1545
         End
         Begin VB.TextBox txtMBMDateJoin 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            MaxLength       =   25
            TabIndex        =   128
            Top             =   1020
            Width           =   1875
         End
         Begin VB.TextBox txtMBMPolicyYear 
            Enabled         =   0   'False
            Height          =   315
            Left            =   9960
            MaxLength       =   25
            TabIndex        =   129
            Top             =   1010
            Width           =   1200
         End
         Begin VB.TextBox txtMBMEmployeeNo 
            Height          =   315
            Left            =   6840
            MaxLength       =   15
            TabIndex        =   130
            Top             =   1290
            Width           =   4335
         End
         Begin VB.TextBox txtDeptName 
            Height          =   315
            Left            =   6840
            MaxLength       =   150
            TabIndex        =   131
            Top             =   1580
            Width           =   4335
         End
         Begin VB.TextBox txtMBMGroupCompany 
            Height          =   315
            Left            =   6840
            MaxLength       =   150
            TabIndex        =   132
            Top             =   1860
            Width           =   4335
         End
         Begin VB.TextBox txtAGECode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            TabIndex        =   133
            Top             =   2140
            Width           =   4335
         End
         Begin VB.ComboBox cboSRVCode 
            Height          =   315
            Left            =   6840
            Sorted          =   -1  'True
            TabIndex        =   134
            Top             =   2410
            Width           =   4335
         End
         Begin VB.ComboBox cboMBMRelCode 
            Height          =   315
            Left            =   6840
            TabIndex        =   137
            Top             =   2700
            Width           =   1815
         End
         Begin VB.ComboBox CboInsMode 
            Height          =   315
            Left            =   9840
            TabIndex        =   135
            Top             =   2700
            Width           =   1305
         End
         Begin VB.TextBox txtPLNCode 
            Height          =   285
            Left            =   6840
            MaxLength       =   4
            TabIndex        =   138
            Top             =   2950
            Width           =   1815
         End
         Begin MSMask.MaskEdBox txtNewPolEffDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Left            =   9840
            TabIndex        =   136
            Top             =   2955
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label55 
            Caption         =   "Ins mode"
            Height          =   255
            Left            =   8760
            TabIndex        =   169
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label45 
            Caption         =   "Renewal "
            Height          =   240
            Left            =   120
            TabIndex        =   167
            Top             =   2640
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Dept. Name"
            Height          =   255
            Left            =   5640
            TabIndex        =   166
            Top             =   1680
            Width           =   1545
         End
         Begin VB.Label Label22 
            Caption         =   "Service Provider"
            Height          =   255
            Left            =   5640
            TabIndex        =   165
            Top             =   2480
            Width           =   1485
         End
         Begin VB.Label Label18 
            Caption         =   "Group Company"
            Height          =   255
            Left            =   5640
            TabIndex        =   164
            Top             =   1935
            Width           =   1545
         End
         Begin VB.Label Label19 
            Caption         =   "Employee Num"
            Height          =   255
            Left            =   5640
            TabIndex        =   163
            Top             =   1400
            Width           =   1545
         End
         Begin VB.Label Label10 
            Caption         =   "Date Join"
            Height          =   255
            Left            =   5640
            TabIndex        =   162
            Top             =   1080
            Width           =   780
         End
         Begin VB.Label Label8 
            Caption         =   "Batch Num"
            Height          =   255
            Left            =   8880
            TabIndex        =   161
            Top             =   760
            Width           =   885
         End
         Begin VB.Label Label17 
            Caption         =   "Bordx. Date"
            Height          =   255
            Left            =   5640
            TabIndex        =   160
            Top             =   780
            Width           =   1275
         End
         Begin VB.Label Label14 
            Caption         =   "Agent Code"
            Height          =   255
            Left            =   2520
            TabIndex        =   159
            Top             =   2640
            Width           =   1140
         End
         Begin VB.Label Label20 
            Caption         =   "Ins Type"
            Height          =   255
            Left            =   5640
            TabIndex        =   158
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label21 
            Caption         =   "Plan Code"
            Height          =   255
            Left            =   8880
            TabIndex        =   157
            Top             =   210
            Width           =   825
         End
         Begin VB.Label Label29 
            Caption         =   "Age Band"
            Height          =   255
            Left            =   5640
            TabIndex        =   156
            Top             =   2170
            Width           =   735
         End
         Begin VB.Label Label72 
            Caption         =   "PostCode"
            Height          =   255
            Left            =   120
            TabIndex        =   155
            Top             =   1015
            Width           =   1140
         End
         Begin VB.Label Label69 
            Caption         =   "Birthdate"
            Height          =   255
            Left            =   120
            TabIndex        =   154
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label Label60 
            Caption         =   "State"
            Height          =   255
            Left            =   120
            TabIndex        =   153
            Top             =   1320
            Width           =   465
         End
         Begin VB.Label Label57 
            Caption         =   "City/Town"
            Height          =   255
            Left            =   2520
            TabIndex        =   152
            Top             =   1020
            Width           =   1020
         End
         Begin VB.Label Label54 
            Caption         =   "Address"
            Height          =   255
            Left            =   120
            TabIndex        =   151
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label56 
            Caption         =   "Sex "
            Height          =   255
            Left            =   120
            TabIndex        =   150
            Top             =   1560
            Width           =   825
         End
         Begin VB.Label Label9 
            Caption         =   "Pol Year"
            Height          =   255
            Left            =   8880
            TabIndex        =   149
            Top             =   1080
            Width           =   1050
         End
         Begin VB.Label Label16 
            Caption         =   "Exp Date"
            Height          =   255
            Left            =   8880
            TabIndex        =   148
            Top             =   450
            Width           =   795
         End
         Begin VB.Label Label15 
            Caption         =   "Eff Date"
            Height          =   255
            Left            =   5640
            TabIndex        =   147
            Top             =   510
            Width           =   1275
         End
         Begin VB.Label Label63 
            Caption         =   "TakeOver"
            Height          =   255
            Left            =   2520
            TabIndex        =   146
            Top             =   1620
            Width           =   825
         End
         Begin VB.Label Label12 
            Caption         =   "Branch"
            Height          =   255
            Left            =   120
            TabIndex        =   145
            Top             =   2400
            Width           =   960
         End
         Begin VB.Label Label70 
            Caption         =   "Race"
            Height          =   255
            Left            =   120
            TabIndex        =   144
            Top             =   2140
            Width           =   420
         End
         Begin VB.Label Label71 
            Caption         =   "Nationality"
            Height          =   275
            Left            =   2520
            TabIndex        =   143
            Top             =   2080
            Width           =   960
         End
         Begin VB.Label Label27 
            Caption         =   "Marital Status"
            Height          =   255
            Left            =   2520
            TabIndex        =   142
            Top             =   1860
            Width           =   975
         End
         Begin VB.Label Label30 
            Caption         =   "Supp Type"
            Height          =   255
            Left            =   5640
            TabIndex        =   141
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label38 
            Caption         =   "New Eff Date"
            Height          =   255
            Left            =   8760
            TabIndex        =   140
            Top             =   3000
            Width           =   1005
         End
         Begin VB.Label Label39 
            Caption         =   "New PLN Code"
            Height          =   255
            Left            =   5640
            TabIndex        =   139
            Top             =   3045
            Width           =   1125
         End
         Begin VB.Label txtMBMBatchNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   9960
            TabIndex        =   168
            Top             =   780
            Width           =   1200
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4965
         Left            =   -74760
         TabIndex        =   83
         Top             =   360
         Width           =   11220
         Begin VB.TextBox txtMBMTelNoH 
            Height          =   315
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   87
            Top             =   285
            Width           =   3585
         End
         Begin VB.TextBox txtMBMExclusion 
            Height          =   2385
            Left            =   5760
            MaxLength       =   4000
            MultiLine       =   -1  'True
            TabIndex        =   86
            Top             =   360
            Width           =   5340
         End
         Begin VB.TextBox txtRemarks 
            Height          =   780
            Left            =   240
            MaxLength       =   4000
            MultiLine       =   -1  'True
            TabIndex        =   84
            Top             =   4080
            Width           =   4890
         End
         Begin VB.TextBox txtMBMTelNoM 
            Height          =   315
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   88
            Top             =   555
            Width           =   3585
         End
         Begin VB.TextBox txtMBMTelNoO 
            Height          =   315
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   89
            Top             =   840
            Width           =   3585
         End
         Begin VB.TextBox txtMBMPrevPlan 
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   90
            Top             =   1125
            Width           =   1215
         End
         Begin VB.TextBox txtMBMRBAmt 
            Enabled         =   0   'False
            Height          =   330
            Left            =   3960
            MaxLength       =   8
            TabIndex        =   91
            Top             =   1125
            Width           =   1215
         End
         Begin VB.TextBox txtMBMPrevINSCom 
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   60
            TabIndex        =   92
            Top             =   1395
            Width           =   3615
         End
         Begin VB.TextBox txtMBMPAYRemarks 
            Height          =   495
            Left            =   1560
            MaxLength       =   60
            MultiLine       =   -1  'True
            TabIndex        =   93
            Top             =   1680
            Width           =   3615
         End
         Begin VB.ComboBox cboCOPayRB 
            Height          =   315
            ItemData        =   "frmMBMEnquiry.frx":0134
            Left            =   1560
            List            =   "frmMBMEnquiry.frx":013B
            TabIndex        =   94
            Top             =   2145
            Width           =   3615
         End
         Begin MSMask.MaskEdBox TxtClmNonPayFr 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Left            =   2040
            TabIndex        =   240
            Top             =   2430
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   476
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtClmNonPayTo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Left            =   3840
            TabIndex        =   241
            Top             =   2430
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMAllergic 
            Height          =   900
            Left            =   240
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   85
            Top             =   3000
            Width           =   4935
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            Caption         =   "To"
            Height          =   255
            Left            =   3480
            TabIndex        =   243
            Top             =   2440
            Width           =   285
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            Caption         =   "From"
            Height          =   255
            Left            =   1560
            TabIndex        =   242
            Top             =   2440
            Width           =   405
         End
         Begin VB.Label Label28 
            Caption         =   "Tel Num. (H)"
            Height          =   255
            Left            =   240
            TabIndex        =   106
            Top             =   315
            Width           =   1005
         End
         Begin VB.Label Label32 
            Caption         =   "H/P Num"
            Height          =   255
            Left            =   240
            TabIndex        =   105
            Top             =   585
            Width           =   1005
         End
         Begin VB.Label Label33 
            Caption         =   "Tel Num (O)"
            Height          =   255
            Left            =   240
            TabIndex        =   104
            Top             =   870
            Width           =   1005
         End
         Begin VB.Label Label40 
            Caption         =   "Principal Warranty"
            Height          =   255
            Left            =   5760
            TabIndex        =   103
            Top             =   120
            Width           =   1590
         End
         Begin VB.Label Label41 
            Caption         =   "Allergic"
            Height          =   255
            Left            =   240
            TabIndex        =   102
            Top             =   2760
            Width           =   1590
         End
         Begin VB.Label Label95 
            Caption         =   "P. Pln Code"
            Height          =   255
            Left            =   240
            TabIndex        =   101
            Top             =   1200
            Width           =   1125
         End
         Begin VB.Label Label96 
            Caption         =   "Prev Ins Com"
            Height          =   255
            Left            =   240
            TabIndex        =   100
            Top             =   1455
            Width           =   1140
         End
         Begin VB.Label Label97 
            Caption         =   "R&&B Amount"
            Height          =   255
            Left            =   3000
            TabIndex        =   99
            Top             =   1200
            Width           =   1140
         End
         Begin VB.Label Label101 
            Caption         =   "Payor Remarks"
            Height          =   255
            Left            =   240
            TabIndex        =   98
            Top             =   1755
            Width           =   1095
         End
         Begin VB.Label Label94 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   3880
            Width           =   1590
         End
         Begin VB.Label Label104 
            Caption         =   "Co-Payment RB"
            Height          =   240
            Left            =   240
            TabIndex        =   96
            Top             =   2190
            Width           =   1140
         End
         Begin VB.Label Label13 
            Caption         =   "CLM Non-Pay"
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   2440
            Width           =   2085
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   36
         Top             =   360
         Width           =   11100
         Begin VB.TextBox txtCovExclusion 
            Height          =   810
            Left            =   6000
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   40
            Top             =   2760
            Width           =   4935
         End
         Begin VB.TextBox txtCovAllergic 
            Height          =   800
            Left            =   240
            MaxLength       =   60
            MultiLine       =   -1  'True
            TabIndex        =   39
            Top             =   2760
            Width           =   5175
         End
         Begin VB.ComboBox cboCovSalutation 
            Height          =   315
            Left            =   1320
            TabIndex        =   38
            Top             =   240
            Width           =   1245
         End
         Begin VB.TextBox txtCovName 
            Height          =   330
            Left            =   3915
            MaxLength       =   50
            TabIndex        =   37
            Top             =   195
            Width           =   7020
         End
         Begin VB.TextBox txtCovIcBcPp 
            Height          =   285
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   41
            Top             =   490
            Width           =   4125
         End
         Begin MSMask.MaskEdBox txtCovBirthdate 
            Height          =   285
            Left            =   1320
            TabIndex        =   42
            Top             =   720
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboCovRelationship 
            Height          =   315
            Left            =   1320
            TabIndex        =   43
            Top             =   980
            Width           =   4125
         End
         Begin VB.ComboBox cboCovSex 
            Height          =   315
            Left            =   1320
            TabIndex        =   44
            Top             =   1260
            Width           =   1485
         End
         Begin VB.ComboBox cboCovAdultChild 
            Height          =   315
            Left            =   3960
            TabIndex        =   45
            Top             =   1260
            Width           =   1485
         End
         Begin VB.TextBox txtPrevPLNCode 
            Height          =   285
            Left            =   1320
            TabIndex        =   46
            Top             =   1550
            Width           =   1455
         End
         Begin VB.TextBox txtRBAmt 
            Height          =   330
            Left            =   3960
            MaxLength       =   8
            TabIndex        =   47
            Top             =   1550
            Width           =   1485
         End
         Begin VB.TextBox txtPrevINSCom 
            Height          =   330
            Left            =   1320
            MaxLength       =   60
            TabIndex        =   48
            Top             =   1800
            Width           =   4120
         End
         Begin VB.TextBox txtCOVPayRemarks 
            Height          =   495
            Left            =   1320
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   2100
            Width           =   4120
         End
         Begin VB.ComboBox cboCovBlood 
            Height          =   315
            Left            =   7440
            TabIndex        =   50
            Top             =   500
            Width           =   3495
         End
         Begin VB.ComboBox cboOccupation 
            Height          =   315
            Left            =   7440
            TabIndex        =   51
            Top             =   780
            Width           =   3495
         End
         Begin VB.ComboBox cboWorkNature 
            Height          =   315
            Left            =   7440
            TabIndex        =   52
            Top             =   1070
            Width           =   3495
         End
         Begin VB.ComboBox cboSmoking 
            Height          =   315
            Left            =   7440
            TabIndex        =   53
            Top             =   1350
            Width           =   1215
         End
         Begin VB.ComboBox cboAlcohol 
            Height          =   315
            Left            =   9720
            TabIndex        =   54
            Top             =   1350
            Width           =   1215
         End
         Begin VB.TextBox txtCovAnnualLimit 
            Height          =   315
            Left            =   7440
            TabIndex        =   55
            Top             =   1640
            Width           =   1245
         End
         Begin VB.TextBox txtAvailLimit 
            Height          =   315
            Left            =   9720
            TabIndex        =   56
            Top             =   1640
            Width           =   1215
         End
         Begin VB.TextBox txtSuppPrem 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7440
            TabIndex        =   57
            Top             =   1920
            Width           =   3495
         End
         Begin MSMask.MaskEdBox TxtSuppClmNonPayFr 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   7440
            TabIndex        =   58
            Top             =   2180
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtSuppClmNonPayTo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   9720
            TabIndex        =   59
            Top             =   2180
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label44 
            Caption         =   "Occupation"
            Height          =   255
            Left            =   6000
            TabIndex        =   82
            Top             =   810
            Width           =   1095
         End
         Begin VB.Label Label46 
            Caption         =   "Work Nature"
            Height          =   255
            Left            =   6000
            TabIndex        =   81
            Top             =   1100
            Width           =   1095
         End
         Begin VB.Label Label43 
            Caption         =   "Smoking"
            Height          =   255
            Left            =   6000
            TabIndex        =   80
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label Label62 
            Caption         =   "Alcohol"
            Height          =   240
            Left            =   8880
            TabIndex        =   79
            Top             =   1380
            Width           =   855
         End
         Begin VB.Label Label42 
            Caption         =   "Mem Name"
            Height          =   255
            Left            =   2640
            TabIndex        =   78
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Balance Limit"
            Height          =   255
            Left            =   8760
            TabIndex        =   77
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Supp Limit"
            Height          =   255
            Left            =   6000
            TabIndex        =   76
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label53 
            Caption         =   "Blood"
            Height          =   255
            Left            =   6000
            TabIndex        =   75
            Top             =   530
            Width           =   1005
         End
         Begin VB.Label Label52 
            Caption         =   "R'ship"
            Height          =   255
            Left            =   240
            TabIndex        =   74
            Top             =   1050
            Width           =   1005
         End
         Begin VB.Label Label47 
            Caption         =   "Supplementary Warranty "
            Height          =   255
            Left            =   6000
            TabIndex        =   73
            Top             =   2520
            Width           =   2580
         End
         Begin VB.Label Label48 
            Caption         =   "Allergic"
            Height          =   255
            Left            =   240
            TabIndex        =   72
            Top             =   2520
            Width           =   600
         End
         Begin VB.Label Label66 
            Caption         =   "Salutation"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label65 
            Caption         =   "IC Num"
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   520
            Width           =   1125
         End
         Begin VB.Label Label49 
            Caption         =   "Birthdate"
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   785
            Width           =   1005
         End
         Begin VB.Label Label50 
            Caption         =   "Sex (M/F)"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   1340
            Width           =   825
         End
         Begin VB.Label Label51 
            Caption         =   "Adult/Child"
            Height          =   255
            Left            =   2880
            TabIndex        =   67
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label Label91 
            Caption         =   "Supp. Prem"
            Height          =   255
            Left            =   6000
            TabIndex        =   66
            Top             =   1960
            Width           =   1005
         End
         Begin VB.Label Label93 
            Caption         =   "P. Plan Code"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   1560
            Width           =   1125
         End
         Begin VB.Label Label98 
            Caption         =   "P. Ins Com"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   1800
            Width           =   1140
         End
         Begin VB.Label Label99 
            Caption         =   "R&&B Amount"
            Height          =   255
            Left            =   2880
            TabIndex        =   63
            Top             =   1560
            Width           =   1140
         End
         Begin VB.Label Label100 
            Caption         =   "Payor Remarks"
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label34 
            Caption         =   "CLM Non-Pay Date"
            Height          =   255
            Left            =   6000
            TabIndex        =   61
            Top             =   2205
            Width           =   1605
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Caption         =   "To"
            Height          =   255
            Left            =   8880
            TabIndex        =   60
            Top             =   2280
            Width           =   645
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1275
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   8625
         Begin VB.Line Line1 
            X1              =   120
            X2              =   8520
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   8520
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label txtMBMProviderRedCharges 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Left            =   7560
            TabIndex        =   35
            Top             =   555
            Width           =   945
         End
         Begin VB.Label txtMBMProviderEffCharges 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Left            =   7560
            TabIndex        =   34
            Top             =   915
            Width           =   945
         End
         Begin VB.Label Label77 
            Caption         =   "Effective Access Fee"
            Height          =   255
            Left            =   5880
            TabIndex        =   33
            Top             =   915
            Width           =   1620
         End
         Begin VB.Label Label79 
            Caption         =   "Acces Fee Refund"
            Height          =   255
            Left            =   5880
            TabIndex        =   32
            Top             =   555
            Width           =   1380
         End
         Begin VB.Label Label61 
            Caption         =   "MCO Refund"
            Height          =   255
            Left            =   2955
            TabIndex        =   31
            Top             =   555
            Width           =   975
         End
         Begin VB.Label txtMBMMCORed 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Left            =   4440
            TabIndex        =   30
            Top             =   555
            Width           =   945
         End
         Begin VB.Label txtMBMPremiumRed 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            TabIndex        =   29
            Top             =   560
            Width           =   1035
         End
         Begin VB.Label Label58 
            Caption         =   "Prem Refund"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   560
            Width           =   1020
         End
         Begin VB.Label txtMBMPremium 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            TabIndex        =   27
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label lblPremium 
            Caption         =   "Premium"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   195
            Width           =   960
         End
         Begin VB.Label txtMBMProviderCharges 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Left            =   7560
            TabIndex        =   25
            Top             =   180
            Width           =   945
         End
         Begin VB.Label txtMCOFeesEff 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Left            =   4440
            TabIndex        =   24
            Top             =   915
            Width           =   945
         End
         Begin VB.Label txtMBMPremiumEff 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            TabIndex        =   23
            Top             =   915
            Width           =   1035
         End
         Begin VB.Label txtMCOFees 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Left            =   4440
            TabIndex        =   22
            Top             =   180
            Width           =   945
         End
         Begin VB.Label Label83 
            Caption         =   "Effective Prem "
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   915
            Width           =   1140
         End
         Begin VB.Label Label85 
            Caption         =   "Effective MCO Fee "
            Height          =   255
            Left            =   2955
            TabIndex        =   20
            Top             =   915
            Width           =   1455
         End
         Begin VB.Label lblMCOFee 
            Caption         =   "MCO Fees"
            Height          =   255
            Left            =   2955
            TabIndex        =   19
            Top             =   195
            Width           =   870
         End
         Begin VB.Label Label11 
            Caption         =   "Access Fee"
            Height          =   255
            Left            =   5880
            TabIndex        =   18
            Top             =   180
            Width           =   1140
         End
      End
      Begin MSComctlLib.ListView lstMBMHis 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   196
         Top             =   720
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8070
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
            Text            =   "Mem No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pol No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Eff Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Exp Date"
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
            Text            =   "Pol Status"
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
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "MBM Reg Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstCovAdd 
         Height          =   1425
         Left            =   -74640
         TabIndex        =   197
         Top             =   3960
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   2514
         SortKey         =   9
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
         NumItems        =   25
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Salutation"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "IC No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Birthdate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sex"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Adult/Child"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Relationship"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Blood"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Allergic"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Exclusion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Status"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "AnnualLimit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Available Limit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Occupation"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Work Nature"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Smoking"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Alcohol"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Premium"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Prev Plan"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "RB Amt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Payor Remarks"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Prev Insurance Com"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Clm Non Pay Fr"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Clm Non Pay To"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstCaseList 
         Height          =   4995
         Left            =   -74880
         TabIndex        =   198
         Top             =   360
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   8811
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Mem No"
            Object.Width           =   2117
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
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Claimability"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Treatment Reason"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Case Status"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView lstAdjustmentHis 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   199
         Top             =   360
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   8705
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
            Text            =   "Bordx Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Batch No"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Adjust. Date"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Endt. Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Endt. Eff Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstBenefitDetails 
         Height          =   4755
         Left            =   -74940
         TabIndex        =   200
         Top             =   600
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   8387
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "BNF Desc"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cal."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "RM PDay"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "No PAnn"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "RM PDay"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "%"
            Object.Width           =   970
         EndProperty
      End
      Begin MSComctlLib.ListView lstMBMAcc 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   201
         Top             =   3120
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   3836
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MBMNumber"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Cover ID"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Approved Amt"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Pay2CLN"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Paid2CLN"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "PAY2MBM"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Paid2MBM"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "PaidByMBM"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Reimburse"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Excess"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Bordx Amt"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Excess Paid"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView lstLimitation 
         Height          =   4905
         Left            =   -74880
         TabIndex        =   209
         Top             =   360
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   8652
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Relation"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Gen WP"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "End of Gen WP"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Pre WP"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "End of Pre WP"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Spec Illness WP"
            Object.Width           =   2910
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "End of Spec Illness WP"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Staff Disclosure Req"
            Object.Width           =   2822
         EndProperty
      End
      Begin MSComctlLib.ListView lstSurgicalProc 
         Height          =   4995
         Left            =   -74880
         TabIndex        =   210
         Top             =   360
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   8811
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sur Proc Cat"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Sur Proc Cat Desc"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Sur Proc Code"
            Object.Width           =   2294
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Sur Proc Code Desc"
            Object.Width           =   9879
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Percentage"
            Object.Width           =   1941
         EndProperty
      End
      Begin MSComctlLib.ListView lstNewBenefitDetails 
         Height          =   4755
         Left            =   -67965
         TabIndex        =   211
         Top             =   600
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   8387
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "BNF Desc"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cal."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "RM PFee"
            Object.Width           =   1587
         EndProperty
      End
      Begin VB.Label Label26 
         Caption         =   "Double Click to Open Old Membership data (Members With Same IC are Listed Below)"
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
         Left            =   -74880
         TabIndex        =   214
         Top             =   480
         Width           =   8055
      End
      Begin VB.Label Label37 
         Caption         =   "Member BNF Listing"
         Height          =   255
         Left            =   -74880
         TabIndex        =   213
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label105 
         Caption         =   "Clinic BNF Listing"
         Height          =   255
         Left            =   -67920
         TabIndex        =   212
         Top             =   360
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport crptSummary 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame10 
      Height          =   975
      Left            =   120
      TabIndex        =   223
      Top             =   1080
      Width           =   11655
      Begin VB.ComboBox cboMBMSalutation 
         Height          =   315
         Left            =   1200
         TabIndex        =   227
         Top             =   195
         Width           =   915
      End
      Begin VB.TextBox txtMBMName 
         Height          =   330
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   226
         Top             =   187
         Width           =   3435
      End
      Begin VB.ComboBox cboPAYCode 
         Height          =   315
         Left            =   1200
         TabIndex        =   225
         Top             =   500
         Width           =   915
      End
      Begin VB.TextBox txtMBMPolNo 
         Height          =   330
         Left            =   3120
         MaxLength       =   25
         TabIndex        =   224
         Top             =   500
         Width           =   3435
      End
      Begin VB.TextBox txtMBMIcBcPP 
         Height          =   285
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   228
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblVisitLmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   8520
         TabIndex        =   244
         Top             =   165
         Width           =   1080
      End
      Begin VB.Label cboHLTCode 
         BackColor       =   &H80000009&
         Height          =   285
         Left            =   7320
         TabIndex        =   230
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label24 
         Caption         =   "Salutation"
         Height          =   255
         Left            =   240
         TabIndex        =   239
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label64 
         Caption         =   "IC Num"
         Height          =   330
         Left            =   6600
         TabIndex        =   238
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label31 
         Caption         =   "Mem Name"
         Height          =   255
         Left            =   2160
         TabIndex        =   237
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label73 
         Caption         =   "Policy Num"
         Height          =   255
         Left            =   2160
         TabIndex        =   236
         Top             =   525
         Width           =   900
      End
      Begin VB.Label Label74 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   235
         Top             =   520
         Width           =   915
      End
      Begin VB.Label Label75 
         Caption         =   "Hlt Type"
         Height          =   255
         Left            =   6600
         TabIndex        =   234
         Top             =   165
         Width           =   795
      End
      Begin VB.Label txtAnnualLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   10440
         TabIndex        =   231
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label txtAvailableAnnualLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   10440
         TabIndex        =   229
         Top             =   495
         Width           =   1080
      End
      Begin VB.Label Label5 
         Caption         =   "Ann Lmt"
         Height          =   255
         Left            =   9795
         TabIndex        =   233
         Top             =   165
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Bal Lmt"
         Height          =   255
         Left            =   9795
         TabIndex        =   232
         Top             =   480
         Width           =   885
      End
      Begin VB.Label Label102 
         Caption         =   "Visit Lmt"
         Height          =   255
         Left            =   7880
         TabIndex        =   245
         Top             =   165
         Width           =   870
      End
   End
   Begin VB.Frame Frame7 
      Height          =   525
      Left            =   120
      TabIndex        =   215
      Top             =   7440
      Width           =   11655
      Begin VB.TextBox txtFirstMeMNo 
         Height          =   285
         Left            =   4080
         TabIndex        =   222
         Text            =   "FirstMeMNo"
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   10200
         TabIndex        =   221
         Top             =   150
         Width           =   960
      End
      Begin VB.TextBox txtMBMAdultChild 
         Height          =   285
         Left            =   840
         TabIndex        =   220
         Tag             =   "T"
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtLastEndStatus 
         Height          =   285
         Left            =   1680
         TabIndex        =   219
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtLastEdtDate 
         Height          =   285
         Left            =   2760
         TabIndex        =   218
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   9240
         TabIndex        =   217
         Top             =   150
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox txtcount 
         Height          =   285
         Left            =   360
         TabIndex        =   216
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmMBMEnquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bExit As Boolean
Public bDataStatus As Boolean
Public ReportCriteria As String
Dim ScanCode As String
Private m_blnDirection(1 To 9) As Boolean
Dim ALmtInd As String

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
        
    If txtMBMNumber <> "" Then
        
        crptSummary.ReportFileName = crdPath & "\Report\RptMbmEnquiry.rpt"
        crptSummary.Destination = crptToWindow
        crptSummary.WindowState = crptMaximized
        crptSummary.SelectionFormula = "{VIEW_MBM_ENQUIRY.MBMNumber} = '" & Trim(txtMBMNumber) & "'"
        crptSummary.ParameterFields(0) = "MemberNo;" & Trim(txtMBMNumber) & ";true"
        crptSummary.Action = 1
        
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
              
    Screen.MousePointer = Default
End Sub

Private Sub CmdTotal_Click()
Dim Acc As New ADODB.Recordset
Dim strsql As String
Dim masterlist As ListItem
Dim strsqlMBM As String
Dim MBM As New ADODB.Recordset
Dim strsqlMBMArc As String
Dim MBMArc As New ADODB.Recordset
Dim TotalApprove As Variant
Dim totalPayable2Clinic As Variant
Dim totalPaid2Clinic As Variant
Dim totalPayable2Mbm As Variant
Dim totalPaidByMbm As Variant
Dim totalPaid2MBM As Variant
Dim TotalExcess As Variant
Dim TotalReimburse As Variant
Dim MBMTotalAnnualLmt As Variant
Dim MBMTotalAvailLmt As Variant

    If lstMBMAcc.ListItems.Count <> 0 Then
        strsqlMBM = " Select MBMNumber, MBMMemberType, MBMCAnnualLimit, " & _
                    " MBMCAvailableLimit " & _
                    " From View_MBM_MemberType where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    If MBM.RecordCount Then
        If MBM!MBMMemberType = "P" Then
            txtAnnualLmt = txtAnnualLimit
            txtAvailableAnnualLmt = txtAvailableAnnualLimit
        Else
            Do Until MBM.EOF
                MBMTotalAnnualLmt = Format(MBMTotalAnnualLmt + MBM!MBMCAnnualLimit, "#,##0.00")
                MBMTotalAvailLmt = Format(MBMTotalAvailLmt + MBM!MBMCAvailableLimit, "#,##0.00")
            MBM.MoveNext
            Loop
                txtAnnualLmt = Format(MBMTotalAnnualLmt + txtAnnualLimit, "#,##0.00")
                txtAvailableAnnualLmt = Format(MBMTotalAvailLmt + txtAvailableAnnualLimit, "#,##0.00")
        End If
        MBM.Close
        Set MBM = Nothing
    Else
    
        strsqlMBMArc = " Select MBMNumber, MBMMemberType, MBMCAnnualLimit, " & _
                       " MBMCAvailableLimit, From View_MBM_MemTypeArc where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        MBMArc.Open strsqlMBMArc, medixmasterconn, adOpenKeyset, adLockOptimistic

        If MBMArc.RecordCount Then
            If MBMArc!MBMMemberType = "P" Then
                txtAnnualLmt = Format(txtAnnualLimit, "#,##0.00")
                txtAvailableAnnualLmt = Format(txtAvailableAnnualLimit, "#,##0.00")
            Else
                Do Until MBM.EOF
                    MBMTotalAnnualLmt = Format(MBMTotalAnnualLmt + MBMArc!MBMCAnnualLimit, "#,##0.00")
                    MBMTotalAvailLmt = Format(MBMTotalAvailLmt + MBMArc!MBMCAvailableLimit, "#,##0.00")
                MBMArc.MoveNext
                Loop
                    txtAnnualLmt = Format(MBMTotalAnnualLmt + txtAnnualLimit, "#,##0.00")
                    txtAvailableAnnualLmt = Format(MBMTotalAvailLmt + txtAvailableAnnualLimit, "#,##0.00")
            End If
        End If
        MBMArc.Close
        Set MBMArc = Nothing
    End If

        strsql = " select totalApprove, totalPayable2Clinic, " & _
                " totalPaid2Clinic, totalPayable2Mbm, totalPaidByMbm, CoverID, " & _
                " totalPaid2MBM , totalExcess , totalReimburse from view_CISMBMTotal_summary where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        Acc.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If Not Acc.EOF Then
            Do Until Acc.EOF
            With Acc
                TotalApprove = Format(TotalApprove + !TotalApprove, "#,##0.00")
                totalPayable2Clinic = Format(totalPayable2Clinic + !totalPayable2Clinic, "#,##0.00")
                If Trim(!totalPaid2Clinic) <> "" Then
                    totalPaid2Clinic = Format(totalPaid2Clinic + !totalPaid2Clinic, "#,##0.00")
                End If
                totalPayable2Mbm = Format(totalPayable2Mbm + !totalPayable2Mbm, "#,##0.00")
                If Trim(!totalPaid2MBM) <> "" Then
                    totalPaid2MBM = Format(totalPaid2MBM + totalPaid2MBM, "#,##0.00")
                End If
                If Trim(!totalPaidByMbm) <> "" Then
                    totalPaidByMbm = Format(totalPaidByMbm + totalPaidByMbm, "#,##0.00")
                End If
                TotalReimburse = Format(TotalReimburse + !TotalReimburse, "#,##0.00")
                TotalExcess = Format(TotalExcess + !TotalExcess, "#,##0.00")
            End With
            Acc.MoveNext
            Loop
            
            txtApprovedAmt = Format(TotalApprove, "#,##0.00")
            
            txtPayable2Hos = totalPayable2Clinic
            If IsNumeric(txtPayable2Hos) = False Then
                txtPayable2Hos = "0.00"
            End If
            
            txtPaid2Hos = totalPaid2Clinic
            If IsNumeric(txtPaid2Hos) = False Then
                txtPaid2Hos = "0.00"
            End If
            ' Out standing hos
            txtHosOS = Format(CDbl(txtPayable2Hos) - CDbl(txtPaid2Hos), "#,##0.00;(#,##0.00)")
            
            txtPayableMem = Format(totalPayable2Mbm, "#,##0.00;(#,##0.00)")
            
            txtPaid2MBM = Format(totalPaid2MBM, "#,##0.00")
            If IsNumeric(txtPaid2MBM) = False Then
                txtPaid2MBM = "0.00"
            End If
                
            txtPaidByMem = Format(totalPaidByMbm, "#,##0.00")
            If IsNumeric(txtPaidByMem) = False Then
                txtPaidByMem = "0.00"
            End If
            ' Out standing MBM
            txtMemOS = Format(CDbl(txtPayableMem) - CDbl(txtPaid2MBM) + CDbl(txtPaidByMem), "#,#.00;(#,#.00)")
            txtReimbursement = Format(TotalReimburse, "#,##0.00;(#,##0.00)")
            txtExcess = Format(TotalExcess, "#,##0.00;(#,##0.00)")
        End If
        Acc.Close
        Set Acc = Nothing
    End If
    'End If
    
End Sub

Private Sub Form_Load()
    Call ComboListing
    bDataStatus = True
    cmdShow.Enabled = True
    txtMBMNumber.Enabled = True
    cmdSearchMBM.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub


Private Sub ClearEntries()
    On Error Resume Next
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            'If ctl.Enabled = True Then
            If ctl.Name <> "txtMBMNumber" Then
               ctl.Enabled = True
               ctl.Text = ""
               'ctl.Enabled = False
            End If
            'End If
        ElseIf TypeOf ctl Is ComboBox Then
            'If ctl.Enabled = True Then
                ctl.Enabled = True
                ctl.Text = ""
                'ctl.Enabled = False
            'End If
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = True
            
        ElseIf TypeOf ctl Is MaskEdBox Then
            'If ctl.Enabled = True Then
                ctl.Enabled = True
                ctl.mask = ""
                ctl.Text = ""
                ctl.mask = "##/##/####"
                'ctl.Enabled = False
            'End If
        End If
    Next
    'Check1.Value = 0
End Sub

Private Sub ComboListing()
    cboMBMTakeOver.AddItem "N - None"
    cboMBMTakeOver.AddItem "Y - Take Over"
    cboMBMTakeOver.AddItem "C - Conversion"
    
    CboInsMode.AddItem "A - Annually"
    CboInsMode.AddItem "Q - Quarterly"
    CboInsMode.AddItem "S - Semi Annual"
    CboInsMode.AddItem "M - Monthly"
    
End Sub

Private Sub CISBNFListing()
    Dim rst2 As New ADODB.Recordset
    Dim BNF As ListItem
    Dim sql As String
    If lstBenefitDetails.ListItems.Count = 0 And Len(Trim(cboHLTCode)) <> 0 _
     And Len(Trim(cboINSCode)) <> 0 And Len(Trim(cboPLNCode)) <> 0 Then
        sql = " Select BNFCode, BNFNoOfDays , BNFPercentage, BNFNoOfAnn, " & _
              " BNFCATDescription, BNFCalculation, BNFRMPerVisit " & _
              " from View_CISBNF where PLNCode ='" & Trim(cboPLNCode) & "'" & _
              " AND HLTCode ='" & Trim(cboHLTCode) & "'" & _
              " AND INSCODe ='" & Trim(cboINSCode) & "'"
        rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        Do Until rst2.EOF
            Set BNF = lstBenefitDetails.ListItems.Add(, , rst2!BNFCode)
                If Trim(rst2!BNFCATDescription) <> "" Then
                    BNF.SubItems(1) = Trim(rst2!BNFCATDescription)
                End If
                
                BNF.SubItems(2) = IIf(IsNull(rst2!BNFCalculation), 0, rst2!BNFCalculation)
                If Trim(rst2!BNFNoOfDays) <> "" Then
                    BNF.SubItems(3) = IIf(IsNull(rst2!BNFNoOfDays), 0, rst2!BNFNoOfDays)
                End If
                If Trim(rst2!BNFNoOfAnn) <> "" Then
                    BNF.SubItems(4) = IIf(IsNull(rst2!BNFNoOfAnn), 0, rst2!BNFNoOfAnn)
                End If
                If Trim(rst2!BNFRMPerVisit) <> "" Then
                    BNF.SubItems(5) = Format(rst2!BNFRMPerVisit, "#,##,0.00")
                End If
                If Trim(rst2!BNFPercentage) <> "" Then
                    BNF.SubItems(6) = rst2!BNFPercentage
                End If
                
                rst2.MoveNext
        Loop
       
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub

Private Sub ClinicBNFlisting()
    Dim rst2 As New ADODB.Recordset
    Dim BNF As ListItem
    Dim sql As String
    If lstNewBenefitDetails.ListItems.Count = 0 And Len(Trim(cboHLTCode)) <> 0 _
     And Len(Trim(cboINSCode)) <> 0 And Len(Trim(cboPLNCode)) <> 0 Then
        sql = " Select BNFCode, BNFCATDescription, " & _
              " BNFCalculation, BNFRMPerVisit " & _
              " from View_CLNBNF where PLNCode ='" & Trim(cboPLNCode) & "'" & _
              " AND HLTCode ='" & Trim(cboHLTCode) & "'" & _
              " AND INSCODe ='" & Trim(cboINSCode) & "'"
        rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        Do Until rst2.EOF
            Set BNF = lstNewBenefitDetails.ListItems.Add(, , rst2!BNFCode)
                If Trim(rst2!BNFCATDescription) <> "" Then
                    BNF.SubItems(1) = Trim(rst2!BNFCATDescription)
                End If
                
                BNF.SubItems(2) = IIf(IsNull(rst2!BNFCalculation), 0, rst2!BNFCalculation)
                If Trim(rst2!BNFRMPerVisit) <> "" Then
                    BNF.SubItems(3) = Format(rst2!BNFRMPerVisit, "#,##,0.00")
                End If
                
                rst2.MoveNext
        Loop
       
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub

Private Sub MBMHISListing()
    Dim MBMHIS As New ADODB.Recordset
    Dim HISList As ListItem
    Dim sql As String
    If Len(Trim(txtMBMIcBcPP)) <> 0 Then
        If lstMBMHis.ListItems.Count = 0 Then
            sql = " select MBMNumber, MBMPolicyNo, " & _
                " MBMPayorEffDate, MBMPayorExpDate,MBMENDtBordxDate, " & _
                " INSCode, PLNCode, MBMStatus, MBMENDTEffDate, MBMDataStatus, MBMBordxDate, MBMValueDate from View_MBM_MBMOthers " & _
                " where MBMIcBcPp ='" & Trim(txtMBMIcBcPP) & "'"
            
            MBMHIS.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
            Do Until MBMHIS.EOF
                Set HISList = lstMBMHis.ListItems.Add(, , MBMHIS!MBMNumber)
                    With MBMHIS
                        HISList.SubItems(1) = !MBMPolicyNo
                        If Trim(!MBMPayorEffDate) <> "" Then
                            HISList.SubItems(2) = !MBMPayorEffDate
                        End If
                        If Trim(!MBMPayorEXPDate) <> "" Then
                            HISList.SubItems(3) = !MBMPayorEXPDate
                        End If
                        If !MBMStatus = "C" Then
                            If Trim(!MBMENDtEffDate) <> "" Then
                                HISList.SubItems(4) = !MBMENDtEffDate
                            End If
                        Else
                            HISList.SubItems(4) = ""
                        End If
                        If Len(!MBMDataStatus) Then
                            HISList.SubItems(5) = !MBMDataStatus
                        End If
                        If Len(!MBMStatus) Then
                            HISList.SubItems(6) = !MBMStatus
                        End If
                        If Len(!MBMBordxDate) Then
                            HISList.SubItems(7) = !MBMBordxDate
                        End If
                        If Len(!MBMENDtBordxDate) Then
                            HISList.SubItems(8) = !MBMENDtBordxDate
                        End If
                        If Len(!MBMValueDate) Then
                            HISList.SubItems(9) = !MBMValueDate
                        End If
                        
                        .MoveNext
                    End With
            Loop
           
            MBMHIS.Close
            Set MBMHIS = Nothing
        End If
    End If
End Sub


Private Sub ADJListing()
    Dim ADJ As New ADODB.Recordset
    Dim ADJList As ListItem
    Dim sql As String
    
    If Len(Trim(txtMBMNumber)) Then
         sql = " select MBMAmendID, MBMNumber, " & _
           " MBMAmendEdtType, MBMAmendDate, MBMAENDtEffDate, " & _
           " MBMAEndtBordxDate, MBMAEndtBatchNo " & _
           " From MBMAmendHistory where MBMnumber = '" & Trim(txtMBMNumber) & "'"
         ADJ.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
         
         Do Until ADJ.EOF
         
             Set ADJList = lstAdjustmentHis.ListItems.Add(, , IIf(Len(Trim(ADJ!MBMAEndtBordxDate)) <> 0, ADJ!MBMAEndtBordxDate, "-"))
                 If Len(Trim(ADJ!MBMAEndtBatchNo)) Then
                    ADJList.SubItems(1) = ADJ!MBMAEndtBatchNo
                 End If
                 If Len(Trim(ADJ!MBMAmendDate)) Then
                    ADJList.SubItems(2) = ADJ!MBMAmendDate
                 End If

                If Len(ADJ!MBMAmendEdtType) Then
                    ADJList.SubItems(3) = ADJ!MBMAmendEdtType
                End If
                
                If Len(ADJ!MBMAEndtEffDate) Then
                    ADJList.SubItems(4) = ADJ!MBMAEndtEffDate
                End If

               
                 ADJ.MoveNext
         Loop
         ADJ.Close
         Set ADJ = Nothing
    End If
End Sub

Private Sub MBMLimitationListing()
Dim MBMLimit As New ADODB.Recordset
Dim MBMCovLimit As New ADODB.Recordset
Dim LimitList As ListItem
Dim sql As String
Dim sql2 As String
    
    lstLimitation.ListItems.Clear
    
    If Len(Trim(txtMBMNumber)) Then
         sql = " select MBMNumber, MBMName, MBMGenWP, MBMPreExWP, " & _
           " MBMSpeIllWP, MBMStaffDec, MBMGenEndWP, " & _
           " MBMPreExEndWP, MBMSpeIllEndWP, MBMCOVID " & _
           " From VIEW_MBM_Limitation where MBMNumber = '" & Trim(txtMBMNumber) & "'"
         MBMLimit.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
         
         Do Until MBMLimit.EOF
         
             Set LimitList = lstLimitation.ListItems.Add(, , Trim(MBMLimit!MBMName))
                 If Len(Trim(MBMLimit!MBMCOVID)) Then
                    LimitList.SubItems(1) = MBMLimit!MBMCOVID
                 End If
                 If Len(Trim(MBMLimit!MBMGenWP)) Then
                    LimitList.SubItems(2) = MBMLimit!MBMGenWP
                 End If

                If Len(MBMLimit!MBMGenEndWP) Then
                    LimitList.SubItems(3) = MBMLimit!MBMGenEndWP
                End If
                
                If Len(MBMLimit!MBMPreExWP) Then
                    LimitList.SubItems(4) = MBMLimit!MBMPreExWP
                End If
                
                If Len(MBMLimit!MBMPreExEndWP) Then
                    LimitList.SubItems(5) = MBMLimit!MBMPreExEndWP
                End If
                
                If Len(MBMLimit!MBMSpeIllWP) Then
                    LimitList.SubItems(6) = MBMLimit!MBMSpeIllWP
                End If
                
                If Len(MBMLimit!MBMSpeIllEndWP) Then
                    LimitList.SubItems(7) = MBMLimit!MBMSpeIllEndWP
                End If
                
                If Len(MBMLimit!MBMStaffDec) Then
                    LimitList.SubItems(8) = MBMLimit!MBMStaffDec
                End If

               
                 MBMLimit.MoveNext
         Loop
         MBMLimit.Close
         Set MBMLimit = Nothing
         
         sql2 = " select MBMCNumber, MBMCName, MBMCGenWP, MBMCPreExWP, " & _
                " MBMCSpeIllWP, MBMCStaffDec, MBMCGenEndWP, " & _
                " MBMCPreExEndWP, MBMSpeIllEndWP, MBMCCoverID " & _
                " From VIEW_MBMCoveredPersons_Limitation where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
         MBMCovLimit.Open sql2, medixmasterconn, adOpenKeyset, adLockOptimistic
         
         Do Until MBMCovLimit.EOF
         
             Set LimitList = lstLimitation.ListItems.Add(, , Trim(MBMCovLimit!MBMCName))
                 If Len(Trim(MBMCovLimit!MBMCCoverID)) Then
                    LimitList.SubItems(1) = MBMCovLimit!MBMCCoverID
                 End If
                 If Len(Trim(MBMCovLimit!MBMCGenWP)) Then
                    LimitList.SubItems(2) = MBMCovLimit!MBMCGenWP
                 End If

                If Len(MBMCovLimit!MBMCGenEndWP) Then
                    LimitList.SubItems(3) = MBMCovLimit!MBMCGenEndWP
                End If
                
                If Len(MBMCovLimit!MBMCPreExWP) Then
                    LimitList.SubItems(4) = MBMCovLimit!MBMCPreExWP
                End If
                
                If Len(MBMCovLimit!MBMCPreExEndWP) Then
                    LimitList.SubItems(5) = MBMCovLimit!MBMCPreExEndWP
                End If
                
                If Len(MBMCovLimit!MBMCSpeIllWP) Then
                    LimitList.SubItems(6) = MBMCovLimit!MBMCSpeIllWP
                End If
                
                If Len(MBMCovLimit!MBMSpeIllEndWP) Then
                    LimitList.SubItems(7) = MBMCovLimit!MBMSpeIllEndWP
                End If
                
                If Len(MBMCovLimit!MBMCStaffDec) Then
                    LimitList.SubItems(8) = MBMCovLimit!MBMCStaffDec
                End If
                MBMCovLimit.MoveNext
         Loop
         MBMCovLimit.Close
         Set MBMCovLimit = Nothing
         
    End If
End Sub

' ######## List out all the cases which have been discharged
Private Sub ListCases()
Dim CASList As ListItem
Dim strsql As String
Dim CAS As New ADODB.Recordset

    strsql = " Select CASNumber, MBMCName, " & _
             " ReasonDesc, CoverID, " & _
             " DateVisit, MBMCRELCode, " & _
             " CASStatus, MBMName, MBMNumber, Claimability  "
    strsql = strsql & " From VIEW_CISMBMInquiry_CASE where 0=0"
    strsql = strsql & " AND MBMNumber = '" & RemStr(txtMBMNumber) & "' "
            
    CAS.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If CAS.RecordCount <> 0 Then
        Screen.MousePointer = vbHourglass
        lstCaseList.ListItems.Clear
        Do Until CAS.EOF
            Set CASList = lstCaseList.ListItems.Add(, , CAS!CASNumber)
                CASList.SubItems(1) = CAS!MBMNumber
                
                If CAS!CoverID = "00" Then
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
                
                CAS.MoveNext
        Loop
        Screen.MousePointer = vbDefault
    End If
    CAS.Close
    Set CAS = Nothing
End Sub


' ####################### Calculation's Function ###############
Private Sub MemberCallAge()
'On Error Resume Next
    Dim SelectAge As String
    If Trim(txtMBMPayorEffDate) <> "__/__/____" _
        And Len(Trim(txtMBMBirthDate)) <> 0 _
        And Trim(txtMBMBirthDate) <> "__/__/____" _
        And Len(Trim(cboHLTCode)) <> 0 Then
        If IsDate(txtMBMBirthDate) Then
            SelectAge = GetAge(txtMBMPayorEffDate, txtMBMBirthDate)
            txtAGECode = GetAgeBand(SelectAge)
            If Len(Trim(txtAGECode)) = 0 Then
                MsgBox "Please Enter A Valid Date OR Age Out of Acceptable Range", vbOKOnly + vbCritical, DialogBoxTitle
                txtMBMPayorEffDate.SetFocus
            Else
                If Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)) = "01" Then
                    txtMBMAdultChild = "C"
                Else
                    txtMBMAdultChild = "A"
                End If
            End If
        End If
    End If
End Sub

Private Sub MemberCallFees()
    txtMBMPremium = ""
    txtMCOFees = ""
    If InStr(1, cboPLNCode, "-") > 1 Then
        txtMBMPremium = Format(GetPremium(Trim(Mid(cboINSCode, 1, 1)), Trim(Mid(cboPAYCode, 1, InStr(1, cboPAYCode, "-") - 1)), Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)), Trim(Mid(cboPLNCode, 1, InStr(1, cboPLNCode, "-") - 1)), Trim(Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1)), CDate(txtMBMPayorEffDate)), "#,##0.00")
        txtMCOFees = Format(GetMCO(Trim(Mid(cboINSCode, 1, 1)), Trim(Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1)), txtMBMAdultChild, CDate(txtMBMPayorEffDate), Trim(cboPLNCode)), "#,##0.00")
    End If
End Sub

Private Sub ProviderCallFees()
'On Error Resume Next
    Dim GetProString As String
    
    If Len(Trim(cboINSCode)) <> 0 And Len(Trim(cboSRVCode)) <> 0 And Len(Trim(cboHLTCode)) <> 0 Then
        txtMBMProviderCharges = ""
       
        GetProString = GetProviderCharges(Trim(cboSRVCode), Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild)
        If Len(Trim(GetProString)) Then
            txtMBMProviderCharges = Format(GetProString, "###,##0.00")
        End If
    
    Else
        txtMBMProviderCharges = 0
    End If
            
End Sub

Private Sub CallLimit(ALInd As String, MBMtype As String)
'On Error Resume Next
    txtAnnualLimit = ""
    If Trim(ALInd) = "Y" Then
        If Len(Trim(cboPLNCode)) Then
            txtAnnualLimit = Format(GetAnnualLimit(cboINSCode, cboPLNCode, cboHLTCode, txtMBMPayorEffDate, MBMtype), "#,##0.00")
        End If
    Else
        txtAnnualLimit = "FULL"
    End If
End Sub

' ######################### END CALCULATIOn###########
' ######################### START CHECKING ###########

' ######################### END CHECKING ###########
' ######################### START BUTTON CLICKING ###########

Private Sub CmdPrev_Click()
Dim MBM As New ADODB.Recordset
Dim sql As String
    
        sql = " select MBMNumber from MBMTwo where MBMnumber < '" & RemStr(txtMBMNumber) & "'  and MBMClinical = 'Y'" & _
            " Order by MBMnumber DESC "
        MBM.MaxRecords = 1
        MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBM.RecordCount Then
            txtMBMNumber.Text = MBM!MBMNumber
            cmdShow_Click
        Else
            MsgBox "First Record Reached! " & vbNewLine & "Not a valid Membership Number Format! ", vbCritical + vbOKOnly
        End If
        MBM.Close
        Set MBM = Nothing
    
End Sub

Private Sub CmdNext_Click()
    Dim MBM As New ADODB.Recordset
    Dim sql As String
    sql = " Select MBMNumber From View_MBM_Clinical where MBMNumber > '" & RemStr(txtMBMNumber) & "' AND MBMClinical = 'Y'" & _
          " Order by MBMNumber "
    MBM.MaxRecords = 1
    MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBM.RecordCount And MBM.RecordCount <> -1 Then
        txtMBMNumber.Text = Trim(MBM!MBMNumber)
        cmdShow_Click
    Else
        MsgBox " Last Record Reached! ", vbCritical + vbOKOnly
    End If
    
    MBM.Close
    Set MBM = Nothing
End Sub


Private Sub lstAdjustmentHis_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstAdjustmentHis.SortKey = ColumnHeader.Index - 1
    lstAdjustmentHis.Sorted = True
    If lstAdjustmentHis.SortOrder = lvwAscending Then
        lstAdjustmentHis.SortOrder = lvwDescending
    Else
        lstAdjustmentHis.SortOrder = lvwAscending
    End If

End Sub

Private Sub lstBenefitDetails_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstBenefitDetails.SortKey = ColumnHeader.Index - 1
    lstBenefitDetails.Sorted = True
    If lstBenefitDetails.SortOrder = lvwAscending Then
        lstBenefitDetails.SortOrder = lvwDescending
    Else
        lstBenefitDetails.SortOrder = lvwAscending
    End If
End Sub

Private Sub lstCaseList_DblClick()
    If lstCaseList.ListItems.Count Then
        frmInvoiceAdj.txtCaseNo = ChkStr(lstCaseList.SelectedItem.Text)
        Call frmInvoiceAdj.CmdGo_Click
        frmInvoiceAdj.Show
    End If
End Sub

Private Sub lstNewBenefitDetails_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstNewBenefitDetails.SortKey = ColumnHeader.Index - 1
    lstNewBenefitDetails.Sorted = True
    If lstNewBenefitDetails.SortOrder = lvwAscending Then
        lstNewBenefitDetails.SortOrder = lvwDescending
    Else
        lstNewBenefitDetails.SortOrder = lvwAscending
    End If
End Sub

Private Sub lstCaseList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstCovAdd.SortKey = ColumnHeader.Index - 1
    lstCovAdd.Sorted = True
    If lstCovAdd.SortOrder = lvwAscending Then
        lstCovAdd.SortOrder = lvwDescending
    Else
        lstCovAdd.SortOrder = lvwAscending
    End If
End Sub

Private Sub lstCovAdd_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstCovAdd.SortKey = ColumnHeader.Index - 1
    lstCovAdd.Sorted = True
    If lstCovAdd.SortOrder = lvwAscending Then
        lstCovAdd.SortOrder = lvwDescending
    Else
        lstCovAdd.SortOrder = lvwAscending
    End If

End Sub

Private Sub lstCovAdd_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstCovAdd.ListItems.Count Then
        txtCovName = lstCovAdd.SelectedItem.Text
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate.mask = ""
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        cboCovBlood = lstCovAdd.SelectedItem.SubItems(7)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(12), "###,##0.00")
        txtAvailLimit = lstCovAdd.SelectedItem.SubItems(13)
        cboOccupation = lstCovAdd.SelectedItem.SubItems(14)
        cboWorkNature = lstCovAdd.SelectedItem.SubItems(15)
        cboSmoking = lstCovAdd.SelectedItem.SubItems(16)
        cboAlcohol = lstCovAdd.SelectedItem.SubItems(17)
        txtSuppPrem = Format(lstCovAdd.SelectedItem.SubItems(18), "#,##0.00")
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(19)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(20)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(21)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(22)
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(23)
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(24)
        
    End If
    Call GetAccountSummary
End Sub

Private Sub lstCovAdd_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstCovAdd.ListItems.Count Then
        txtCovName = lstCovAdd.SelectedItem.Text
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate.mask = ""
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        cboCovBlood = lstCovAdd.SelectedItem.SubItems(7)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(12), "###,##0.00")
        txtAvailLimit = lstCovAdd.SelectedItem.SubItems(13)
        cboOccupation = lstCovAdd.SelectedItem.SubItems(14)
        cboWorkNature = lstCovAdd.SelectedItem.SubItems(15)
        cboSmoking = lstCovAdd.SelectedItem.SubItems(16)
        cboAlcohol = lstCovAdd.SelectedItem.SubItems(17)
        txtSuppPrem = Format(lstCovAdd.SelectedItem.SubItems(18), "#,##0.00")
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(19)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(20)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(21)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(22)
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(23)
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(24)
        
    End If
    Call GetAccountSummary
End Sub

Private Sub lstMBMAcc_Click()
Dim strsql As String
Dim MBM As New ADODB.Recordset
    
    If lstMBMAcc.ListItems.Count <> 0 Then
    
        strsql = " Select MBMNumber, MBMMemberType, MBMCAnnualLimit, " & _
                 " MBMCAvailableLimit From View_MBM_MemberType where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBM.RecordCount Then
        If MBM!MBMMemberType = "P" Then
            txtAnnualLmt = txtAnnualLimit
            txtAvailableAnnualLmt = txtAvailableAnnualLimit
        ElseIf MBM!MBMMemberType = "Q" And lstMBMAcc.SelectedItem.SubItems(1) <> 0 Then
            txtAnnualLmt = MBM!MBMCAnnualLimit
            txtAvailableAnnualLmt = MBM!MBMCAvailableLimit
        ElseIf MBM!MBMMemberType = "Q" And lstMBMAcc.SelectedItem.SubItems(1) = 0 Then
            txtAnnualLmt = txtAnnualLimit
            txtAvailableAnnualLmt = txtAvailableAnnualLimit
        End If
        
        txtApprovedAmt = Format(lstMBMAcc.SelectedItem.SubItems(2), "#,##0.00")
        txtPayable2Hos = Format(lstMBMAcc.SelectedItem.SubItems(3), "#,##0.00")
        If IsNumeric(txtPayable2Hos) = False Then
            txtPayable2Hos = 0
        End If
        If Trim(lstMBMAcc.SelectedItem.SubItems(4)) <> "" Then
            txtPaid2Hos = Format(lstMBMAcc.SelectedItem.SubItems(4), "#,##0.00")
        End If
        If IsNumeric(txtPaid2Hos) = False Then
            txtPaid2Hos = 0
        End If
        
        ' Out standing Clinic
        txtHosOS = Format(CDbl(txtPayable2Hos) - CDbl(txtPaid2Hos), "#,##0.00;(#,##0.00)")
        
        txtPayableMem = Format(lstMBMAcc.SelectedItem.SubItems(5), "#,##0.00;(#,##0.00)")
        If IsNumeric(txtPayableMem) = False Then
            txtPayableMem = 0
        End If
        If Trim(lstMBMAcc.SelectedItem.SubItems(6)) <> "" Then
            txtPaid2MBM = Format(lstMBMAcc.SelectedItem.SubItems(6), "#,##0.00")
        End If
        If IsNumeric(txtPaid2MBM) = False Then
            txtPaid2MBM = 0
        End If
        If Trim(lstMBMAcc.SelectedItem.SubItems(7)) <> "" Then
            txtPaidByMem = Format(lstMBMAcc.SelectedItem.SubItems(7), "#,##0.00")
        End If
        If IsNumeric(txtPaidByMem) = False Then
            txtPaidByMem = 0
        End If
        txtMemOS = Format(CDbl(txtPayableMem) - CDbl(txtPaid2MBM) + CDbl(txtPaidByMem), "#,##0.00;(#,##0.00)")
        txtReimbursement = Format(lstMBMAcc.SelectedItem.SubItems(8), "#,##0.00")
        txtExcess = Format(lstMBMAcc.SelectedItem.SubItems(9), "#,##0.00;(#,##0.00)")
        txtBordxAmt = Format(lstMBMAcc.SelectedItem.SubItems(10), "#,##0.00;(#,##0.00)")
        If lstMBMAcc.SelectedItem.SubItems(9) = "" Then
            lstMBMAcc.SelectedItem.SubItems(9) = "0"
            txtExcessPaid = Format(lstMBMAcc.SelectedItem.SubItems(9) - lstMBMAcc.SelectedItem.SubItems(5), "#,#.00;(#,#.00)")
        End If

    End If
    End If
End Sub

Private Sub lstMBMHis_DblClick()
    If lstMBMHis.ListItems.Count Then
        txtMBMNumber = lstMBMHis.SelectedItem.Text
        cmdShow_Click
        SSTab1.Tab = 0
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Screen.MousePointer = vbHourglass
    
    If SSTab1.Tab = 2 Then
        If Mid(cboINSCode, 1, 1) = "I" Or Mid(cboINSCode, 1, 1) = "G" Then
            MsgBox "This Plan does not have supplementary", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 0
        Else
        SSTab1.Tab = 2
        End If
        
    End If
    
    If SSTab1.Tab = 4 Then
        If lstBenefitDetails.ListItems.Count = 0 Then
            Call CISBNFListing
        End If
        If lstNewBenefitDetails.ListItems.Count = 0 Then
            'If txtPLNCode <> "" Or txtNewPolEffDate <> "__/__/____" Then
                Call ClinicBNFlisting
            'End If
        End If
    End If
    
    If SSTab1.Tab = 3 Then
        lstAdjustmentHis.ListItems.Clear
        If lstAdjustmentHis.ListItems.Count = 0 Then
            Call ADJListing
        End If
    End If
    
    If SSTab1.Tab = 5 Then
        If Trim(txtMBMNumber) <> "" And lstCaseList.ListItems.Count = 0 Then
            Call ListCases
        End If
    End If
    
    If SSTab1.Tab = 6 Then
        If lstMBMHis.ListItems.Count = 0 Then
            Call MBMHISListing
        End If
    End If
    
    If SSTab1.Tab = 7 Then
        lstMBMAcc.ListItems.Clear
        Call GetAccountSummary
    End If
        
        
    If SSTab1.Tab = 8 Then
        Call MBMLimitationListing
    End If


    If SSTab1.Tab = 9 Then
        Call SGCProcListing
    End If
Screen.MousePointer = Default
End Sub


Public Sub cmdShow_Click()
Dim getRecord As Integer
Dim MBMNumber As String
    
    MBMNumber = txtMBMNumber
    Call ClearEntries
    lstMBMAcc.ListItems.Clear
    lstSurgicalProc.ListItems.Clear
    lstLimitation.ListItems.Clear
    lstMBMHis.ListItems.Clear
    lstCaseList.ListItems.Clear
    lstNewBenefitDetails.ListItems.Clear
    lstBenefitDetails.ListItems.Clear
    lstAdjustmentHis.ListItems.Clear
    lstCovAdd.ListItems.Clear
    
    txtMBMNumber = Trim(MBMNumber)
    'getRecord = showMBM
    Screen.MousePointer = vbHourglass
    'If getRecord = 1 Then
        If txtMBMNumber = "" Then
            MsgBox "Please Enter Membership No", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMNumber.SetFocus
        Else
        
            ScanCode = Trim(txtMBMNumber)
            If Mid(ScanCode, 1, 2) = "LO" Then
                txtMBMNumber = ""
                Call SearchForm(ScanCode)
            Else

                Call FindMBM
            
            End If
        End If
    
    Screen.MousePointer = Default
    
End Sub

Private Function FindMBM() As Integer
Dim strCount As String
Dim sqlstr2 As String
Dim rst2 As New ADODB.Recordset
Dim getRecord As Integer
   sqlstr2 = " Select MBMNumber from MBM Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
   rst2.Open sqlstr2, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rst2.RecordCount = 0 Then
        MsgBox "No record found!", vbCritical
    Else
        getRecord = showMBM
        lstAdjustmentHis.ListItems.Clear
        lstBenefitDetails.ListItems.Clear
        lstNewBenefitDetails.ListItems.Clear
        lstCaseList.ListItems.Clear
        lstMBMHis.ListItems.Clear
        lstSurgicalProc.ListItems.Clear
        Call showMBMCov
        Call CallLimit(ALmtInd, "P")
        
    End If
    rst2.Close
    Set rst2 = Nothing
End Function

Private Function showMBM() As Integer
Dim sqlstr, sqlPolStatus As String
Dim sqlstqMBMRef As String
Dim MBM As New ADODB.Recordset
Dim MBMRef As New ADODB.Recordset
Dim PolStatus As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim MBMNumber As String
Dim strCount As String

    ALmtInd = 0
    txtMBMPremiumRed = ""
    txtMBMMCORed = ""
    txtMBMProviderRedCharges = ""
    txtMBMProviderEffCharges = ""
    txtMCOFeesEff = ""
    txtMBMPremiumEff = ""
    txtMBMPremium = 0
    txtMCOFees = 0
    txtMBMProviderCharges = 0
    txtAnnualLimit = 0
    
    sqlstr = " Select MBMNumber, HLTCode, MBMPolicyNo, MBMName, MBMIcBcPp, " & _
            " MBMMemberType, MBMStatus, INSCode, PLNCode, MBMRenewFlag, " & _
            " MBMAddress1, MBMAddress2, MBMAddress3, MBMPostCode, CTYCode, " & _
            " MBMTelnoH, MBMEmail, MBMSalutation, MBMSex, SRVCodeList, MBMTelNoM, " & _
            " MBMTelNoO, MBMOccupation, MBMHeight, MBMWeight, " & _
            " MBMWorkNature, MBMGroupCompany, MBMEmployeeNo, MBMBlood, " & _
            " MBMSmoking, MBMAlcohol, MBMExclusion, MBMAllergic, " & _
            " MBMMaritalStatus, MBMAccessFee, MBMCLNPremium, MBMCLNMCOfee, " & _
            " MBMFirstJoinedDate, BRNCode, MBMDOB, NATCode, RACCode," & _
            " LGNCode, STACode, MBMFirstPolNo, MBMFirstMEMNo, MBMExpiredStatus, " & _
            " BRNCode, MBMTakeOver, MBMAgentCode, MBMPolicyYear, MBMDataStatus"
            
    sqlstr = sqlstr & ", PAYCode, MBMBatchNo,ANNLmtInd,VISLmtInd,PRMInd,MCOInd, " & _
            " MBMBordxDate, AGECode, MBMPrevMemNo, MBMPrevPolNo , MBMCLNVisitLmt," & _
            " MBMPayorEffDate, MBMPayorExpDate, MBMAdultChild, MBMCLNBalLmt, HSCLNInd, " & _
            " MBMLastEndorseStatus, MBMENDtBordxDate, MBMCLMNonPayFr, MBMCLMNonPayTo, " & _
            " MBMReinstate, MBMLastAdjustmentDate, MBMInstallment, LabelStatus, PLNType, INSType, MBMPolEffDate, MBMPolExpDate, " & _
            " MBMPrevPLNCode, MBMRBAmt, MBMPAYRemarks, MBMPrevInsCom, MBMDepartment, MBMRemarks, MBMCoPayRB,MBMClmNonPayFr,MBMClmNonPayTo, MBMNewPlanCode,MBMNewPolEffDate "
        
    sqlstr = sqlstr & " From View_Clinical_MBM_MBMOthers Where MBMNumber ='" & Trim(txtMBMNumber) & "'"
    
    ReportCriteria = sqlstr
           
    If bDataStatus = True Then
        MBM.Open ReportCriteria, medixmasterconn, adOpenKeyset, adLockOptimistic
    Else
        MBM.Open ReportCriteria, medixmasterconnA, adOpenKeyset, adLockOptimistic
    End If
    Call ClearEntries
    
    If MBM.RecordCount Then
        showMBM = 1
        With MBM
            txtMBMName = IIf(!MBMStatus = "S", Trim(!MBMName) & " (Suspended)", Trim(!MBMName))
            If Len(!MBMIcBcPp) Then
                txtMBMIcBcPP = Trim(!MBMIcBcPp)
            End If
            If Len(!MBMStatus) Then
                txtMBMStatus = Trim(!MBMStatus)
            End If
            
            sqlPolStatus = " select pstatuscode ,pstatusDesc from PolicyStatus where PStatusCode ='" & Trim(txtMBMStatus) & "'"
            PolStatus.Open sqlPolStatus, medixmasterconn, adOpenKeyset, adLockOptimistic
            If PolStatus.RecordCount Then
                txtMBMStatus = PolStatus!pstatusDesc
            End If
            
            PolStatus.Close
            Set PolStatus = Nothing
            
            If Len(!MBMExpiredStatus) Then
                lblEXPStatus = !MBMExpiredStatus
            End If
            
            If Len(!MBMPolicyNo) Then
                txtMBMPolNo = Trim(!MBMPolicyNo)
            End If
            
            
            If Len(!INSType) Then
                cboINSCode = Trim(!INSType)
            End If
            If Len(!PLNType) Then
                cboPLNCode = Trim(!PLNType)
            End If
            
            If Trim(!MBMPolEffDate) <> "" Then
                txtMBMPayorEffDate = Trim(!MBMPolEffDate)
            End If
        
            If Trim(!MBMPolExpDate) <> "" Then
                txtMBMPayorExpDate = Trim(!MBMPolExpDate)
            End If
                        
            If Len(!MBMAddress1) Then
                txtMBMAdd1 = Trim(!MBMAddress1)
            End If
            
            If Len(!MBMAddress2) Then
                txtMBMAdd2 = Trim(!MBMAddress2)
            End If
            
            If Len(!MBMAddress3) Then
                txtMBMAdd3 = Trim(!MBMAddress3)
            End If
            
            If Len(!MBMPostCode) Then
                txtMBMPostCode = Trim(!MBMPostCode)
            End If
            If Len(!CTYCode) Then
                cboCTYCode = Trim(!CTYCode)
            End If
            If Len(!MBMRenewFlag) Then
                cboMBMRenewal = !MBMRenewFlag
            End If
            
            If Trim(!MBMTelnoH) <> "" Then
                txtMBMTelNoH = Trim(!MBMTelnoH)
            End If
            
            If Trim(!MBMSalutation) <> "" Then
                cboMBMSalutation = Trim(!MBMSalutation)
            End If
            If Len(!MBMSex) Then
                cboMBMSex = Trim(!MBMSex)
            End If
            If Len(Trim(!SRVCodeList)) Then
                cboSRVCode = Trim(!SRVCodeList)
            End If
            If Trim(!MBMTelNoM) <> "" Then
                txtMBMTelNoM = Trim(!MBMTelNoM)
            End If
            If Trim(!MBMTelNoO) <> "" Then
                txtMBMTelNoO = Trim(!MBMTelNoO)
            End If
            
            If Trim(!MBMGroupCompany) <> "" Then
                txtMBMGroupCompany = Trim(!MBMGroupCompany)
            End If
            If Trim(!MBMEmployeeNo) <> "" Then
                txtMBMEmployeeNo = Trim(!MBMEmployeeNo)
            End If
            
            If Trim(!MBMExclusion) <> "" Then
                txtMBMExclusion = Trim(!MBMExclusion)
            End If
            If Trim(!MBMAllergic) <> "" Then
                txtMBMAllergic = Trim(!MBMAllergic)
            End If
            If Trim(!MBMInstallment) <> "" Then
                CboInsMode = Trim(!MBMInstallment)
            End If
            If Trim(!MBMMaritalStatus) <> "" Then
                cboMBMMaritalStatus = Trim(!MBMMaritalStatus)
            End If
            If Trim(!MBMFirstJoinedDate) <> "" Then
                txtMBMDateJoin = Trim(!MBMFirstJoinedDate)
            End If
            If Trim(!BRNCode) <> "" Then
                cboBRNCode = Trim(!BRNCode)
            End If
            If Trim(!NATCode) <> "" Then
                cboNATCode = Trim(!NATCode)
            End If
            If Trim(!RACCode) <> "" Then
                cboRACCode = Trim(!RACCode)
            End If
            
            If Trim(!STACode) <> "" Then
                cboSTACode = Trim(!STACode)
            End If
            If Trim(!MBMDepartment) <> "" Then
                txtDeptName = Trim(!MBMDepartment)
            End If
            
            If Trim(!MBMDOB) <> "" Then
                txtMBMBirthDate.mask = ""
                txtMBMBirthDate.Text = Trim(!MBMDOB)
            End If
            If Trim(!MBMFirstMEMNo) <> "" Then
                txtFirstMeMNo = Trim(!MBMFirstMEMNo)
            End If
            If Trim(!BRNCode) <> "" Then
                cboBRNCode = Trim(!BRNCode)
            End If
            If Trim(!MBMTakeOver) = "C" Then
                cboMBMTakeOver.ListIndex = 2
            ElseIf Trim(!MBMTakeOver) = "Y" Then
                cboMBMTakeOver.ListIndex = 1
            Else
                cboMBMTakeOver.ListIndex = 0
            End If
            If Len(!MBMAgentCode) Then
                txtMBMAgentCode = Trim(!MBMAgentCode)
            End If
            'If Trim(!LabelStatus) = "Y" Then
            '    Check1.Value = 1
            'End If
            If Len(!MBMPrevPLNCode) Then
                txtMBMPrevPlan = !MBMPrevPLNCode
            End If
            If Len(!MBMRBAmt) Then
                txtMBMRBAmt = !MBMRBAmt
            End If
            If Len(!MBMPAYRemarks) Then
                txtMBMPAYRemarks = !MBMPAYRemarks
            End If
            If Len(!MBMPrevInsCom) Then
                txtMBMPrevINSCom = !MBMPrevInsCom
            End If
            If Trim(!MBMMemberType) = "P" Then
                cboMBMRelCode = "C - Combined Limit"
            ElseIf Trim(!MBMMemberType) = "Q" Then
                cboMBMRelCode = "S - Separated Limit"
            Else
                cboMBMRelCode = Trim(!MBMMemberType)
            End If
            If Trim(!MBMDataStatus) = "A" Then
                txtDataStatus = "Active"
            Else
                txtDataStatus = "Deleted"
            End If
            If Len(!PAYCode) Then
                cboPAYCode = Trim(!PAYCode)
            End If
            If Len(!AgeCode) Then
                txtAGECode = Trim(!AgeCode)
            End If
            If Len(!MBMBordxDate) Then
                txtMBMBordxDate.mask = ""
                txtMBMBordxDate = Trim(!MBMBordxDate)
            End If
            If Len(!MBMBatchNo) Then
                txtMBMBatchNo = Trim(!MBMBatchNo)
            End If
            If Len(!HLTCode) Then
                cboHLTCode = Trim(!HLTCode)
            End If
            
            If Len(!MBMPolicyYear) Then
                txtMBMPolicyYear = Trim(!MBMPolicyYear)
            End If
            
            If !AnnLmtIND = "1" Then
                If Len(!MBMCLNBalLmt) Then
                    txtAvailableAnnualLimit = Format(Trim(!MBMCLNBalLmt), "#,##0.00")
                End If
            Else
                txtAvailableAnnualLimit = "FULL"
            End If
            
            If !VisLmtIND = "1" Then
                If Len(!MBMCLNVisitLmt) Then
                    lblVisitLmt = Format(Trim(!MBMCLNVisitLmt), "#,##0.00")
                End If
            Else
                    lblVisitLmt = "FULL"
            End If

                    
            If Len(!MBMAdultChild) Then
                txtMBMAdultChild = Trim(!MBMAdultChild)
            End If
            If Len(!MBMRemarks) Then
                txtRemarks = Trim(!MBMRemarks)
            End If

            If Len(!MBMCoPayRB) Then
                cboCOPayRB = !MBMCoPayRB
            End If
            
            If Trim(!MBMCLMNonPayFr) <> "" Then
                TxtClmNonPayFr = Trim(!MBMCLMNonPayFr)
            End If
            
            If Trim(!MBMCLMNonPayTo) <> "" Then
                TxtClmNonPayTo = Trim(!MBMCLMNonPayTo)
            End If
            
            'If !PRMInd = "1" Then
                If Len(!MBMCLNPremium) Then
                    txtMBMPremium = Format(IIf(IsNull(!MBMCLNPremium) = True, Format(GetPremium(Trim(cboINSCode), Trim(cboPAYCode), Trim(txtAGECode), Trim(cboPLNCode), Trim(cboHLTCode), txtMBMPayorEffDate), "#,##0.00"), !MBMCLNPremium), "####0.00")
                Else
                    txtMBMPremium = "0.00"
                End If
            'End If
                                    
            'If !MCOInd = "1" Then
                If Len(!MBMCLNMCOFee) Then
                    txtMCOFees = Format(IIf(IsNull(!MBMCLNMCOFee) = True, Format(GetMCO(Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild, txtMBMPayorEffDate, cboPLNCode), "#,##0.00"), !MBMCLNMCOFee), "####0.00")
                Else
                    txtMCOFees = "0.00"
                End If
            'End If
            
            If Len(!MBMAccessFee) Then
                txtMBMProviderCharges = Format(IIf(IsNull(!MBMAccessFee) = True, Format(GetProviderCharges(Trim(Mid(cboSRVCode, 1, 2)), Trim(cboINSCode), Trim(cboHLTCode), Trim(txtMBMAdultChild)), "####0.00"), !MBMAccessFee), "#,##0.00")
            Else
                txtMBMProviderCharges = "0.00"
            End If
            
            If Len(!MBMNewPlanCode) Then
                txtPLNCode = !MBMNewPlanCode
            End If
            
            If Len(!MBMNewPolEffDate) Then
                txtNewPolEffDate = Format(!MBMNewPolEffDate, "dd/mm/yyyy")
            End If
            If Len(!AnnLmtIND) Then
                ALmtInd = !AnnLmtIND
            Else
                ALmtInd = 0
            End If
        End With
        
            If MBM!MBMStatus = "C" Then
            If Len(MBM!MBMNumber) Then
                MBMNumber = Trim(MBM!MBMNumber)
            End If
            sqlstqMBMRef = " Select * from MBMCLNOthers where MBMNumber =  '" & Trim(MBMNumber) & "'"
            MBMRef.Open sqlstqMBMRef, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            If MBMRef.RecordCount Then
            
                With MBMRef
                    txtMBMPremiumRed = Format(Round(!MBMPremReduction, 0), "#,##0.00")
                    txtMBMMCORed = Format(Round(!MBMMCORefund, 0), "#,##0.00")
                    txtMBMProviderRedCharges = Format(Round(!MBMProviderRef, 2), "#,##0.00")
                    txtMBMProviderEffCharges = Format(Round(!MBMProviderEff, 2), "#,##0.00")
                    txtMCOFeesEff = Format(Round(!MBMMCOEffective, 0), "#,##0.00")
                    txtMBMPremiumEff = Format(Round(!MBMPREMEffective, 0), "#,##0.00")
                End With
            
            End If
            
            MBMRef.Close
            Set MBMRef = Nothing
        End If
    End If
    MBM.Close
    Set MBM = Nothing
End Function

Private Sub showMBMCov()
    Dim MBMCov As New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "Select MBMCNumber, MBMCCoverID, MBMCName, MBMCSalutation," & _
        " MBMCICBCPPNo , MBMCDOB, MBMCSex, MBMCAdultChild, MBMCAgeband, " & _
        " MBMCRELCode, MBMCBloodType, MBMCAllergic,  " & _
        " MBMCExclusion, MBMCStatus , MBMCAnnualLimit , MBMCAvailableLimit, " & _
        " MBMCOccupation, MBMCCLMNonPayFr, MBMCCLMNonPayTo, " & _
        " MBMCWorkNature, MBMCSmoking, MBMCAlcohol, MBMCPremium, " & _
        " MBMCPAYRemarks, MBMCRBAmt, MBMCPrevInsCom, MBMCPrevPLNCode "
        
    If bDataStatus = True Then
        sqlstr = sqlstr & " from VIEW_Clinical_MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
        MBMCov.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    Else
        sqlstr = sqlstr & " from MBMCoveredPersonsArchive where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
        MBMCov.Open sqlstr, medixmasterconnA, adOpenKeyset, adLockOptimistic
    End If

    lstCovAdd.ListItems.Clear
    If MBMCov.RecordCount Then
        Dim MemberConveredPerson As ListItem
        Do While Not MBMCov.EOF
            Set MemberConveredPerson = lstCovAdd.ListItems.Add(, , MBMCov!MBMCName)
            With MBMCov
                If Len(Trim(!MBMCSalutation)) Then
                    MemberConveredPerson.SubItems(1) = ChkStr(!MBMCSalutation)
                End If
                If Len(Trim(!MBMCICBCPPNo)) Then
                    MemberConveredPerson.SubItems(2) = !MBMCICBCPPNo
                End If
                If Len(Trim(!MBMCDOB)) Then
                    MemberConveredPerson.SubItems(3) = !MBMCDOB
                End If
                
                If Len(Trim(!MBMCSex)) Then
                    MemberConveredPerson.SubItems(4) = ChkStr(!MBMCSex)
                End If
                
                If Len(Trim(!MBMCAdultChild)) Then
                    MemberConveredPerson.SubItems(5) = ChkStr(!MBMCAdultChild)
                End If
                
                If Len(Trim(!MBMCRELCode)) Then
                    MemberConveredPerson.SubItems(6) = ChkStr(!MBMCRELCode)
                End If
                
                If Len(Trim(!MBMCBloodType)) Then
                    MemberConveredPerson.SubItems(7) = ChkStr(!MBMCBloodType)
                End If
                
                If Len(Trim(!MBMCAllergic)) Then
                    MemberConveredPerson.SubItems(8) = Trim(!MBMCAllergic)
                End If
                
                If Len(Trim(!MBMCExclusion)) Then
                    MemberConveredPerson.SubItems(9) = Trim(!MBMCExclusion)
                End If
                MemberConveredPerson.SubItems(10) = !MBMCCoverID
                
                If Len(Trim(!MBMCAnnualLimit)) Then
                    MemberConveredPerson.SubItems(12) = !MBMCAnnualLimit
                End If
                
                If Len(Trim(!MBMCAvailableLimit)) Then
                    MemberConveredPerson.SubItems(13) = !MBMCAvailableLimit
                End If
                
                If Len(Trim(!MBMCOccupation)) Then
                    MemberConveredPerson.SubItems(14) = !MBMCOccupation
                End If
 
                If Len(Trim(!MBMCWorkNature)) Then
                    MemberConveredPerson.SubItems(15) = !MBMCWorkNature
                End If
                
                If Len(Trim(!MBMCSmoking)) Then
                    MemberConveredPerson.SubItems(16) = !MBMCSmoking
                End If
                
                If Len(Trim(!MBMCAlcohol)) Then
                    MemberConveredPerson.SubItems(17) = !MBMCAlcohol
                End If
                If Len(Trim(!MBMCAvailableLimit)) Then
                    MemberConveredPerson.SubItems(17) = !MBMCAvailableLimit
                End If
                
                If Len(Trim(!MBMCPremium)) Then
                    MemberConveredPerson.SubItems(18) = !MBMCPremium
                End If
                If Len(Trim(!MBMCPrevPLNCode)) Then
                    MemberConveredPerson.SubItems(19) = !MBMCPrevPLNCode
                End If
                If Len(Trim(!MBMCRBAmt)) Then
                    MemberConveredPerson.SubItems(20) = !MBMCRBAmt
                End If
                If Len(Trim(!MBMCPAYRemarks)) Then
                    MemberConveredPerson.SubItems(21) = !MBMCPAYRemarks
                End If
                If Len(Trim(!MBMCPrevInsCom)) Then
                    MemberConveredPerson.SubItems(22) = !MBMCPrevInsCom
                End If
                If Len(Trim(!MBMCCLMNonPayFr)) Then
                    MemberConveredPerson.SubItems(23) = !MBMCCLMNonPayFr
                End If
                If Len(Trim(!MBMCCLMNonPayTo)) Then
                    MemberConveredPerson.SubItems(24) = !MBMCCLMNonPayTo
                End If
            
            End With
            MBMCov.MoveNext
        Loop
    End If
    MBMCov.Close
    Set MBMCov = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearchMBM_Click()
    bDataStatus = True
    frmSearchMBM.Source = 6
    frmSearchMBM.Show vbModal
End Sub


Private Sub lstCovAdd_Click()
    If lstCovAdd.ListItems.Count Then
        txtCovName = lstCovAdd.SelectedItem.Text
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate.mask = ""
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        cboCovBlood = lstCovAdd.SelectedItem.SubItems(7)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(12), "###,##0.00")
        txtAvailLimit = lstCovAdd.SelectedItem.SubItems(13)
        cboOccupation = lstCovAdd.SelectedItem.SubItems(14)
        cboWorkNature = lstCovAdd.SelectedItem.SubItems(15)
        cboSmoking = lstCovAdd.SelectedItem.SubItems(16)
        cboAlcohol = Format(lstCovAdd.SelectedItem.SubItems(17), "#,##0.00")
        txtSuppPrem = Format(lstCovAdd.SelectedItem.SubItems(18), "#,##0.00")
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(19)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(20)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(21)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(22)
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(23)
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(24)
        
    End If

End Sub

'  ################# END Suppplementary ######################

' ################################ Enable Tab / button   ######
Private Sub EnableEntries(status As Boolean)
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = status
        End If
    Next ctl
End Sub

Private Sub InsertPrem()
Dim MBMInsertPrem As Double
Dim MBMPrem As Double
Dim PayExpDate As String

MBMPrem = txtMCOFees
PayExpDate = txtMBMPayorExpDate

    MBMInsertPrem = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtMBMPremium, "#,##0.00"))
    txtMBMPremium = ""
    txtMBMPremium = Format(MBMInsertPrem, "#,##0.00")
End Sub

Private Sub GetAccountSummary()
    Dim Acc As New ADODB.Recordset
    Dim strsql As String
    Dim masterlist As ListItem
    lstMBMAcc.ListItems.Clear
    strsql = " select totalApprove, totalPayable2Clinic, " & _
            " totalPaid2Clinic, totalPayable2Mbm, totalPaidByMbm, CoverID, " & _
            " totalPaid2MBM , totalExcess , totalReimburse from VIEW_CISMBMTotal_Summary where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    Acc.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If Not Acc.EOF Then
        Do Until Acc.EOF
        With Acc
        Set masterlist = lstMBMAcc.ListItems.Add(, , txtMBMNumber)
            masterlist.SubItems(1) = Trim(!CoverID)
            masterlist.SubItems(2) = Format(!TotalApprove, "#,##0.00")
            masterlist.SubItems(3) = Format(!totalPayable2Clinic, "#,##0.00")
            
            If Trim(!totalPaid2Clinic) <> "" Then
                masterlist.SubItems(4) = Format(!totalPaid2Clinic, "#,##0.00")
            End If
            If IsNumeric(masterlist.SubItems(4)) = False Then
                masterlist.SubItems(4) = "0.00"
            End If
            
            masterlist.SubItems(5) = Format(!totalPayable2Mbm, "#,##0.00;(#,##0.00)")
            If IsNumeric(masterlist.SubItems(3)) = False Then
                masterlist.SubItems(5) = "0.00"
            End If
            If Trim(!totalPaid2MBM) <> "" Then
                masterlist.SubItems(6) = Format(!totalPaid2MBM, "#,##0.00")
            End If
            If IsNumeric(masterlist.SubItems(6)) = False Then
                masterlist.SubItems(6) = "0.00"
            End If
            If Trim(!totalPaidByMbm) <> "" Then
                masterlist.SubItems(7) = Format(!totalPaidByMbm, "#,##0.00")
            End If
            If IsNumeric(masterlist.SubItems(7)) = False Then
                masterlist.SubItems(7) = "0.00"
            End If
            If Trim(!TotalReimburse) <> "" Then
                masterlist.SubItems(8) = Format(!TotalReimburse, "#,##0.00")
            End If
            If IsNumeric(masterlist.SubItems(8)) = False Then
                masterlist.SubItems(8) = "0.00"
            End If
            If Trim(!TotalExcess) <> "" Then
                masterlist.SubItems(9) = Format(!TotalExcess, "#,##0.00;(#,##0.00)")
            End If
            If IsNumeric(masterlist.SubItems(9)) = False Then
                masterlist.SubItems(9) = "0.00"
            End If
            If Trim(!TotalApprove) <> "" Then
                masterlist.SubItems(10) = Format(!TotalApprove, "#,##0.00;(#,##0.00")
            End If
            If IsNumeric(masterlist.SubItems(10)) = False Then
                masterlist.SubItems(10) = "0.00"
            End If
        
        End With
        Acc.MoveNext
        Loop
    End If
    Acc.Close
    Set Acc = Nothing

End Sub



Private Sub lstLimitation_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstLimitation, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstLimitation, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstLimitation, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    Case 4
        SortListView lstLimitation, ColumnHeader.Index, ldtDatetime, m_blnDirection(4)
    Case 5
        SortListView lstLimitation, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstLimitation, ColumnHeader.Index, ldtDatetime, m_blnDirection(6)
    Case 7
        SortListView lstLimitation, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstLimitation, ColumnHeader.Index, ldtDatetime, m_blnDirection(8)
    Case 9
        SortListView lstLimitation, ColumnHeader.Index, ldtString, m_blnDirection(9)
    End Select
End Sub

Private Sub SGCProcListing()
    Dim rst2 As New ADODB.Recordset
    Dim SGC As ListItem
    Dim sql As String
    
    If lstSurgicalProc.ListItems.Count = 0 And Len(Trim(cboHLTCode)) <> 0 _
     And Len(Trim(cboINSCode)) <> 0 And Len(Trim(cboPLNCode)) <> 0 And Len(Trim(cboPAYCode)) <> 0 Then
        sql = " select SGCProcCode, SGCProcCodeDesc , " & _
          " SGCProcCatCode, SGCProcCATDesc, SGCProcCodePerc " & _
          " from View_SGCProc where PLNCOde ='" & Trim(cboPLNCode) & "'" & _
          " AND HLTCode ='" & Trim(cboHLTCode) & "'" & _
          " AND INSType ='" & Trim(cboINSCode) & "'" & _
          " AND PAYCode = '" & Trim(cboPAYCode) & "'"
        rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            Set SGC = lstSurgicalProc.ListItems.Add(, , rst2!SGCProcCatCode)
                
                If Trim(rst2!SGCProcCATDesc) <> "" Then
                    SGC.SubItems(1) = rst2!SGCProcCATDesc
                End If
                
                If rst2!SGCProcCode <> "" Then
                    SGC.SubItems(2) = rst2!SGCProcCode
                End If
                
                If Trim(rst2!SGCProcCodeDesc) <> "" Then
                    SGC.SubItems(3) = rst2!SGCProcCodeDesc
                End If
                
                If Trim(rst2!SGCProcCodePerc) <> "" Then
                    SGC.SubItems(4) = rst2!SGCProcCodePerc
                End If
                
                rst2.MoveNext
        Loop
       
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub

Private Sub txtMBMNumber_GotFocus()
cmdShow.Default = True
End Sub

Private Sub txtMBMNumber_LostFocus()
txtMBMNumber = ChkStr(txtMBMNumber)
cmdShow.Default = False
End Sub


