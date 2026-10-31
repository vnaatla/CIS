VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLGGeneration 
   Caption         =   "LG Generation"
   ClientHeight    =   9795
   ClientLeft      =   0
   ClientTop       =   240
   ClientWidth     =   12015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   12015
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport crptLGPreview 
      Left            =   3120
      Top             =   8760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   3945
      Left            =   0
      TabIndex        =   15
      Top             =   240
      Width           =   11775
      Begin VB.ComboBox cmbGLType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   125
         Top             =   2160
         Width           =   2205
      End
      Begin VB.TextBox txtComplications 
         Alignment       =   1  'Right Justify
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
         Left            =   10320
         MaxLength       =   8
         TabIndex        =   143
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtCaesarean 
         Alignment       =   1  'Right Justify
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
         Left            =   8280
         MaxLength       =   8
         TabIndex        =   141
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtmaternitynormal 
         Alignment       =   1  'Right Justify
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
         Left            =   6480
         MaxLength       =   8
         TabIndex        =   139
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInternalRemarks 
         Height          =   975
         Left            =   6480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   130
         Top             =   2480
         Width           =   4605
      End
      Begin VB.TextBox txtFaxNumber 
         Height          =   315
         Left            =   13080
         MaxLength       =   120
         TabIndex        =   128
         Top             =   3000
         Width           =   885
      End
      Begin VB.CheckBox chkCLN 
         Caption         =   "Check1"
         Height          =   255
         Left            =   11160
         TabIndex        =   42
         Top             =   240
         Width           =   255
      End
      Begin VB.ComboBox cboCLNCode 
         Height          =   315
         Left            =   6480
         TabIndex        =   4
         Top             =   240
         Width           =   4605
      End
      Begin VB.TextBox txtMBMNo 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   3525
      End
      Begin VB.TextBox txtPatient 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   530
         Width           =   3525
      End
      Begin VB.TextBox txtICNo 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   810
         Width           =   3525
      End
      Begin VB.ComboBox cboLetterHead 
         Height          =   315
         Left            =   12600
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3360
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.ComboBox txtDoctor 
         Height          =   315
         Left            =   6480
         TabIndex        =   5
         Top             =   530
         Width           =   4605
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   315
         Left            =   6480
         TabIndex        =   6
         Top             =   810
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboIndicator 
         Height          =   315
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   810
         Width           =   2445
      End
      Begin VB.TextBox txtRemarks 
         Height          =   630
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   1100
         Width           =   3525
      End
      Begin VB.TextBox txtGrpCompany 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6480
         TabIndex        =   8
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox txtRBCoPay 
         Alignment       =   1  'Right Justify
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
         Left            =   9585
         MaxLength       =   8
         TabIndex        =   12
         Top             =   1365
         Width           =   1455
      End
      Begin VB.TextBox txtLmtExceeded 
         Alignment       =   1  'Right Justify
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
         Left            =   9585
         MaxLength       =   8
         TabIndex        =   13
         Top             =   1640
         Width           =   1455
      End
      Begin VB.TextBox txtTotalExcess 
         Alignment       =   1  'Right Justify
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
         Left            =   9585
         MaxLength       =   8
         TabIndex        =   14
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtTotalBill 
         Alignment       =   1  'Right Justify
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
         Left            =   6480
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1350
         Width           =   1695
      End
      Begin VB.TextBox txtRBAmt 
         Alignment       =   1  'Right Justify
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
         Left            =   8520
         MaxLength       =   8
         TabIndex        =   48
         Top             =   555
         Width           =   1695
      End
      Begin VB.TextBox txtRBExcess 
         Alignment       =   1  'Right Justify
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
         Left            =   6480
         MaxLength       =   8
         TabIndex        =   10
         Top             =   1640
         Width           =   1695
      End
      Begin VB.TextBox txtOtherExcess 
         Alignment       =   1  'Right Justify
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
         Left            =   6480
         MaxLength       =   8
         TabIndex        =   11
         Top             =   1925
         Width           =   1695
      End
      Begin VB.TextBox txtGLAmt 
         Alignment       =   1  'Right Justify
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
         Left            =   9585
         MaxLength       =   8
         TabIndex        =   57
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtDiagnosis 
         Height          =   315
         Left            =   1440
         MaxLength       =   120
         TabIndex        =   17
         Top             =   1700
         Width           =   3525
      End
      Begin VB.TextBox txtclinicalremark 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   135
         Top             =   2820
         Width           =   3495
      End
      Begin VB.TextBox txtsplplnremark 
         Height          =   495
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   132
         Top             =   2330
         Width           =   3525
      End
      Begin VB.ComboBox cmdDelayreason 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   145
         Text            =   "cmdDelayreason"
         Top             =   3420
         Width           =   3495
      End
      Begin VB.ComboBox cmbDecline 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   122
         Text            =   "cmbDecline"
         Top             =   2040
         Width           =   3525
      End
      Begin VB.Label lbldelay 
         Caption         =   "Reason for Delay"
         Height          =   375
         Left            =   120
         TabIndex        =   144
         Top             =   3420
         Width           =   1575
      End
      Begin VB.Label lblMaternityComplications 
         Caption         =   "Maternity Complications"
         Height          =   255
         Left            =   8520
         TabIndex        =   142
         Top             =   2520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblMaternityCaesarean 
         Caption         =   "C-Section"
         Height          =   255
         Left            =   7440
         TabIndex        =   140
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblMaternityNormal 
         Caption         =   "Maternity Normal"
         Height          =   255
         Left            =   5160
         TabIndex        =   138
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblclinicalremark 
         Caption         =   "Clinical Remark"
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Top             =   2900
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Special Remark"
         Height          =   255
         Left            =   120
         TabIndex        =   133
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Internal Remarks"
         Height          =   255
         Index           =   1
         Left            =   5160
         TabIndex        =   131
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "Fax Number"
         Height          =   255
         Left            =   12840
         TabIndex        =   129
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label37 
         Caption         =   "GL Type"
         Height          =   255
         Left            =   5160
         TabIndex        =   126
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "Decline Reason"
         Height          =   255
         Left            =   120
         TabIndex        =   123
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label85 
         Caption         =   "GL Amt"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   58
         Top             =   2190
         Width           =   735
      End
      Begin VB.Label Label85 
         Caption         =   "Total Bill"
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
         Left            =   5160
         TabIndex        =   56
         Top             =   1410
         Width           =   1455
      End
      Begin VB.Label Label85 
         Caption         =   "Total Excess"
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
         Index           =   6
         Left            =   8280
         TabIndex        =   54
         Top             =   1900
         Width           =   1575
      End
      Begin VB.Label Label85 
         Caption         =   "R&&B Excess"
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
         Index           =   5
         Left            =   5160
         TabIndex        =   53
         Top             =   1700
         Width           =   1575
      End
      Begin VB.Label Label85 
         Caption         =   "R&&B Co-Pay"
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
         Index           =   4
         Left            =   8280
         TabIndex        =   52
         Top             =   1420
         Width           =   1575
      End
      Begin VB.Label Label85 
         Caption         =   "Room && Board"
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
         Left            =   7200
         TabIndex        =   51
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label85 
         Caption         =   "Other Excess"
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
         Left            =   5160
         TabIndex        =   50
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label Label85 
         Caption         =   "Limit Exceeded"
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
         Index           =   3
         Left            =   8280
         TabIndex        =   49
         Top             =   1665
         Width           =   1575
      End
      Begin VB.Label Label29 
         Caption         =   "Grp. Company"
         Height          =   255
         Left            =   5160
         TabIndex        =   45
         Top             =   1125
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Diagnosis"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Remarks"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Indicator"
         Height          =   255
         Left            =   7920
         TabIndex        =   41
         Top             =   825
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Treatment Date"
         Height          =   255
         Left            =   5160
         TabIndex        =   40
         Top             =   820
         Width           =   1215
      End
      Begin VB.Label lblLG 
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
         Height          =   555
         Left            =   7200
         TabIndex        =   39
         Top             =   300
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label Label7 
         Caption         =   "Treating Doctor"
         Height          =   255
         Left            =   5160
         TabIndex        =   32
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "IC No"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   820
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Hospital"
         Height          =   255
         Index           =   0
         Left            =   5160
         TabIndex        =   30
         Top             =   250
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Member Name"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Member No."
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
         Left            =   120
         TabIndex        =   28
         Top             =   255
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   4935
      Left            =   0
      TabIndex        =   25
      Top             =   4200
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Member List"
      TabPicture(0)   =   "frmLGGeneration.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstMBM"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "LG"
      TabPicture(1)   =   "frmLGGeneration.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "lblLGIssuedAfterLimitExInd"
      Tab(1).Control(2)=   "lblUsingLimit"
      Tab(1).Control(3)=   "lstLG"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Benefit Schedule"
      TabPicture(2)   =   "frmLGGeneration.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstPolicyWording"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Cases"
      TabPicture(3)   =   "frmLGGeneration.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label36(8)"
      Tab(3).Control(1)=   "Label36(5)"
      Tab(3).Control(2)=   "lblTotalSpeVisits"
      Tab(3).Control(3)=   "Label36(7)"
      Tab(3).Control(4)=   "Label36(6)"
      Tab(3).Control(5)=   "lblTotalSpecBordx"
      Tab(3).Control(6)=   "lblTotalSpeBill"
      Tab(3).Control(7)=   "Label36(4)"
      Tab(3).Control(8)=   "lblTotalVisit"
      Tab(3).Control(9)=   "Label36(3)"
      Tab(3).Control(10)=   "lblBillAmt"
      Tab(3).Control(11)=   "lblTotalBordx"
      Tab(3).Control(12)=   "lblTotalFFS"
      Tab(3).Control(13)=   "Label36(2)"
      Tab(3).Control(14)=   "Label36(1)"
      Tab(3).Control(15)=   "Label36(0)"
      Tab(3).Control(16)=   "lstSpecialistCases"
      Tab(3).Control(17)=   "lstCaseList"
      Tab(3).ControlCount=   18
      TabCaption(4)   =   "Case Documents"
      TabPicture(4)   =   "frmLGGeneration.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "List1"
      Tab(4).ControlCount=   1
      Begin VB.ListBox List1 
         Height          =   1815
         ItemData        =   "frmLGGeneration.frx":008C
         Left            =   -74760
         List            =   "frmLGGeneration.frx":008E
         TabIndex        =   127
         ToolTipText     =   "right click to randomize contents"
         Top             =   840
         Width           =   4335
      End
      Begin VB.Frame Frame3 
         Height          =   1575
         Left            =   120
         TabIndex        =   79
         Top             =   3120
         Width           =   11385
         Begin VB.TextBox txtCovID 
            Height          =   315
            Left            =   1560
            TabIndex        =   80
            Top             =   2280
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label lblsptopup 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4365
            TabIndex        =   137
            Top             =   555
            Width           =   1275
         End
         Begin VB.Label Label14 
            Caption         =   "SP Top Up"
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   136
            Top             =   555
            Width           =   1320
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   11400
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label lblSpecAnn 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4365
            TabIndex        =   120
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label lblSpecBal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4365
            TabIndex        =   119
            Top             =   1240
            Width           =   1275
         End
         Begin VB.Label lblSpecAcc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4365
            TabIndex        =   118
            Top             =   920
            Width           =   1275
         End
         Begin VB.Label lblPLNCode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10000
            TabIndex        =   117
            Top             =   180
            Width           =   1200
         End
         Begin VB.Label lblTopUp 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1365
            TabIndex        =   116
            Top             =   560
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "MEP Limit"
            Height          =   255
            Index           =   0
            Left            =   2880
            TabIndex        =   115
            Top             =   2610
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "1st Spec. Lmt"
            Height          =   255
            Index           =   1
            Left            =   300
            TabIndex        =   114
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
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1425
            TabIndex        =   113
            Top             =   2340
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label Label16 
            Caption         =   "Dental Ext Lmt"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   2280
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lblCovDentExtLmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   111
            Top             =   2280
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Line Line4 
            X1              =   0
            X2              =   11280
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label lblCovMepLmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3960
            TabIndex        =   110
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
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   109
            Top             =   2160
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblCovDentLmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1305
            TabIndex        =   108
            Top             =   2595
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblCovPrem 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   107
            Top             =   2400
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label lblCovPremRed 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   106
            Top             =   2280
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label Label30 
            Caption         =   "Premium"
            Height          =   255
            Left            =   480
            TabIndex        =   105
            Top             =   2280
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label31 
            Caption         =   "Prem Refund"
            Height          =   255
            Left            =   360
            TabIndex        =   104
            Top             =   2160
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label34 
            Caption         =   "Dental Limit"
            Height          =   255
            Left            =   180
            TabIndex        =   103
            Top             =   2625
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label35 
            Caption         =   "Spec. Limit"
            Height          =   255
            Left            =   300
            TabIndex        =   102
            Top             =   2010
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label lblAccumAmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1365
            TabIndex        =   101
            Top             =   920
            Width           =   1035
         End
         Begin VB.Line Line5 
            X1              =   0
            X2              =   11280
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label18 
            Caption         =   "Effective Prem"
            Height          =   255
            Left            =   195
            TabIndex        =   100
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
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   99
            Top             =   2220
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblBalLmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1365
            TabIndex        =   98
            Top             =   1240
            Width           =   1035
         End
         Begin VB.Label lblAnnLmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1365
            TabIndex        =   97
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label Label17 
            Caption         =   "PLN Code"
            Height          =   255
            Left            =   8520
            TabIndex        =   96
            Top             =   195
            Width           =   855
         End
         Begin VB.Label Label20 
            Caption         =   "LG Issued Ind"
            Height          =   255
            Left            =   8520
            TabIndex        =   95
            Top             =   555
            Width           =   1095
         End
         Begin VB.Label Label22 
            Caption         =   "Annual Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   195
            Width           =   1095
         End
         Begin VB.Label Label23 
            Caption         =   "Accum Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   920
            Width           =   855
         End
         Begin VB.Label Label24 
            Caption         =   "Balance Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   1240
            Width           =   975
         End
         Begin VB.Label Label25 
            Caption         =   "Spec. Annual Limit"
            Height          =   255
            Left            =   2640
            TabIndex        =   91
            Top             =   195
            Width           =   1575
         End
         Begin VB.Label Label26 
            Caption         =   "Spec. Accum Amt"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   90
            Top             =   920
            Width           =   1455
         End
         Begin VB.Label Label27 
            Caption         =   "Spec. Balance Limit"
            Height          =   255
            Left            =   2640
            TabIndex        =   89
            Top             =   1240
            Width           =   1575
         End
         Begin VB.Label Label14 
            Caption         =   "GP Top Up"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   88
            Top             =   555
            Width           =   960
         End
         Begin VB.Label lblLGIssuedInd 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10000
            TabIndex        =   87
            Top             =   555
            Width           =   1200
         End
         Begin VB.Label Label19 
            Caption         =   "NoAnnLimtDed"
            Height          =   255
            Left            =   5940
            TabIndex        =   86
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label Label28 
            Caption         =   "NoSpecLimtDed"
            Height          =   255
            Left            =   5940
            TabIndex        =   85
            Top             =   555
            Width           =   1215
         End
         Begin VB.Label lblNoAnnLimtDed 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7145
            TabIndex        =   84
            Top             =   180
            Width           =   1200
         End
         Begin VB.Label lblNoSpecLimtDed 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7145
            TabIndex        =   83
            Top             =   555
            Width           =   1200
         End
         Begin VB.Label lblUsingLimitDesc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   10000
            TabIndex        =   81
            Top             =   920
            Width           =   1200
         End
         Begin VB.Label Label33 
            Caption         =   "Using Limit"
            Height          =   255
            Left            =   8520
            TabIndex        =   82
            Top             =   920
            Width           =   2055
         End
      End
      Begin VB.ListBox lstPolicyWording 
         Height          =   4155
         ItemData        =   "frmLGGeneration.frx":0090
         Left            =   -74880
         List            =   "frmLGGeneration.frx":0092
         TabIndex        =   60
         ToolTipText     =   "right click to randomize contents"
         Top             =   430
         Width           =   3615
      End
      Begin MSComctlLib.ListView lstLG 
         Height          =   4365
         Left            =   -74880
         TabIndex        =   26
         Top             =   360
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7699
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   18
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LGNo."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Member No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "LG Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Issued By"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Doctor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Hospital Name"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "GL Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Diagnosis"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "EmployeeName"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "HOS Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "LGFormat"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "PLNLGRemarks"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "GLType"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Internam Remarks"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Alert message"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Cas Delay Reason"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "LG TIME"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstMBM 
         Height          =   2685
         Left            =   120
         TabIndex        =   27
         Top             =   450
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4736
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Member No."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Member Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "IC No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Staff Category"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "PLN Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "EmployeeName"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "LG Remarks"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstCaseList 
         Height          =   1875
         Left            =   -74880
         TabIndex        =   61
         Top             =   600
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
            Object.Width           =   0
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
         Height          =   1635
         Left            =   -74880
         TabIndex        =   62
         Top             =   2880
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   2884
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
            Object.Width           =   0
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
         TabIndex        =   78
         Top             =   2520
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
         Index           =   1
         Left            =   -67560
         TabIndex        =   77
         Top             =   2520
         Width           =   1215
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
         TabIndex        =   76
         Top             =   2520
         Width           =   735
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
         TabIndex        =   75
         Top             =   2520
         Width           =   975
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
         TabIndex        =   74
         Top             =   2520
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
         TabIndex        =   73
         Top             =   2520
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
         Index           =   3
         Left            =   -71520
         TabIndex        =   72
         Top             =   2520
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
         TabIndex        =   71
         Top             =   2520
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
         Index           =   4
         Left            =   -69600
         TabIndex        =   70
         Top             =   4560
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
         TabIndex        =   69
         Top             =   4560
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
         TabIndex        =   68
         Top             =   4560
         Width           =   975
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
         TabIndex        =   67
         Top             =   4560
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
         Index           =   7
         Left            =   -67800
         TabIndex        =   66
         Top             =   4560
         Width           =   615
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
         TabIndex        =   65
         Top             =   4560
         Width           =   735
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
         TabIndex        =   64
         Top             =   360
         Width           =   975
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
         TabIndex        =   63
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblUsingLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -65400
         TabIndex        =   47
         Top             =   1680
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblLGIssuedAfterLimitExInd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -64200
         TabIndex        =   46
         Top             =   1680
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "DOUBLE CLICK TO RE-PRINT LG"
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
         Height          =   375
         Left            =   -74880
         TabIndex        =   38
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "MEMBER LIST"
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
         Index           =   1
         Left            =   -74850
         TabIndex        =   37
         Top             =   555
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   33
      Top             =   9120
      Width           =   11775
      Begin VB.CommandButton cmdDecline 
         Caption         =   "Decline"
         Height          =   300
         Left            =   4550
         TabIndex        =   124
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCopyFile 
         Caption         =   "Copy File"
         Height          =   300
         Left            =   5400
         TabIndex        =   121
         Top             =   240
         Width           =   855
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -120
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   300
         Left            =   10800
         TabIndex        =   24
         Top             =   240
         Width           =   900
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   9840
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   -480
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   8040
         TabIndex        =   21
         Top             =   240
         Width           =   900
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   300
         Left            =   7150
         TabIndex        =   20
         Top             =   240
         Width           =   900
      End
      Begin VB.CommandButton cmdLG 
         Caption         =   "&LG"
         Height          =   300
         Left            =   6240
         TabIndex        =   19
         Top             =   240
         Width           =   900
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "&Preview"
         Height          =   300
         Left            =   -360
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton cmdExcess 
         Caption         =   "Excess"
         Height          =   300
         Left            =   -360
         TabIndex        =   55
         Top             =   480
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lblMBMStatus 
         Caption         =   "lblMBMStatus"
         Height          =   255
         Left            =   -360
         TabIndex        =   59
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Height          =   275
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label9 
         Caption         =   "Records"
         Height          =   275
         Left            =   1320
         TabIndex        =   35
         Top             =   270
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000015&
      Caption         =   "LG Generation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmLGGeneration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
Dim objword As Word.Application
Dim objdoc As Word.Document
Dim Letter As String
Dim LGNo As String
Dim CLNCode As String
Dim TempLG As String
Dim TempMBM As Boolean
Dim recCount As Integer
Dim EditFlag As Boolean
Dim AddFlag As Boolean
Dim ASOInd, tmpMBMCategory, tmpPatientName, tmpicno, tmpSubsidiary As String
Dim UtilisedAmt As Double
Dim IsCombined As Double
Dim tmpStaffCat As String
Dim tmpGLFormat As String
Dim tmpPolicyWording As String
Dim PictureFileName As String
Dim tmpMBMName As String
Dim tmpMBMNum As String
Dim tmpCoverID As String
Dim RecordArray() As String
Dim tmpCaseNo As String
Dim tmpCLMGPFeeBasicMed As Double
Dim tmpCLMGPFeeLTMed As Double
Dim tmpCLMGPFeeMedInj As Double
Dim tmpCLMSPFeeBasicMed As Double
Dim tmpCLMSPFeeLTMed As Double
Dim tmpCLMSPFeeMedInj As Double
Dim tmpCLMDental As Double
Dim tmpCLMMaternity As Double
Dim tmpCLMPRVCheckup As Double
Dim tmpCLMXRay As Double
Dim tmpCLMUltraSound As Double
Dim tmpCLMECG As Double
Dim tmpCLMDressing As Double
Dim tmpCLMNebulizer As Double
Dim tmpCLMID As Double
Dim tmpCLMTSSTO As Double
Dim tmpCLMMinorSurgery As Double
Dim tmpCLMPOthers As Double
Dim tmpCLMBloodDrawing As Double
Dim tmpCLMBloodGlucose As Double
Dim tmpCLMUrine As Double
Dim tmpCLMPapSmear As Double
Dim tmpCLMLOthers As Double
Dim tmpCLMDTPPolio As Double
Dim tmpCLMDTPPolioHIB As Double
Dim tmpCLMHepatitisB As Double
Dim tmpCLMMMR As Double
Dim tmpCLMIOthers As Double
Dim tmpCLMPOthersDesc As String
Dim tmpCLMLOthersDesc As String
Dim tmpCLMIOthersDesc As String
Dim tmpCLMVision As Double
Dim tmpCLMTyphiad As Double
Dim tmpCLMOtherNoBalDed As Double
Dim tmpInvINdex As String
Dim tmpCASnumber As String
Dim tmpMBMNumber As String
Dim tmpCASStatus As String
Dim tmpClaimability As String
Dim tmpCaseType As String
Dim TmpPending As String
Dim tmpRegisterDate As String
Dim tmpREgisterBy As String
Dim tmpLastUpdateDate As String
Dim tmpLastUpdateBy As String
Dim tmpClinicCode As String
Dim tmpInvoiceNo As String
Dim tmpInvoiceDate As String
Dim tmpInvoiceStmtAmt As String
Dim tmpDateVisit As String
Dim tmpMCDays As String
Dim tmpOPBatchNo As String
Dim tmpTimeVisit As String
Dim tmpMedInjType1 As String
Dim tmpMedInjType2 As String
Dim tmpMedInjType3 As String
Dim tmpMedInjType4 As String
Dim tmpMedInjType5 As String
Dim tmpMedInjDays1 As String
Dim tmpMedInjDays2 As String
Dim tmpMedInjDays3 As String
Dim tmpMedInjDays4 As String
Dim tmpMedInjDays5 As String
Dim tmpMedInjQty1 As String
Dim tmpMedInjQty2 As String
Dim tmpMedInjQty3 As String
Dim tmpMedInjQty4 As String
Dim tmpMedInjQty5 As String
Dim tmpMedInjType6 As String
Dim tmpMedInjType7 As String
Dim tmpMedInjDays6 As String
Dim tmpMedInjDays7 As String
Dim tmpMedInjQty6 As String
Dim tmpMedInjQty7 As String
Dim tmpInvoiceReceiptDate As String
Dim tmpInvoiceReceiptMode As String
Dim tmpCasRemark As String
Dim tmpCASTreatReason1 As String
Dim tmpCASTreatReason2 As String
Dim tmpCASTreatReason3 As String
Dim tmpCASTreatReason4 As String
Dim tmpCASTreatReason5 As String
Dim tmpCASDiagnosis5 As String
Dim tmpMCStatus As String
Dim tmpNoOfDiag As String
Dim tmpNoOfMedication As String
Dim tmpINVEnteredDate As String
Dim tmpINVEnteredBy As String
Dim tmpINVLastUpdateBy As String
Dim tmpINVLastUpdateDate As String
Dim tmpBillNo As String
Dim tmpBillDate As String
Dim tmpINVAmount As String
Dim tmpINVPayTo As String
Dim tmpINVPayee As String
Dim tmpINVBillAmt As String
Dim tmpGPFee As String
Dim tmpGPLTMed As String
Dim tmpGPInj As String
Dim tmpGPMed As String
Dim tmpSPFee As String
Dim tmpSPLTMed As String
Dim tmpSPInj As String
Dim tmpSPMed As String
Dim tmpClaimOpenDate As String
Dim tmpClaimOpenBy As String
Dim tmpClaimCloseDate As String
Dim tmpClaimClosedBy As String
Dim tmpClaimLastUpdateDate As String
Dim tmpClaimLastUpdateUser As String
Dim tmpCLMCoPayColl As String
Dim tmpCLMPayableByMedix2C As String
Dim tmpCLMCLNExcess As Double
Dim tmpCLMTotalActual As Double
Dim tmpCLMFinalApprove As Double
Dim tmpCLMExcess As Double
Dim tmpCLMBalanceLimit As Double
Dim tmpCLMPreviousBalance As Double
Dim tmpCLMPaidBYMedix2M As Double
Dim tmpCLMPaidBYMedix2C As Double
Dim tmpCLMAdjAnnLmt As Double
Dim tmpClmNonGPFee As Double
Dim tmpNonGPFee As Double
Dim tmpCLMMCOProcFee As Double
Dim tmpCLMMCOProcStatus As String
Dim tmpCLMProcMth As String
Dim tmpCLMProcYR As String
Dim tmpCLMCoPayCLNColl As Double
Dim tmpCLMBordxAmt As Double
Dim tmpCLMNo As String
Dim strGRPBordxFees As String
Dim CASENumberMonth As String
Dim MontlyStatus As String
Dim FinalAmountForMonthLimit As String
Dim tmpcboDrugList As Integer
Dim tmpcboDrugcnt As String
Dim tmpcboDrugLen As String
Dim tmpSaveStatus As Boolean
Dim strPLNMMPerc As String
Dim tmpLimitExceed As String
Dim tmpCLMTotalApproved As Double
Dim NewCaseNo As String
Dim Clinic As String
Dim TreatmentDate As String
Dim LGNum As String
Dim IsValidLg As Boolean
Dim AnnLmt As String
Dim tmpPLNGLNO As String
Dim tmpProdName As String
Dim tempdate As Date
Dim TemplateName As String
Dim SuppTempName As String
Dim SuppInd As String
Dim VipInd As String
Dim TmpRemarks As String
Dim PLNGPSPShared As String
Dim tmpGLAmt_Chk_update As Double
Dim tmpEmailLG As String
Dim path As String
Dim FinalPath As String
Dim tmpPymtStatus, tmpclnpanelstatus As Boolean
Dim tmpIsPending As Boolean
Dim tmpStatusinfo, chkvisitstatus As Boolean
Dim mbmtmpexpirydate As Date
Dim tmpSpecSuppLimitStatus As String
Dim Resultforsp As String
Dim tmpCoverID_template As String
Dim AfterLimitExInd As String
Dim tmpStatus, isinvoiceno, btbgtmpStatus As String
Dim tmpCLNName, tmpclnadd1, tmpclnadd2, tmpclnadd3, tmpCLNState As String
Dim tmpGLtype, tmpglpreviousremark As String
Dim tmpLGCasDate, tmpLGCasTime As String


Private IsPharmacy              As Boolean
Private IsMasterCase            As Boolean
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
Private m_blnDirection(1 To 18) As Boolean
Dim TmpEffDate As String
Dim tmpPrinStatus As String

Private Sub cboCLNCode_GotFocus()
Const METHOD_NAME = "cboCLNCode_GotFocus"
On Error GoTo debugs
Dim tmpclncode As String
Dim strsqlCLN As String
Dim CLN As New ADODB.Recordset

    strsqlCLN = ""
    tmpclncode = ""
    tmpclncode = Right(cboCLNCode, Len(cboCLNCode) - InStr(Trim(cboCLNCode), "["))
    tmpclncode = Trim(tmpclncode)
    ''tmpclncode = Mid(tmpclncode, 1, 4)
    tmpclncode = Mid(Trim(tmpclncode), 1, Len(tmpclncode) - 1)
    
    strsqlCLN = "Select Top 1 * from CLN where CLNCode = '" & tmpclncode & "' "
    CLN.Open strsqlCLN, medixmasterconn, adOpenForwardOnly, adLockOptimistic
    
    Do Until CLN.EOF
            tmpGLFormat = IIf(Len(CLN!CLNGLFormat), CLN!CLNGLFormat, "")
        CLN.MoveNext
    Loop
    CLN.Close
    Set CLN = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cboCLNCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNCode.Text, cboCLNCode.SelStart) + Chr(KeyAscii) + Mid(cboCLNCode.Text, cboCLNCode.SelStart + cboCLNCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCLNCode.ListCount - 1
 
        If NewText = LCase(Left(cboCLNCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNCode.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCLNCode.Text = ValidValue
                cboCLNCode.SelStart = 0
                cboCLNCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCLNCode_LostFocus()
Const METHOD_NAME = "cboCLNCode_LostFocus"
On Error GoTo debugs
Dim tmpclncode As String
Dim strsqlCLN As String
Dim CLN As New ADODB.Recordset

    strsqlCLN = ""
    tmpclncode = ""
    tmpclncode = Right(cboCLNCode, Len(cboCLNCode) - InStr(Trim(cboCLNCode), "["))
    tmpclncode = Trim(tmpclncode)
    ''tmpclncode = Mid(tmpclncode, 1, 4)
    tmpclncode = Mid(Trim(tmpclncode), 1, Len(tmpclncode) - 1)
    
    strsqlCLN = "Select Top 1 * from CLN where CLNCode = '" & tmpclncode & "' "
    CLN.Open strsqlCLN, medixmasterconn, adOpenForwardOnly, adLockOptimistic
    
    Do Until CLN.EOF
       ' If StrComp(Trim(Mid(cboCLNCode, 1, Count)), Trim(CLN!CLNName)) = 0 Then
       '     CLNCode = CLN!CLNCode
            tmpGLFormat = IIf(Len(CLN!CLNGLFormat), CLN!CLNGLFormat, "")
      '      Exit Do
      '  End If
        CLN.MoveNext
    Loop
    CLN.Close
    Set CLN = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub cboIndicator_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboIndicator.Text, cboIndicator.SelStart) + Chr(KeyAscii) + Mid(cboIndicator.Text, cboIndicator.SelStart + cboIndicator.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboIndicator.ListCount - 1
 
        If NewText = LCase(Left(cboIndicator.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboIndicator.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboIndicator.Text = ValidValue
                cboIndicator.SelStart = 0
                cboIndicator.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboLetterHead_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboLetterHead.Text, cboLetterHead.SelStart) + Chr(KeyAscii) + Mid(cboLetterHead.Text, cboLetterHead.SelStart + cboLetterHead.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboLetterHead.ListCount - 1
 
        If NewText = LCase(Left(cboLetterHead.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboLetterHead.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboLetterHead.Text = ValidValue
                cboLetterHead.SelStart = 0
                cboLetterHead.SelLength = Len(ValidValue)
            End If
        End If
End Sub
Public Sub BindDelayreasons()
Const METHOD_NAME = "BindDelayreasons"
On Error GoTo debugs
Dim rsSqlReason As New ADODB.Recordset
Dim strSqlReason As String
strSqlReason = ""
strSqlReason = "SELECT MessageContent FROM HISDB..MessageComboList WHERE MessageModule= 'LG Generation' AND MessageForm = 'Generate Guarantee Letter-CIS' AND MessageTag = 'Reason for delay'"
    rsSqlReason.Open strSqlReason, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
    With cmdDelayreason
        .Clear
        Do While Not rsSqlReason.EOF
            .AddItem rsSqlReason.Fields("MessageContent").Value
            rsSqlReason.MoveNext
        Loop
    End With
            
    rsSqlReason.Close
    Set rsSqlReason = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub BindDecline()
Const METHOD_NAME = "BindDecline"
On Error GoTo debugs
Dim DecReason As New ADODB.Recordset
Dim sqlstr As String
sqlstr = ""
cmbDecline.Clear
sqlstr = "Select * from Tbl_DeclineReason WHERE COALESCE(DeclineReason,'') <> ''"
DecReason.Open sqlstr, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
Do Until DecReason.EOF
    cmbDecline.AddItem DecReason!DeclineReason
    cmbDecline.ItemData(cmbDecline.NewIndex) = DecReason!ID
DecReason.MoveNext
Loop
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Public Sub Maternitytrue()
    lblMaternityNormal.Visible = True
    txtmaternitynormal.Visible = True
    lblMaternityCaesarean.Visible = True
    txtCaesarean.Visible = True
    lblMaternityComplications.Visible = False
    txtComplications.Visible = False
    txtInternalRemarks.Top = "2840"
    txtInternalRemarks.Height = "615"
End Sub
Public Sub MaternitytrueComplications()
    lblMaternityNormal.Visible = False
    txtmaternitynormal.Visible = False
    lblMaternityCaesarean.Visible = False
    txtCaesarean.Visible = False
    lblMaternityComplications.Visible = True
    txtComplications.Visible = True
    txtInternalRemarks.Top = "2840"
    txtInternalRemarks.Height = "615"
End Sub
Public Sub Maternityfalse()
    lblMaternityNormal.Visible = False
    txtmaternitynormal.Visible = False
    lblMaternityCaesarean.Visible = False
    txtCaesarean.Visible = False
    lblMaternityComplications.Visible = False
    txtComplications.Visible = False
    txtInternalRemarks.Top = "2480"
    txtInternalRemarks.Height = "975"
End Sub
Public Sub BindGLType()
Const METHOD_NAME = "BindGLType"
On Error GoTo debugs
Dim DecReason As New ADODB.Recordset
Dim sqlstr As String
sqlstr = ""
cmbGLType.Clear
sqlstr = "Select GLID,GLType,GLDesc from Tbl_GLType"
DecReason.Open sqlstr, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
Do Until DecReason.EOF
    cmbGLType.AddItem DecReason!GLType
DecReason.MoveNext
Loop
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub
Private Sub chkCLN_Click()
Const METHOD_NAME = "chkCLN_Click"
On Error GoTo debugs

Dim CLN As New ADODB.Recordset
Dim StrSql As String
   
    If lstMBM.ListItems.Count > 0 And txtGrpCompany <> "" Then
        cboCLNCode.Clear
        StrSql = "Exec SearchGrpComCLN '" & txtGrpCompany & "'"
        CLN.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        
        If chkCLN.Value = 1 Then
            If CLN.EOF = False Then
                Do Until CLN.EOF
                    cboCLNCode.AddItem CLN!CLNName & "     [ " & CLN!CLNCode & "]"
                    CLN.MoveNext
                Loop
            Else
                MsgBox "Provider Listing is not available!" & vbNewLine & _
               " Please verify with Administrator!", vbCritical + vbOKOnly
            End If
        End If
    
        CLN.Close
        Set CLN = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub



Private Sub cmbGLType_Click()
Call ValidateLG1
  If cmbGLType.Text = "Inpatient Delivery Private" Then
    Call Maternitytrue
  ElseIf cmbGLType.Text = "Inpatient Delivery Private-Complications" Then
    Call MaternitytrueComplications
  Else
    Call Maternityfalse
  End If
End Sub

Private Sub CmdAdd_Click()
Const METHOD_NAME = "cmdAdd_Click"
On Error GoTo debugs
Dim ScomiMsg As String
ScomiMsg = "Hospital KPJ Damansara & Prince Court only Applicable For Shah Hakim Zain & Family.(SCHA0000867*01)"

' code to check for lg issue.

'Dim IcNo As String
'IcNo = txtICNo.Text
'Call ValidateLg(IcNo)
Dim CBalLmt As Double

If IsValidLg Then
    If TempMBM = True Then
        If Trim(lblLGIssuedInd) = "N" Then
            MsgBox "Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
            Exit Sub
         ElseIf CDate(Format(tempdate, "DD/MM/YYYY")) < CDate(Format(DateTime.Now, "DD/MM/YYYY")) Then
            MsgBox "Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
            Exit Sub
        ElseIf Trim(lblLGIssuedInd) = "Y" Then
                If Trim(lblUsingLimit) = "A" Then
                    If Trim(lblLGIssuedAfterLimitExInd) = "N" Then
                        If IsNumeric(lblSpecBal) Then
                            If IsNumeric(lblBalLmt) Then
                            CBalLmt = CDbl(lblBalLmt) + CDbl(lblSpecBal)
                            End If
                        Else
                            If IsNumeric(lblBalLmt) Then
                            CBalLmt = CDbl(lblBalLmt)
                            Else
                            'CBalLmt = (lblBalLmt)
                            End If
                        End If
                        
                    
                        If IsNumeric(lblBalLmt) Then
                            If Format(CDbl(CBalLmt)) <= 0 Then
                                MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                'Member Using Annual Limit for LG
                                If MsgBox("Member Entitled Max LG Amount RM " & Format((CDbl(CBalLmt)), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                             
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "Y" Then
                        If IsNumeric(lblBalLmt) Then
                        
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                        Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "C" Then
                        If IsNumeric(lblBalLmt) Then
                            If Format(CDbl(lblBalLmt)) <= 0 Then
                                If MsgBox("Limit Exceeded. Please verify with " & Trim(txtGrpCompany) & " HR Dept or Health Connect. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                If MsgBox("Member Entitled Max LG Amount RM " & Format((CDbl(lblBalLmt)), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                                If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                    If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                 Else
                                    If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                 End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed . Please Check Annual Limit and Annual Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    Else
                        MsgBox "Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.", vbOKOnly + vbCritical
                        Exit Sub
                    End If
                ElseIf Trim(lblUsingLimit) = "S" Then
                    If Trim(lblLGIssuedAfterLimitExInd) = "N" Then
                        If IsNumeric(lblSpecBal) Then
                            If Format(CDbl(lblSpecBal)) <= 0 Then
                                MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                If MsgBox("Member Entitled Max LG Amount RM " & Format((CDbl(lblSpecBal)), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblSpecAnn) = "FULL" And Trim(lblSpecBal) = "FULL" Then
                                If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                    If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                 Else
                                    If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                 End If
                                 
                             '   If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                             '       Exit Sub
                             '   End If
                            ElseIf Trim(lblSpecAnn) = "N/A" And Trim(lblSpecBal) = "N/A" Then
                                If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                    If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                 Else
                                    If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                 End If
                                 
                             'If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                             '       Exit Sub
                             'End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "Y" Then
                        If IsNumeric(lblSpecBal) Then
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                        Else
                            If Trim(lblSpecAnn) = "FULL" And Trim(lblSpecBal) = "FULL" Then
                            
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                            ElseIf Trim(lblSpecAnn) = "N/A" And Trim(lblSpecBal) = "N/A" Then
                                MsgBox "Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "C" Then
                        If IsNumeric(lblSpecBal) Then
                            If Format(CDbl(lblSpecBal)) <= 0 Then
                                If MsgBox("Limit Exceeded. Please verify with " & Trim(txtGrpCompany) & " HR Dept or Health Connect. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                If MsgBox("Member Entitled Max LG Amount RM " & Format((CDbl(lblSpecBal)), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        ElseIf Trim(lblSpecBal) = Trim("FULL") Then
                            If MsgBox("Member Entitled Max LG Amount RM " & (((lblSpecBal))) & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                        Else
                            If Trim(lblSpecAnn) = "FULL" And Trim(lblSpecBal) = "FULL" Then
                            
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                            ElseIf Trim(lblSpecAnn) = "N/A" And Trim(lblSpecBal) = "N/A" Then
                                MsgBox "Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    Else
                         MsgBox "Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.", vbOKOnly + vbCritical
                         Exit Sub
                    End If
                ElseIf Trim(lblUsingLimit) = "B" Then
                    If Trim(lblLGIssuedAfterLimitExInd) = "N" Then
                         If IsNumeric(lblBalLmt) Then
                            If IsNumeric(lblSpecBal) Then
                                If (Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))) <= 0 Then
                                    MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                    Exit Sub
                                Else
                                    If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                If Trim(lblSpecBal) = "N/A" Then
                                    If Trim(lblSpecAnn) = "FULL" Then
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                                    ElseIf Trim(lblSpecAnn) = "N/A" Then
                                        If Format(CDbl(lblBalLmt)) <= 0 Then
                                            MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                            Exit Sub
                                        Else
                                            If MsgBox("Member Entitled Max LG Amount RM " & Format((CDbl(lblBalLmt)), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!", vbOKOnly + vbCritical
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit", vbOKOnly + vbCritical
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "Y" Then
                         If IsNumeric(lblBalLmt) Then
                            If IsNumeric(lblSpecBal) Then
                            
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                            Else
                                If Trim(lblSpecBal) = "N/A" Then
                                    If Trim(lblSpecAnn) = "FULL" Then
                                    
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                                    ElseIf Trim(lblSpecAnn) = "N/A" Then
                                        If Format(CDbl(lblBalLmt)) <= 0 Then
                                            MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed.", vbOKOnly + vbCritical
                                            Exit Sub
                                        Else
                                            If MsgBox("Member Entitled Max LG Amount RM " & Format((CDbl(lblBalLmt)), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!", vbOKOnly + vbCritical
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit !!", vbOKOnly + vbCritical
                                    Exit Sub
                                End If
                            End If

                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                            
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!.", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "C" Then
                         If IsNumeric(lblBalLmt) Then
                            If IsNumeric(lblSpecBal) Then
                                If (Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))) <= 0 Then
                                    If MsgBox("Limit Exceeded. Please verify with " & Trim(txtGrpCompany) & " HR Dept or Health Connect. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                Else
                                    If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                If Trim(lblSpecBal) = "N/A" Then
                                    If Trim(lblSpecAnn) = "FULL" Then
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                                    ElseIf Trim(lblSpecAnn) = "N/A" Then
                                        If Format(CDbl(lblBalLmt)) <= 0 Then
                                            MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                            Exit Sub
                                        Else
                                            If MsgBox("Member Entitled Max LG Amount RM " & Format((CDbl(lblBalLmt)), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!", vbOKOnly + vbCritical
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit", vbOKOnly + vbCritical
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                            
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    Else
                          MsgBox "Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.", vbOKOnly + vbCritical
                    End If
                Else
                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Limit Used for LG !!", vbOKOnly + vbCritical
                End If
        ElseIf Trim(lblLGIssuedInd) = "C" Then
            If MsgBox("Please verify with " & Trim(txtGrpCompany) & " HR Dept or Health Connect ! Press OK to Add LG and Cancel to Abort", _
                vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
                If Trim(lblUsingLimit) = "A" Then
                    If Trim(lblLGIssuedAfterLimitExInd) = "N" Then
                        If IsNumeric(lblBalLmt) Then
                            If Format(CDbl(lblBalLmt)) <= 0 Then
                                MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                'Member Using Annual Limit for LG
                                If MsgBox("Member Entitled Max LG Amount RM " & Format((CDbl(lblBalLmt)), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                            
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "Y" Then
                        If IsNumeric(lblBalLmt) Then
                        
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                            
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "C" Then
                        If IsNumeric(lblBalLmt) Then
                            If Format(CDbl(lblBalLmt)) <= 0 Then
                                If MsgBox("Limit Exceeded. Please verify with " & Trim(txtGrpCompany) & " HR Dept or Health Connect. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                If MsgBox("Member Entitled Max LG Amount RM " & Format((CDbl(lblBalLmt)), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                                If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                    If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                 Else
                                    If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                 End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed . Please Check Annual Limit and Annual Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    Else
                        MsgBox "Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.", vbOKOnly + vbCritical
                        Exit Sub
                    End If
                ElseIf Trim(lblUsingLimit) = "S" Then
                    If Trim(lblLGIssuedAfterLimitExInd) = "N" Then
                        If IsNumeric(lblSpecBal) Then
                            If Format(CDbl(lblSpecBal)) <= 0 Then
                                MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                If MsgBox("Member Entitled Max LG Amount RM " & Format((CDbl(lblSpecBal)), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblSpecAnn) = "FULL" And Trim(lblSpecBal) = "FULL" Then
                            
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "Y" Then
                        If IsNumeric(lblSpecBal) Then
                        
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                        Else
                            If Trim(lblSpecAnn) = "FULL" And Trim(lblSpecBal) = "FULL" Then
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                            ElseIf Trim(lblSpecAnn) = "N/A" And Trim(lblSpecBal) = "N/A" Then
                                MsgBox "Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "C" Then
                        If IsNumeric(lblSpecBal) Then
                            If Format(CDbl(lblSpecBal)) <= 0 Then
                                If MsgBox("Limit Exceeded. Please verify with " & Trim(txtGrpCompany) & " HR Dept or Health Connect. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                If MsgBox("Member Entitled Max LG Amount RM " & Format((CDbl(lblSpecBal)), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblSpecAnn) = "FULL" And Trim(lblSpecBal) = "FULL" Then
                            
                                If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                    If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                 Else
                                    If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                 End If
                            ElseIf Trim(lblSpecAnn) = "N/A" And Trim(lblSpecBal) = "N/A" Then
                                MsgBox "Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    Else
                         MsgBox "Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.", vbOKOnly + vbCritical
                         Exit Sub
                    End If
                ElseIf Trim(lblUsingLimit) = "B" Then
                    If Trim(lblLGIssuedAfterLimitExInd) = "N" Then
                         If IsNumeric(lblBalLmt) Then
                            If IsNumeric(lblSpecBal) Then
                                If (Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))) <= 0 Then
                                    MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                    Exit Sub
                                Else
                                    If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))), , "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                If Trim(lblSpecBal) = "N/A" Then
                                    If Trim(lblSpecAnn) = "FULL" Then
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                                    ElseIf Trim(lblSpecAnn) = "N/A" Then
                                        If Format(CDbl(lblBalLmt)) <= 0 Then
                                            MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                            Exit Sub
                                        Else
                                            If MsgBox("Member Entitled Max LG Amount RM " & Format((CDbl(lblBalLmt)), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!", vbOKOnly + vbCritical
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit", vbOKOnly + vbCritical
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "Y" Then
                         If IsNumeric(lblBalLmt) Then
                            If IsNumeric(lblSpecBal) Then
                             
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                            Else
                                If Trim(lblSpecBal) = "N/A" Then
                                    If Trim(lblSpecAnn) = "FULL" Then
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                                    ElseIf Trim(lblSpecAnn) = "N/A" Then
                                        If Format(CDbl(lblBalLmt)) <= 0 Then
                                            MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed.", vbOKOnly + vbCritical
                                            Exit Sub
                                        Else
                                            If MsgBox("Member Entitled Max LG Amount RM " & Format((CDbl(lblBalLmt)), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!", vbOKOnly + vbCritical
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit !!", vbOKOnly + vbCritical
                                    Exit Sub
                                End If
                            End If

                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                       Exit Sub
                                   End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!.", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "C" Then
                         If IsNumeric(lblBalLmt) Then
                            If IsNumeric(lblSpecBal) Then
                                If (Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))) <= 0 Then
                                    If MsgBox("Limit Exceeded. Please verify with " & Trim(txtGrpCompany) & " HR Dept or Health Connect. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                Else
                                    If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                If Trim(lblSpecBal) = "N/A" Then
                                    If Trim(lblSpecAnn) = "FULL" Then
                                    
                             If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                                If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                                    ElseIf Trim(lblSpecAnn) = "N/A" Then
                                        If Format(CDbl(lblBalLmt)) <= 0 Then
                                            MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                            Exit Sub
                                        Else
                                            If MsgBox("Member Entitled Max LG Amount RM " & Format((CDbl(lblBalLmt)), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!", vbOKOnly + vbCritical
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit", vbOKOnly + vbCritical
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                            If Trim(txtGrpCompany) = Trim("SCOMI & SUBSIDIARIES") Then
                             If MsgBox(ScomiMsg + " Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             Else
                                If MsgBox("Press OK to Add LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                             End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    Else
                          MsgBox "Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.", vbOKOnly + vbCritical
                    End If
                Else
                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Limit Used for LG !!", vbOKOnly + vbCritical
                End If
            Else
                Exit Sub
            End If
        Else
            MsgBox "Letter of Guarantee is Not Allowed. Please Check LG Issued Ind !!", vbOKOnly + vbCritical
            Exit Sub
        End If
        
        lblLG = ""
        lstMBM.Enabled = True
        EditFlag = False
        AddFlag = True
        SSTab.Tab = 1
        cmdLG.Enabled = True
        cmdPreview.Enabled = False
        cmdEdit.Enabled = False
        cmdLG.Caption = "&LG"
        cboIndicator.ListIndex = 0
        cboLetterHead.ListIndex = 1
        txtMBMNo.Enabled = True
        txtDate.Enabled = True
        txtRemarks.Enabled = True
        txtDiagnosis.Enabled = True
        txtDoctor.Enabled = True
        cboLetterHead.Enabled = True
        cboIndicator.Enabled = True
        cboCLNCode.Enabled = True
        chkCLN.Enabled = True
        chkCLN.Value = 0
    Else
        MsgBox "Please select a member before you create the LG", vbOKOnly + vbCritical
    End If
    Else
        MsgBox "No LG For this Plan!!!!", vbOKOnly + vbCritical
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub ValidateLG(IcNumber As String, MBMNumber As String)
Const METHOD_NAME = "ValidateLG"
On Error GoTo debugs
Dim SNo As String
'Dim PLNGPSPShared As String

 If InStr(1, MBMNumber, "-") Then
        Dim sArr() As String
        sArr = Split(MBMNumber, "-")
        SNo = sArr(0)
        tmpCoverID_template = sArr(1)
End If
    
'Call ValidateBSNUser(SNo)

Dim cmd As New ADODB.Command
Dim rst2 As New ADODB.Recordset
Dim PRM As ADODB.Parameter
Dim PrmMBMNo As ADODB.Parameter
Set cmd.ActiveConnection = medixmasterconn
cmd.CommandText = "SP_ValidateLG1"
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 9000
Set PRM = cmd.CreateParameter("ICNo", adVarChar, adParamInput, 2000, IcNumber)
cmd.Parameters.Append PRM
Set PrmMBMNo = cmd.CreateParameter("MBMNumber", adVarChar, adParamInput, 2000, SNo)
cmd.Parameters.Append PrmMBMNo
rst2.CursorLocation = adUseClient
rst2.CursorType = adOpenForwardOnly
rst2.CacheSize = 100
Set rst2 = cmd.Execute

        Do Until rst2.EOF
                If (Len(rst2!LGIssuedInd) > 0) And Trim(rst2!LGIssuedInd) = "Y" Then
                    IsValidLg = True
                ElseIf (Len(rst2!LGIssuedInd) > 0) And Trim(rst2!LGIssuedInd) = "N" Then
                    IsValidLg = False
                End If
                
                If rst2!AnnLmtIND = "N" Then
                    lblAnnLmt = "FULL"
                ElseIf rst2!AnnLmtIND = "Y" Then
                    If Len(rst2!AnnLimit) Then
                     AnnLmt = rst2!AnnLimit
                     'lblAnnLmt = rst2!AnnLimit
                    End If
                End If
                PLNGPSPShared = IIf(Len(rst2!PLNGPSPShared), rst2!PLNGPSPShared, "N")
                AfterLimitExInd = (IIf(Len(rst2!LGIssuedAfterLimitExInd), rst2!LGIssuedAfterLimitExInd, "N"))
                If tmpPrinStatus = "C" Then
                    lblMBMStatus = "C"
                Else
                    lblMBMStatus = rst2!MBMStatus
                End If
                rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
     'If IsValidLg Then
        Dim spVal As Long
        If PLNGPSPShared = "Y" Then
            If Trim(lblBalLmt) = Trim("FULL") Then
                IsValidLg = True
            ElseIf IsNumeric(lblBalLmt) Then
                If CDbl(lblBalLmt) > 0 Then
                IsValidLg = True
                Else
                IsValidLg = False
                End If
            Else
                IsValidLg = False
            End If
        Else
            If Len(lblSpecBal) Then
            
                If Trim(lblSpecBal) = Trim("N/A") Then
                    'IsValidLg = False
                ElseIf Trim(lblSpecBal) = Trim("FULL") Then
                    IsValidLg = True
                ElseIf IsNumeric(lblSpecBal) Then
                    spVal = CDbl(lblSpecBal)
                    If spVal > 0 Then
                        IsValidLg = True
                    Else
                        IsValidLg = False
                    End If
                End If
            End If
        End If
        If AfterLimitExInd = "Y" Then
        IsValidLg = True
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub ValidateBSNUser(MBMNumber As String)
Const METHOD_NAME = "ValidateBSNUser"
On Error GoTo debugs
    Dim txtGrpCompany As String
    Dim Balcmd As New ADODB.Command
    Dim BalPrm1 As New ADODB.Parameter
    Dim BalRes As New ADODB.Recordset
        Set Balcmd.ActiveConnection = MediConnCISDB
        Balcmd.CommandText = "sp_BSNMemberVerification"
        Balcmd.CommandType = adCmdStoredProc
        Balcmd.CommandTimeout = 9000
        Set BalPrm1 = Balcmd.CreateParameter("MemeberNo", adVarChar, adParamInput, 2000, MBMNumber)
        Balcmd.Parameters.Append BalPrm1
        BalRes.CursorLocation = adUseClient
        BalRes.CursorType = adOpenForwardOnly
        BalRes.CacheSize = 100
        Set BalRes = Balcmd.Execute
        Do While Not BalRes.EOF
            txtGrpCompany = BalRes!Number
            BalRes.MoveNext
        Loop
        BalRes.Close
        Set BalRes = Nothing
        If txtGrpCompany <> "0" Then
          MsgBox ("No GL to MO allow except Emergency & from 9 pm to 9 am")
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub
Private Sub cmdClear_Click()
    Call ClearForm(Me)
    tmpMBMNum = ""
    tmpMBMName = ""
    LGNum = ""
    cmdLG.Enabled = False
    cmdPreview.Enabled = False
    cmdEdit.Enabled = False
    cmdLG.Caption = "&LG"
    
    lstMBM.ListItems.Clear
    lstLG.ListItems.Clear
    
    cboIndicator.Clear
    cboCLNCode.Clear
    chkCLN.Value = 0
    cboLetterHead.Clear
    
    txtMBMNo = ""
    txtPatient = ""
    txtICNo = ""
    txtRemarks = ""
    txtDiagnosis = ""
    txtDoctor = ""
    txtDate = "__/__/____"
    txtGrpCompany = ""

    txtMBMNo.Enabled = True
    txtPatient.Enabled = False
    txtICNo.Enabled = False
    txtRemarks.Enabled = False
    txtDiagnosis.Enabled = False
    txtDoctor.Enabled = False
    txtDate.Enabled = False
    txtGrpCompany.Enabled = False
 
    cboIndicator.Enabled = False
    cboCLNCode.Enabled = False
    chkCLN.Enabled = False
    cboLetterHead.Enabled = False
    '''''''''''''''''''''''
  
    lblAnnLmt = ""
    lblAccumAmt = ""
    lblBalLmt = ""
    lblSpecAnn = ""
    lblSpecAcc = ""
    lblSpecBal = ""
    lblTopUp = ""
    lblsptopup = ""
    lblNoAnnLimtDed = ""
    lblNoSpecLimtDed = ""
    lblPLNCode = ""
    lblLGIssuedInd = ""
    lblUsingLimitDesc = ""
    lblUsingLimit = ""
    lblLGIssuedAfterLimitExInd = ""
    '''''''''''''''''''''''
        
    lblCount = 0
    TempLG = ""
    tmpLGCasDate = ""
    tmpLGCasTime = ""
    TempMBM = False
    AddFlag = False
    EditFlag = False
    SSTab.Tab = 0
    Call ComboListing
End Sub


Private Sub cmdCopyFile_Click()
    'LGNum = lstLG.SelectedItem.Text
    
    If Len(LGNum) Then
    Call CopyFile(LGNum)
    LGNum = ""
    Else
      MsgBox "Please Select Record!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub cmdDecline_Click()
Const METHOD_NAME = "cmdDecline_Click"
On Error GoTo debugs
Dim RecordSaved As String
Dim strsqlLG As String

If Len(Trim(cmbDecline)) = 0 Then
    MsgBox ("Please Select Decline Reason")
    Exit Sub
End If

  LGNum = lstLG.SelectedItem.Text
    If Len(Trim(lblLG)) = 0 Then
      MsgBox ("Please Select a LG!")
      Exit Sub
    End If
 RecordSaved = MsgBox("Are you sure you want to decline [" & lblLG & "]?", vbYesNo + vbExclamation)
    If RecordSaved = vbYes Then
                Clinic = Right(cboCLNCode, Len(cboCLNCode) - InStr(Trim(cboCLNCode), "["))
                Clinic = Trim(Clinic)
                ''Clinic = Mid(Clinic, 1, 4)
                Clinic = Mid(Trim(Clinic), 1, Len(Clinic) - 1)
    
            strsqlLG = "EXEC UpdateLG2 '" & ChkStr(lblLG) & "', '" & Format(DateTime.Now, "MM/DD/YYYY") & "', '" & ChkStr(txtDoctor) & "', " & _
                        "'" & Clinic & "', '" & cmbDecline & "', '" & Logon & "','" & Trim(txtDiagnosis) & "','" & Trim(cmbDecline) & "'"
            MediConnCISDB.Execute strsqlLG
            
            If IsPharmacy Then
                strsqlLG = "EXEC sp_UpdateMedReqDeclineReason '" & ChkStr(lblLG) & "', '" & Trim(cmbDecline) & "'"
                MediConnCISDB.Execute strsqlLG
            End If
            
            strsqlLG = ""
            strsqlLG = "Update TblRecord Set IsLGissued =0 , IsRejected=1,CaseProgress='LG Declined' where CaseNo='" + LGNum + "'"
            MediConnCISDB.Execute strsqlLG
            
        Call funcwordDec
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub funcwordDec()
Const METHOD_NAME = "funcwordDec"
On Error GoTo debugs
On Error GoTo openerror
Dim LetterHead As String
Dim AnnSpecLimitrst2 As New ADODB.Recordset
Dim SqlQry As String
Dim X As Boolean
Dim MBMCNUmber As String
Dim TemplateName As String
'Dim VipInd As String
Dim tmpclncode As String


'VipInd = ""

  LGNum = lstLG.SelectedItem.Text
  
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = False
'            If Mid(txtMBMNo, 1, 2) <> "AH" Then
'                TemplateName = "DECLINEGL.dot"
'                'Set objdoc = objword.Documents.Add("C:\Template\CIS LG\DECLINEGL.dot")
'                Set objdoc = objword.Documents.Add("\\Svr-doc\Template\CIS LG\" & TemplateName)
'            ElseIf Mid(txtMBMNo, 1, 2) = "AH" Then
'                TemplateName = "AHOPDecline.dot"
'                Set objdoc = objword.Documents.Add("\\Svr-doc\Template\CIS LG\" & TemplateName)
'            End If
            TemplateName = "DECLINEGL.dot"
            Set objdoc = objword.Documents.Add("\\Svr-doc\Template\CIS LG\" & TemplateName)
        End If
        
'    If Mid(txtMBMNo, 1, 2) = "AH" And AHbookmark(LetterHead) = True Then
'        Me.Show
'        Dim path1 As String
'        Dim Filename1 As String
'        'path = "\\mars\Outpatient\Guarantee Letter\" & LGNum
'        path1 = "\\Svr-doc\Outpatient\Guarantee Letter\" & LGNum
'        X = CreatePath(path1)
'        Filename1 = "\\Svr-doc\Outpatient\Guarantee Letter\" & LGNum & "\D" & Trim(LGNum) & ".doc"
'        objdoc.SaveAs2 Filename:=Filename1, FileFormat:=Word.wdFormatDocument
'        objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
'        Filename1 = "\\Svr-doc\Outpatient\Guarantee Letter\" & LGNum & "\D" & Trim(LGNum) & ".pdf"
'        objword.ActiveDocument.SaveAs Filename:=Filename1, FileFormat:=Word.wdFormatPDF
'        Set objword = Nothing
'        objdoc.Close
'        MsgBox ("Decline Letter Generated Successfully")
'            Dim Shex1 As Object
'            Set Shex1 = CreateObject("Shell.Application")
'            Shex1.Open (Filename1)
'    ElseIf bookmarkDecline(LetterHead) = True And Mid(txtMBMNo, 1, 2) <> "AH" Then
'        Me.Show
'        Dim path As String
'        Dim Filename As String
'        path = "\\Svr-doc\Outpatient\Guarantee Letter\" & LGNum
'        X = CreatePath(path)
'        Filename = "\\Svr-doc\Outpatient\Guarantee Letter\" & LGNum & "\D" & Trim(LGNum) & ".doc"
'        objdoc.SaveAs2 Filename:=Filename, FileFormat:=Word.wdFormatDocument
'        objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
'        Filename = "\\Svr-doc\Outpatient\Guarantee Letter\" & LGNum & "\D" & Trim(LGNum) & ".pdf"
'        objword.ActiveDocument.SaveAs Filename:=Filename, FileFormat:=Word.wdFormatPDF
'        Set objword = Nothing
'        objdoc.Close
'        MsgBox ("Decline Letter Generated Successfully")
'            Dim Shex As Object
'            Set Shex = CreateObject("Shell.Application")
'            Shex.Open (Filename)
'    End If
        bookmarkDecline (LetterHead)
        Me.Show
        Dim path As String
        Dim Filename As String
        path = "\\Svr-doc\Outpatient\Guarantee Letter\" & LGNum
        X = CreatePath(path)
        Filename = "\\Svr-doc\Outpatient\Guarantee Letter\" & LGNum & "\D" & Trim(LGNum) & ".doc"
        objdoc.SaveAs2 Filename:=Filename, FileFormat:=Word.wdFormatDocument
        objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
        Filename = "\\Svr-doc\Outpatient\Guarantee Letter\" & LGNum & "\D" & Trim(LGNum) & ".pdf"
        objword.ActiveDocument.SaveAs Filename:=Filename, FileFormat:=Word.wdFormatPDF
        Set objword = Nothing
        objdoc.Close
        MsgBox ("Decline Letter Generated Successfully")
        Dim Shex As Object
        Set Shex = CreateObject("Shell.Application")
        Shex.Open (Filename)
    
    If txtTotalExcess <> "" And txtTotalExcess <> "0" Then
        LetterHead = "ExcessLetter.dot"
        Set objdoc = objword.Documents.Add(LocCardPath & "\CIS\LG\" & LetterHead)
    End If
    Exit Sub
openerror:
    Me.Show
    If Err.Number = "5356" Then
        MsgBox Err.Description & "Please close the previous document first using Word or save it to other name", vbExclamation + vbOKOnly, DialogBoxTitle
    ElseIf Err.Number = "462" Then
        Set objword = New Word.Application
        objword.Visible = True
    ElseIf Err.Number = "4198" Then
        Err.Clear
    Else
        MsgBox Err.Description & Err.Number, vbExclamation + vbOKOnly, DialogBoxTitle
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub PaymentInfo()
Const METHOD_NAME = "PaymentInfo"
On Error GoTo debugs
Dim StrSql As String
Dim PayRec As New ADODB.Recordset
Dim PayList As ListItem
    
    StrSql = "SELECT PVCLNNo, PVCLNDate,  CLNPaidAmt, PVCLNChequeAmt, " & _
             "CLNRemark From Payment2CLN WHERE CASNumber='" & Trim(TempLG) & "'"
    PayRec.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockOptimistic
    Do While Not PayRec.EOF
        tmpPymtStatus = True
        PayRec.MoveNext
    Loop
    PayRec.Close
    Set PayRec = Nothing
    
    StrSql = "SELECT PVNo, PVDate, MBMPaidAmt, PVPayto, PVCLNChequeAmt, " & _
             "Remark From Payment2MBM WHERE CASNumber='" & Trim(TempLG) & "'"
    PayRec.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockOptimistic
    Do While Not PayRec.EOF
        tmpPymtStatus = True
    PayRec.MoveNext
    Loop
    PayRec.Close
    Set PayRec = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub statusInfo()
Const METHOD_NAME = "statusInfo"
On Error GoTo debugs
Dim StrSql As String
Dim StatusRec As New ADODB.Recordset
    tmpStatusinfo = True
    If tmpCoverID_template = "00" Then
        StrSql = "Select top 1 MBMCLNOthers.MBMPolExpDate from CISDB..CISCas Inner join HISDB..MBMCLNOthers ON MBMCLNOthers.MBMNumber=CISCas.MBMnumber " & _
                 "where MBMPolExpDate>= GETDATE()-30 and CISCas.CASNumber='" & Trim(lstLG.SelectedItem.Text) & "'"
    ElseIf tmpCoverID_template <> "00" Then
        StrSql = "Select top 1 MBMCoveredPersonsCLN.MBMCCLNExpDate AS MBMPolExpDate from CISDB..CISCas Inner join HISDB..MBMCoveredPersonsCLN ON MBMCoveredPersonsCLN.MBMCNumber=CISCas.MBMnumber  and MBMCoveredPersonsCLN.MBMCCoverID=CISCas.CoverID  " & _
                 "where MBMCCLNExpDate>= GETDATE()-30 and CASNumber='" & Trim(lstLG.SelectedItem.Text) & "' and CoverID = '" & Trim(tmpCoverID_template) & "'"
    End If
    StatusRec.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockOptimistic
    Do While Not StatusRec.EOF
        tmpStatusinfo = False
        mbmtmpexpirydate = StatusRec!MBMPolExpDate
        StatusRec.MoveNext
    Loop
    StatusRec.Close
    Set StatusRec = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub dailyvisitstatus()
Const METHOD_NAME = "dailyvisitstatus"
On Error GoTo debugs
Dim StrSql, MBMCNUmber As String
Dim dailylimitStatus As New ADODB.Recordset
    chkvisitstatus = False
    
    
    If Len(txtMBMNo) Then
        If InStr(txtMBMNo, "-") > 0 Then
            MBMCNUmber = Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1))
        Else
            MBMCNUmber = Trim(txtMBMNo)
        End If
    End If
    
        StrSql = ""
        StrSql = " Exec WebChkVisitLMTperDay_sp_cln '" & Trim(MBMCNUmber) & "' ,'" & Trim(tmpCoverID_template) & "' ,'" + CStr(Format(txtDate, "yyyy/mm/dd")) + "'   "
        dailylimitStatus.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
        
        If Not dailylimitStatus.EOF Then
           If dailylimitStatus!NumberofVisits = 0 Then
                chkvisitstatus = False
            Else
                chkvisitstatus = False
            End If
        End If
    
    dailylimitStatus.Close
    Set dailylimitStatus = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub spclnforpanel_check()
Const METHOD_NAME = "spclnforpanel_check"
On Error GoTo debugs
Dim StrSql As String
Dim ADOrstCLN As New ADODB.Recordset
tmpclnpanelstatus = False
StrSql = ""
StrSql = "Select CLNPanel from HISDB..CLN where  ltrim(rtrim(CLNName)) +'['+ CONVERT(varchar,ltrim(rtrim(clncode)))+']' ='" & Trim(Replace(cboCLNCode, "'", "''")) & "'"
ADOrstCLN.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockOptimistic

If Not ADOrstCLN.EOF Then
    If ADOrstCLN!CLNPanel = "N" Or ADOrstCLN!CLNPanel = "D" Then
        tmpclnpanelstatus = True
    Else
        tmpclnpanelstatus = False
    End If
Else
    tmpclnpanelstatus = True
End If

ADOrstCLN.Close
Set ADOrstCLN = Nothing

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CheckForPending()
Const METHOD_NAME = "CheckForPending"
On Error GoTo debugs
Dim StrSql As String
Dim ADOrstPend As New ADODB.Recordset
tmpclnpanelstatus = False
StrSql = ""
StrSql = "select Pending from cisdb..ciscas where CASNumber='" & lstLG.SelectedItem.Text & "'"
ADOrstPend.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockOptimistic

If Not ADOrstPend.EOF Then
    If ADOrstPend!Pending = "Y" Then
        tmpIsPending = True
    Else
        tmpIsPending = False
    End If
Else
    tmpIsPending = False
End If

ADOrstPend.Close
Set ADOrstPend = Nothing

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub spLGforSP_check()
Const METHOD_NAME = "spLGforSP_check"
On Error GoTo debugs

Dim StrSql As String
Dim ADOrstPLN As New ADODB.Recordset
Dim tmpclncode As String
Dim sArr() As String

    tmpclncode = ""
    tmpclncode = Right(cboCLNCode, Len(cboCLNCode) - InStr(Trim(cboCLNCode), "["))
    tmpclncode = Trim(tmpclncode)
'    tmpclncode = Mid(tmpclncode, 1, 4)
    tmpclncode = Mid(Trim(tmpclncode), 1, Len(tmpclncode) - 1)
    sArr = Split(txtMBMNo, "-")
    'txtMBMNo = sArr(0)

StrSql = ""
StrSql = "Exec spLGforSP_check '" & Trim(sArr(0)) & "','" & Trim(tmpclncode) & "'"
ADOrstPLN.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockOptimistic

If Not ADOrstPLN.EOF Then
    Resultforsp = ADOrstPLN!Result
End If

ADOrstPLN.Close
Set ADOrstPLN = Nothing

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdEdit_Click()
Const METHOD_NAME = "cmdEdit_Click"
On Error GoTo debugs
List1.Clear
If Trim(TempLG) = "" Then
    MsgBox "Please Select the LG !!", vbOKOnly + vbCritical
    Exit Sub
End If
        
        
        Dim Bordxstatus As New ADODB.Recordset
        Dim sqlstr As String
        sqlstr = "Select ciscas.OPBatchNo,ciscas.invoiceno,ClaimBordxStatus,claim.BTBGBordxStatus from CISWorksheet..claim Inner join CISdb..ciscas ON ciscas.CASNumber=claim.CASNumber where claim.CASNumber='" & TempLG & "'"
        Bordxstatus.Open sqlstr, MediConnCISWS, adOpenForwardOnly, adLockOptimistic
        If Not Bordxstatus.EOF Then
        tmpStatus = IIf(Len(Bordxstatus!ClaimBordxStatus), Bordxstatus!ClaimBordxStatus, "")
        isinvoiceno = IIf(Len(Trim(Bordxstatus!InvoiceNo)), Trim(Bordxstatus!InvoiceNo), "")
        btbgtmpStatus = IIf(Len(Bordxstatus!BTBGBordxStatus), Bordxstatus!BTBGBordxStatus, "")
        Bordxstatus.Close
        Set Bordxstatus = Nothing
        If tmpStatus = True Or Len(isinvoiceno) Or btbgtmpStatus = True Then
        MsgBox "This case cannot amend already bordx Or Invoice no already inserted !!", vbOKOnly + vbCritical
        tmpStatus = False
        btbgtmpStatus = False
        Exit Sub
        End If
        End If
        
        tmpPymtStatus = False
        Call PaymentInfo
        If tmpPymtStatus = True Then
        MsgBox "THE CASE CANNOT AMEND ALREADY PAYMENT DONE !!", vbOKOnly + vbCritical
        tmpPymtStatus = False
        Exit Sub
        End If
        'Strat:new PANEL clinic checking
        Resultforsp = True
        If tmpMBMCategory <> "PEMC" Then
            Call spLGforSP_check
            If Resultforsp = False Then
                MsgBox "THE SELECTED CLINIC NOT IN PANEL !!", vbOKOnly + vbCritical
                Resultforsp = True
                Exit Sub
            End If
        End If
        ''End: new PANEL clinic checking
        
        'check panel start
        If tmpMBMCategory <> "PEMC" Then
            Resultforsp = False
            Call spclnforpanel_check
                If tmpclnpanelstatus = True Then
                    MsgBox "The selected clinic is not in HC panel. GL not allowed..!!", vbOKOnly + vbCritical
                    Resultforsp = False
                    Exit Sub
                End If
        End If
        ''check panel end
        
        ''check visits start
            chkvisitstatus = False
            Call dailyvisitstatus
             If chkvisitstatus = True Then
                MsgBox "Specialist GL issued already for the same visit date...!!", vbOKOnly + vbCritical
                chkvisitstatus = False
                Exit Sub
            End If
        ''check visits end
        
        ''check status
            Call statusInfo
            If tmpStatusinfo = True Then
                MsgBox "THE ABOVE MEMBER EXPIRY..!!", vbOKOnly + vbCritical
                tmpStatusinfo = False
               Exit Sub
            End If
        ''End check status
        
        If Trim(lblLGIssuedInd) = "N" Then
            MsgBox "Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
            Exit Sub
        ElseIf Trim(lblLGIssuedInd) = "Y" Then
                If Trim(lblUsingLimit) = "A" Then
                    If Trim(lblLGIssuedAfterLimitExInd) = "N" Then
                        If IsNumeric(lblBalLmt) Then
                            Dim ChkAmt As Double
                            If IsNumeric(lblSpecBal) Then
                               
                                
                                If (Format(CDbl(lblBalLmt)) <= 0) And (CDbl(lblSpecBal) <= 0) Then
                                    MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                    Exit Sub
                                Else
                                
                                If PLNGPSPShared = "Y" Then
                                
                                    If Format(CDbl(lblBalLmt)) > 0 Then
                                         ChkAmt = CDbl(lblBalLmt)
                                    ElseIf Format(CDbl(lblSpecBal)) > 0 Then
                                         ChkAmt = CDbl(lblSpecBal)
                                    End If
                                Else
                                    If Format(CDbl(lblSpecBal)) > 0 Then
                                         ChkAmt = CDbl(lblSpecBal)
                                    End If
                                End If
                                
                                If ChkAmt = 0 Then
                                    MsgBox "Insufficient Balance, Please Check", vbOKOnly + vbCritical, DialogBoxTitle
                                    'Exit Sub
                                Else
                                
                                    If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(ChkAmt))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                End If
                               End If
                            Else
                            
                                If Format(CDbl(lblBalLmt)) <= 0 Then
                                    MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                    Exit Sub
                                Else
                                    'Member Using Annual Limit for LG
                                    If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                End If
                            
                                
                            End If
                            
                            
                            
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "Y" Then
                        If IsNumeric(lblBalLmt) Then
                            If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                Exit Sub
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "C" Then
                        If IsNumeric(lblBalLmt) Then
                            If Format(CDbl(lblBalLmt)) <= 0 Then
                                If MsgBox("Limit Exceeded. Please verify with " & Trim(txtGrpCompany) & " HR Dept or Health Connect. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed . Please Check Annual Limit and Annual Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    Else
                        MsgBox "Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.", vbOKOnly + vbCritical
                        Exit Sub
                    End If
                ElseIf Trim(lblUsingLimit) = "S" Then
                    If Trim(lblLGIssuedAfterLimitExInd) = "N" Then
                        If IsNumeric(lblSpecBal) Then
                            If Format(CDbl(lblSpecBal)) <= 0 Then
                                MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblSpecBal))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblSpecAnn) = "FULL" And Trim(lblSpecBal) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                           ' Else
                           '     MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!", vbOKOnly + vbCritical
                           '     Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "Y" Then
                        If IsNumeric(lblSpecBal) Then
                            If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                Exit Sub
                            End If
                        Else
                            If Trim(lblSpecAnn) = "FULL" And Trim(lblSpecBal) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            ElseIf Trim(lblSpecAnn) = "N/A" And Trim(lblSpecBal) = "N/A" Then
                                MsgBox "Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "C" Then
                        If IsNumeric(lblSpecBal) Then
                            If Format(CDbl(lblSpecBal)) <= 0 Then
                                If MsgBox("Limit Exceeded. Please verify with " & Trim(txtGrpCompany) & " HR Dept or Health Connect. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblSpecBal))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblSpecAnn) = "FULL" And Trim(lblSpecBal) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            ElseIf Trim(lblSpecAnn) = "N/A" And Trim(lblSpecBal) = "N/A" Then
                                MsgBox "Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    Else
                         MsgBox "Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.", vbOKOnly + vbCritical
                         Exit Sub
                    End If
                ElseIf Trim(lblUsingLimit) = "B" Then
                    If Trim(lblLGIssuedAfterLimitExInd) = "N" Then
                         If IsNumeric(lblBalLmt) Then
                            If IsNumeric(lblSpecBal) Then
                                If (Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))) <= 0 Then
                                    MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                    Exit Sub
                                Else
                                    If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                If Trim(lblSpecBal) = "N/A" Then
                                    If Trim(lblSpecAnn) = "FULL" Then
                                        If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                            Exit Sub
                                        End If
                                    ElseIf Trim(lblSpecAnn) = "N/A" Then
                                        If Format(CDbl(lblBalLmt)) <= 0 Then
                                            MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                            Exit Sub
                                        Else
                                            If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!", vbOKOnly + vbCritical
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit", vbOKOnly + vbCritical
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "Y" Then
                         If IsNumeric(lblBalLmt) Then
                            If IsNumeric(lblSpecBal) Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                If Trim(lblSpecBal) = "N/A" Then
                                    If Trim(lblSpecAnn) = "FULL" Then
                                        If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                            Exit Sub
                                        End If
                                    ElseIf Trim(lblSpecAnn) = "N/A" Then
                                        If Format(CDbl(lblBalLmt)) <= 0 Then
                                            MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed.", vbOKOnly + vbCritical
                                            Exit Sub
                                        Else
                                            If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!", vbOKOnly + vbCritical
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit !!", vbOKOnly + vbCritical
                                    Exit Sub
                                End If
                            End If

                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!.", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "C" Then
                         If IsNumeric(lblBalLmt) Then
                            If IsNumeric(lblSpecBal) Then
                                If (Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))) <= 0 Then
                                    If MsgBox("Limit Exceeded. Please verify with " & Trim(txtGrpCompany) & " HR Dept or Health Connect. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                Else
                                    If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                If Trim(lblSpecBal) = "N/A" Then
                                    If Trim(lblSpecAnn) = "FULL" Then
                                        If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                            Exit Sub
                                        End If
                                    ElseIf Trim(lblSpecAnn) = "N/A" Then
                                        If Format(CDbl(lblBalLmt)) <= 0 Then
                                            MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                            Exit Sub
                                        Else
                                            If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!", vbOKOnly + vbCritical
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit", vbOKOnly + vbCritical
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    Else
                          MsgBox "Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.", vbOKOnly + vbCritical
                    End If
                Else
                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Limit Used for LG !!", vbOKOnly + vbCritical
                End If
        ElseIf Trim(lblLGIssuedInd) = "C" Then
            If MsgBox("Please verify with " & Trim(txtGrpCompany) & " HR Dept or Health Connect ! Press OK to Edit LG and Cancel to Abort", _
                vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
                If Trim(lblUsingLimit) = "A" Then
                    If Trim(lblLGIssuedAfterLimitExInd) = "N" Then
                        If IsNumeric(lblBalLmt) Then
                            If Format(CDbl(lblBalLmt)) <= 0 Then
                                MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                'Member Using Annual Limit for LG
                                If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "Y" Then
                        If IsNumeric(lblBalLmt) Then
                            If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                Exit Sub
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "C" Then
                        If IsNumeric(lblBalLmt) Then
                            If Format(CDbl(lblBalLmt)) <= 0 Then
                                If MsgBox("Limit Exceeded. Please verify with " & Trim(txtGrpCompany) & " HR Dept or Health Connect. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed . Please Check Annual Limit and Annual Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    Else
                        MsgBox "Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.", vbOKOnly + vbCritical
                        Exit Sub
                    End If
                ElseIf Trim(lblUsingLimit) = "S" Then
                    If Trim(lblLGIssuedAfterLimitExInd) = "N" Then
                        If IsNumeric(lblSpecBal) Then
                            If Format(CDbl(lblSpecBal)) <= 0 Then
                                MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblSpecBal))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblSpecAnn) = "FULL" And Trim(lblSpecBal) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "Y" Then
                        If IsNumeric(lblSpecBal) Then
                            If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                Exit Sub
                            End If
                        Else
                            If Trim(lblSpecAnn) = "FULL" And Trim(lblSpecBal) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            ElseIf Trim(lblSpecAnn) = "N/A" And Trim(lblSpecBal) = "N/A" Then
                                MsgBox "Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "C" Then
                        If IsNumeric(lblSpecBal) Then
                            If Format(CDbl(lblSpecBal)) <= 0 Then
                                If MsgBox("Limit Exceeded. Please verify with " & Trim(txtGrpCompany) & " HR Dept or Health Connect. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblSpecBal))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblSpecAnn) = "FULL" And Trim(lblSpecBal) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            ElseIf Trim(lblSpecAnn) = "N/A" And Trim(lblSpecBal) = "N/A" Then
                                MsgBox "Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!", vbOKOnly + vbCritical
                                Exit Sub
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    Else
                         MsgBox "Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.", vbOKOnly + vbCritical
                         Exit Sub
                    End If
                ElseIf Trim(lblUsingLimit) = "B" Then
                    If Trim(lblLGIssuedAfterLimitExInd) = "N" Then
                         If IsNumeric(lblBalLmt) Then
                            If IsNumeric(lblSpecBal) Then
                                If (Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))) <= 0 Then
                                    MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                    Exit Sub
                                Else
                                    If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                If Trim(lblSpecBal) = "N/A" Then
                                    If Trim(lblSpecAnn) = "FULL" Then
                                        If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                            Exit Sub
                                        End If
                                    ElseIf Trim(lblSpecAnn) = "N/A" Then
                                        If Format(CDbl(lblBalLmt)) <= 0 Then
                                            MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                            Exit Sub
                                        Else
                                            If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!", vbOKOnly + vbCritical
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit", vbOKOnly + vbCritical
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "Y" Then
                         If IsNumeric(lblBalLmt) Then
                            If IsNumeric(lblSpecBal) Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                If Trim(lblSpecBal) = "N/A" Then
                                    If Trim(lblSpecAnn) = "FULL" Then
                                        If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                            Exit Sub
                                        End If
                                    ElseIf Trim(lblSpecAnn) = "N/A" Then
                                        If Format(CDbl(lblBalLmt)) <= 0 Then
                                            MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed.", vbOKOnly + vbCritical
                                            Exit Sub
                                        Else
                                            If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!", vbOKOnly + vbCritical
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit !!", vbOKOnly + vbCritical
                                    Exit Sub
                                End If
                            End If

                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!.", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    ElseIf Trim(lblLGIssuedAfterLimitExInd) = "C" Then
                         If IsNumeric(lblBalLmt) Then
                            If IsNumeric(lblSpecBal) Then
                                If (Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))) <= 0 Then
                                    If MsgBox("Limit Exceeded. Please verify with " & Trim(txtGrpCompany) & " HR Dept or Health Connect. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                Else
                                    If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt)) + Format(CDbl(lblSpecBal))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                If Trim(lblSpecBal) = "N/A" Then
                                    If Trim(lblSpecAnn) = "FULL" Then
                                        If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                            Exit Sub
                                        End If
                                    ElseIf Trim(lblSpecAnn) = "N/A" Then
                                        If Format(CDbl(lblBalLmt)) <= 0 Then
                                            MsgBox "Limit Exceeded. Letter of Guarantee is Not Allowed !!", vbOKOnly + vbCritical
                                            Exit Sub
                                        Else
                                            If MsgBox("Member Entitled Max LG Amount RM " & Format(Trim(Format(CDbl(lblBalLmt))), "#,##0.00") & ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!", vbOKOnly + vbCritical
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit", vbOKOnly + vbCritical
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Trim(lblAnnLmt) = "FULL" And Trim(lblBalLmt) = "FULL" Then
                                If MsgBox("Press OK to Edit LG and Cancel to Abort.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.", vbOKOnly + vbCritical
                                Exit Sub
                            End If
                        End If
                    Else
                          MsgBox "Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.", vbOKOnly + vbCritical
                    End If
                Else
                    MsgBox "Letter of Guarantee is Not Allowed. Please Check Limit Used for LG !!", vbOKOnly + vbCritical
                End If
            Else
                Exit Sub
            End If
        Else
            MsgBox "Letter of Guarantee is Not Allowed. Please Check LG Issued Ind !!", vbOKOnly + vbCritical
            Exit Sub
        End If
                
        
    EditFlag = True
    AddFlag = False
    
    txtDoctor.Enabled = True
    txtDate.Enabled = True
    txtRemarks.Enabled = True
    '''''''''''''''''''''''
    txtDiagnosis.Enabled = True
    '''''''''''''''''''''''
    cmdLG.Enabled = True
    cmdPreview.Enabled = False
    cmdLG.Caption = "&Update"
    
    cboLetterHead.Enabled = True
    cboIndicator.Enabled = True
    cboCLNCode.Enabled = True
    chkCLN.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLG_Click()
Dim tmpIsValid As Boolean
Dim tmpBal As Double
Dim RecordSaved As String
tmpBal = 0
tmpIsValid = False
IsNumeric (lblSpecBal)

        'SRI START:check for pending case
            Call CheckForPending
            If tmpIsPending = True Then
                If (MsgBox("The case is still in pending state." & vbNewLine & "If you proceed to issue GL, the pending status is turned back to 'Accepted' and the selected pending reasons will be cleared.", vbOKCancel + vbCritical, "Case Pending")) = vbCancel Then
                    Exit Sub
                End If
            End If
        'SRI END:check for pending case

    If tmpGLAmt_Chk_update = 0 Then
        If Trim(txtGLAmt) >= 0 And AfterLimitExInd = "N" Then
            If PLNGPSPShared = "Y" Then
                If IsNumeric(lblSpecBal) Then
                   tmpBal = tmpBal + CDbl(lblSpecBal)
                Else
                   tmpIsValid = True
                End If
                
                If IsNumeric(lblBalLmt) Then
                   tmpBal = tmpBal + CDbl(lblBalLmt)
                Else
                   tmpIsValid = True
                End If
                
                If lblBalLmt = "FULL" Then
                    tmpIsValid = True
                Else
                
                If IsNumeric(txtGLAmt) And IsNumeric(tmpBal) And cmbGLType = "Normal GL" Then
                    If txtGLAmt <= tmpBal Then
                        tmpIsValid = True
                    Else
                        tmpIsValid = False
                    End If
                ElseIf cmbGLType = "Normal GL" Then
                    tmpIsValid = False
                End If
                End If
            ElseIf PLNGPSPShared = "N" Then
                If IsNumeric(lblSpecBal) Then
                   tmpBal = tmpBal + CDbl(lblSpecBal)
                Else
                   tmpIsValid = True
                End If
                If lblSpecBal = "FULL" Then
                    tmpIsValid = True
                Else
                If IIf(IsNumeric(txtGLAmt), txtGLAmt, 0) <= tmpBal Then
                    tmpIsValid = True
                Else
                    tmpIsValid = False
                End If
                End If
            End If
        Else
                    tmpIsValid = True
        End If
    Else
        If Trim(txtGLAmt) >= 0 And AfterLimitExInd = "N" Then
            If PLNGPSPShared = "Y" Then
                If IsNumeric(lblSpecBal) Then
                   tmpBal = (tmpBal + CDbl(lblSpecBal)) + tmpGLAmt_Chk_update
                Else
                   tmpIsValid = True
                End If
                If IsNumeric(lblBalLmt) Then
                   tmpBal = (tmpBal + CDbl(lblBalLmt)) + tmpGLAmt_Chk_update
                Else
                   tmpIsValid = True
                End If
                If lblBalLmt = "FULL" Then
                    tmpIsValid = True
                Else
                If IsNumeric(txtGLAmt) <= tmpBal Then
                    tmpIsValid = True
                Else
                    tmpIsValid = False
                End If
                End If
                
            ElseIf PLNGPSPShared = "N" Then
                
                If IsNumeric(lblSpecBal) Then
                   tmpBal = (tmpBal + CDbl(lblSpecBal)) + tmpGLAmt_Chk_update
                Else
                   tmpIsValid = True
                End If
                If lblSpecBal = "FULL" Then
                tmpIsValid = True
                Else
                If IsNumeric(txtGLAmt) <= tmpBal Then
                    tmpIsValid = True
                Else
                    tmpIsValid = False
                End If
                End If
            End If
        Else
                    tmpIsValid = True
        End If
    End If
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        'NEW Condition
        
        Dim Bordxstatus As New ADODB.Recordset
        Dim sqlstr As String
        sqlstr = "Select ciscas.OPBatchNo,ciscas.invoiceno,ClaimBordxStatus,claim.BTBGBordxStatus from CISWorksheet..claim Inner join CISdb..ciscas ON ciscas.CASNumber=claim.CASNumber where claim.CASNumber='" & TempLG & "'"
        Bordxstatus.Open sqlstr, MediConnCISWS, adOpenForwardOnly, adLockOptimistic
        If Not Bordxstatus.EOF Then
        tmpStatus = IIf(Len(Bordxstatus!ClaimBordxStatus), Bordxstatus!ClaimBordxStatus, "")
        isinvoiceno = IIf(Len(Trim(Bordxstatus!InvoiceNo)), Trim(Bordxstatus!InvoiceNo), "")
        btbgtmpStatus = IIf(Len(Bordxstatus!BTBGBordxStatus), Bordxstatus!BTBGBordxStatus, "")
        Bordxstatus.Close
        Set Bordxstatus = Nothing
        If tmpStatus = True Or Len(isinvoiceno) Or btbgtmpStatus = True Then
        MsgBox "This case cannot amend already bordx Or Invoice no already inserted !!", vbOKOnly + vbCritical
        tmpStatus = False
        btbgtmpStatus = False
        Exit Sub
        End If
        End If
        
        tmpPymtStatus = False
        Call PaymentInfo
        If tmpPymtStatus = True Then
        MsgBox "THE CASE CANNOT AMEND ALREADY PAYMENT DONE !!", vbOKOnly + vbCritical
        tmpPymtStatus = False
        Exit Sub
        End If
        
        'check panel start
        If tmpMBMCategory <> "PEMC" Then
            Resultforsp = False
            Call spclnforpanel_check
               If tmpclnpanelstatus = True Then
                   MsgBox "The selected clinic is not in HC panel. GL not allowed..!!", vbOKOnly + vbCritical
                   Resultforsp = False
                   Exit Sub
               End If
        End If
        ''check panel end
        
        ''check visits start
            chkvisitstatus = False
            Call dailyvisitstatus
             If chkvisitstatus = True Then
                MsgBox "Specialist GL issued already for the same visit date...!!", vbOKOnly + vbCritical
                chkvisitstatus = False
                Exit Sub
            End If
        ''check visits end
        If tmpMBMCategory <> "PEMC" Then
            Resultforsp = True
            Call spLGforSP_check
            If Resultforsp = False Then
                MsgBox "THE SELECTED CLINIC NOT IN PANEL !!", vbOKOnly + vbCritical
                Resultforsp = True
                Exit Sub
            End If
        End If
         ''check status
            Call statusInfo
            If tmpStatusinfo = True Then
                MsgBox "THE ABOVE MEMBER EXPIRY..!!", vbOKOnly + vbCritical
                tmpStatusinfo = False
               Exit Sub
            End If
        ''End check status
    If Len(tmpLGCasTime) And IsDate(Format(tmpLGCasDate, "dd/mm/yyyy")) And IsDate(Format(txtDate, "dd/mm/yyyy")) Then
        If (Format(tmpLGCasDate, "dd/mm/yyyy") = Format(txtDate, "dd/mm/yyyy")) And DateTime.DateDiff("n", tmpLGCasTime, Format(Time, "hh:mm")) > 30 Then
             If Trim(cmdDelayreason) = "" Then
                MsgBox "Please Enter the reason for Delay.!", vbOKOnly + vbCritical, DialogBoxTitle
                cmdDelayreason.SetFocus
                Exit Sub
             ElseIf Len(Trim(cmdDelayreason)) Then
                RecordSaved = MsgBox("Are you want to Enter or Update the reason for Delay.?", vbYesNo + vbExclamation)
                If RecordSaved = vbYes Then
                   cmdDelayreason.SetFocus
                   Exit Sub
                End If
             End If
         End If
    End If
    If tmpIsValid Then
        tmpGLFormat = ""
        Call SaveLG
        'TempLG = ""
    Else
         MsgBox "Insufficient Balance, Please Check", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
End Sub

Private Sub CmdPreview_Click()
Const METHOD_NAME = "CmdPreview_Click"
On Error GoTo debugs
Dim CLN As New ADODB.Recordset
Dim strsqlCLN As String
Dim Count As Integer
    
    
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
    'tmpGLFormat = ""
    CLNCode = ""
    Count = InStr(1, cboCLNCode, "[")
    If Count <> 0 Then
        Count = Count - 1
    Else
        Count = Len(cboCLNCode)
    End If
    
    If tmpGLFormat = "" Then
        strsqlCLN = "Exec SearchCLN"
        CLN.Open strsqlCLN, medixmasterconn, adOpenForwardOnly, adLockOptimistic
        
        Do Until CLN.EOF
            If StrComp(Trim(Mid(cboCLNCode, 1, Count)), Trim(CLN!CLNName)) = 0 Then
                CLNCode = CLN!CLNCode
                tmpGLFormat = IIf(Len(CLN!CLNGLFormat), CLN!CLNGLFormat, "")
                Exit Do
            End If
            CLN.MoveNext
        Loop
        CLN.Close
        Set CLN = Nothing
    End If
    If tmpGLFormat <> "" Then
        'If Mid(cboLetterHead, 1, 3) = "MED" Then
        If Trim(tmpGLFormat) = "M" Then
            crptLGPreview.ReportFileName = crdPath & "\Report\MEDLGPreview.rpt"
        Else
            crptLGPreview.ReportFileName = crdPath & "\Report\UHCLGPreview.rpt"
        End If
        crptLGPreview.Destination = crptToWindow
        crptLGPreview.WindowState = crptMaximized
        crptLGPreview.SelectionFormula = "{CLN.CLNCode} = '" & Trim(CLNCode) & "'"
        crptLGPreview.ParameterFields(0) = "txtLGNo;" & lblLG & ";true"
        crptLGPreview.ParameterFields(1) = "txtCLNName;" & Trim(Mid(cboCLNCode, 1, Count)) & ";true"
        crptLGPreview.ParameterFields(2) = "txtPatient;" & txtPatient & ";true"
        crptLGPreview.ParameterFields(3) = "txtMBMNo;" & txtMBMNo & ";true"
        crptLGPreview.ParameterFields(4) = "txtICNo;" & txtICNo & ";true"
        crptLGPreview.ParameterFields(5) = "txtTreatmentDate;" & txtDate & ";true"
        crptLGPreview.ParameterFields(6) = "txtDoctor;" & txtDoctor & ";true"
        crptLGPreview.ParameterFields(7) = "txtConsultType;" & cboIndicator & ";true"
        crptLGPreview.Action = 1
    Else
        MsgBox "Please check GL format!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdSearch_Click()
    Call Maternityfalse
    tmpMBMNum = ""
    SSTab.Tab = 0
    lstCaseList.ListItems.Clear
    lstSpecialistCases.ListItems.Clear
    Dim tmpclearMBMNO As String
    tmpclearMBMNO = txtMBMNo
    Call cmdClear_Click
    cboCLNCode.Clear
    lstLG.ListItems.Clear
    lstMBM.ListItems.Clear
    txtMBMNo = tmpclearMBMNO
    lblMBMStatus = ""
    If txtMBMNo = "" Then
        MsgBox "Please key the Membership No", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMNo.SetFocus
        Exit Sub
    End If
    If lstMBM.ListItems.Count <> 0 Then
        If StrComp(lstMBM.SelectedItem.Text, txtMBMNo) = 0 Or StrComp(Mid(lstMBM.SelectedItem.Text, 1, InStr(lstMBM.SelectedItem.Text, "-") - 1), txtMBMNo) = 0 Then
        Else
            txtPatient = ""
            txtICNo = ""
            lstMBM.ListItems.Clear
        End If
    End If
    Call SearchRecords
End Sub

Private Sub Form_Load()
    Call Maternityfalse
    Call ActionLogUpdate(12, "frmLGGeneration")
    cmdLG.Enabled = False
    cmdPreview.Enabled = False
    cmdEdit.Enabled = False
    TempLG = ""
    tmpLGCasDate = ""
    tmpLGCasTime = ""
    TempMBM = False
    txtPatient.Enabled = False
    txtICNo.Enabled = False
    cboCLNCode.Enabled = False
    chkCLN.Enabled = False
    txtDoctor.Enabled = False
    txtDate.Enabled = False
    txtRemarks.Enabled = False
    '''''''''''''''''''''''
    txtDiagnosis.Enabled = False
    '''''''''''''''''''''''
    cboIndicator.Enabled = False
    Call ComboListing
    Call BindDecline
    Call BindGLType
    Call BindDelayreasons
    tmpGLAmt_Chk_update = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
            Case vbKeyReturn
                cmdSearch_Click
        End Select
    End If
End Sub

Private Sub ComboListing()
    With cboIndicator
        .AddItem "N - New"
        .AddItem "F - Follow up"
        .Text = "N - New"
    End With
    
    With cboLetterHead
        .AddItem "MED - MediExpress (Malaysia) Sdn Bhd"
        .AddItem "UHC - United Health Connect Sdn Bhd"
        .Text = "UHC - United Health Connect Sdn Bhd"
    End With
    
    With txtDoctor
        .AddItem "Not Specified"
    End With
End Sub
Private Function SearchRecords()
Const METHOD_NAME = "SearchRecords"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim rst2 As New ADODB.Recordset
Dim PRM As ADODB.Parameter
Dim strAppend As String
Dim StrSql As String
Dim Record As ListItem
Dim MBMCNUmber As String
    tmpMBMNum = ""
    tmpMBMName = ""
    recCount = 0
    StrSql = ""
    If Len(txtMBMNo) Then
        If InStr(txtMBMNo, "-") > 0 Then
            'strAppend = " AND MBMNumber = '" & Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1)) & "'"
            MBMCNUmber = Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1))
        Else
            'strAppend = " AND MBMNumber = '" & Trim(txtMBMNo) & "'"
            MBMCNUmber = Trim(txtMBMNo)
        End If
    End If
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "CLNLG1"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 9000
    Set PRM = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, MBMCNUmber)
    cmd.Parameters.Append PRM
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    If Not rst2.EOF Then
        Do Until rst2.EOF
                Set Record = lstMBM.ListItems.Add(, , rst2!MBMNumber & "-" & rst2!MBMCOVID)
                tmpMBMNum = rst2!MBMNumber
                If (rst2!MBMCOVID) = "00" Then
                    tmpMBMName = rst2!MBMName
                End If
                
                If Len(rst2!MBMName) Then
                    Record.SubItems(1) = rst2!MBMName
                End If
                If Len(rst2!MBMIcBcPp) Then
                    Record.SubItems(2) = rst2!MBMIcBcPp
                End If
                If Len(rst2!MBMGroupCompany) Then
                    txtGrpCompany = rst2!MBMGroupCompany
                    tmpPolicyWording = rst2!MBMGroupCompany
                Else
                    txtGrpCompany = ""
                    tmpPolicyWording = ""
                End If
                If Len(rst2!MBMStaffCategory) Then
                    tmpStaffCat = rst2!MBMStaffCategory
                Else
                    tmpStaffCat = ""
                End If
                txtsplplnremark.BackColor = vbYellow
                txtsplplnremark = IIf(Len(rst2!PLNSpecialRemark), rst2!PLNSpecialRemark, "")
                txtclinicalremark.BackColor = vbGreen
                txtclinicalremark = IIf(Len(rst2!clinicalRemarks), rst2!clinicalRemarks, "")
                tmpPrinStatus = IIf(Len(rst2!MBMCLNStatus), rst2!MBMCLNStatus, "")
                Record.SubItems(4) = IIf(Len(rst2!PLNType), rst2!PLNType, "")
                Record.SubItems(5) = IIf(Len(rst2!MBMName), rst2!MBMName, "")
                Record.SubItems(6) = IIf(Len(rst2!PLNLGRemarks), rst2!PLNLGRemarks, "")
                tempdate = rst2!MBMPolExpDate
                ASOInd = IIf(Len(rst2!ASOInd), rst2!ASOInd, "N")
                tmpMBMCategory = IIf(Len(rst2!MBMCategory), rst2!MBMCategory, "")
                tmpSubsidiary = IIf(Len(rst2!Subsidiary), rst2!Subsidiary, "")
                recCount = recCount + 1
                StrSql = rst2!MBMNumber
                rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        lblCount = recCount
    Else
        MsgBox "Record not found!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    If Trim(txtGrpCompany) = "INTERNATIONAL RESCUE COMMITTEE1" Then
        
    Else
        If Trim(txtGrpCompany) = "OPENET TELECOM MALAYSIA SDN BHD" Or Trim(txtMBMNo) = "STHA3342413*01" Or Trim(txtMBMNo) = "STHA3342413*02" Or Trim(txtGrpCompany) = "SURUHANJAYA PERKHIDMATAN AIR NEGARA" Or Trim(txtGrpCompany) = "COMMUNICATION CORPORATION SDN BHD" Or Trim(txtGrpCompany) = "SAINT-GOBAIN MALAYSIA SDN BHD" Or Trim(txtGrpCompany) = "PERBADANAN KEMAJUAN NEGERI SELANGOR" Or Trim(txtGrpCompany) = "DIGI TELECOMMUNICATIONS SDN BHD" Or Trim(txtMBMNo) = "STHB3164093*01" Then
            Dim SQuery As String
            Dim RSWarenty As New ADODB.Recordset
            
             SQuery = "SELECT     distinct top 1 MBMCLNWarranty From MBMCLNOthers WHERE     (MBMNumber like N'" + txtMBMNo + "%')"
                RSWarenty.Open SQuery, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                Do Until RSWarenty.EOF
                    If Len(RSWarenty!MBMCLNWarranty) Then
                         MsgBox RSWarenty!MBMCLNWarranty, vbOKOnly + vbCritical, DialogBoxTitle
                         Exit Do
                    End If
                    RSWarenty.MoveNext
                Loop
                RSWarenty.Close
                Set RSWarenty = Nothing
        End If
    End If
debugs:
If Err.Description <> "" Then
   Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Public Sub CopyFile(sLGNo As String)
Dim Found As Boolean
On Error GoTo ErrorSave

  If Len(sLGNo) Then
    CommonDialog1.CancelError = False
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.Filter = "All Files (*.*)|*.*"
    CommonDialog1.ShowOpen

    Dim Filename As String
    Dim ext As String
    Dim countRX As Integer
    Filename = CommonDialog1.Filename
    
    Dim fso
    Dim sDFolder As String
    Found = False
    countRX = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If IsPharmacy Then
        Dim validFileTypes As Variant
        
        Dim Index As Integer
        
        validFileTypes = Array("bmp", "gif", "png", "jpg", "jpeg", "pdf")
        ext = fso.GetExtensionName(Filename)
                
        For Index = LBound(validFileTypes) To UBound(validFileTypes)
          If validFileTypes(Index) = ext Then
            Found = True
            Exit For
          End If
        Next
        
        If Found Then
            countRX = TotalRXFiles(sLGNo)
            sDFolder = "\\Svr-doc\Outpatient\OtherScannedDoc\" & sLGNo & "\RX_" & sLGNo & "-" & CStr(countRX) + "." + ext
        End If
    Else
        sDFolder = "\\Svr-doc\Outpatient\Guarantee Letter\" & sLGNo & "\"
    End If
    
    If IsPharmacy And Found = False Then
        MsgBox "Invalid File Type. Please upload a File with extension (bmp, gif, png, jpg, jpeg, pdf)", vbExclamation, "Warning"
    Else
        'Checking If File Is Located in the Source Folder
        If fso.FileExists(Filename) Then
            fso.CopyFile (Filename), sDFolder, Not IsPharmacy
            MsgBox "Specified File Copied Successfully", vbInformation, "Done!"
            addTheFilesClaims (sLGNo)
        End If
    End If
  Else
     MsgBox "No LG Number Selected or Generated"
  End If
  
ErrorSave:
    If Err.Description <> "" Then
        MsgBox Err.Description, vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Function TotalRXFiles(ByVal CaseNo As String) As Integer
Dim fso
Dim fld
Dim fl
    
    TotalRXFiles = 0
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder("\\Svr-doc\Outpatient\OtherScannedDoc\" & CaseNo)
    
    For Each fl In fld.Files
        If Left(fl.Name, 3) = "RX_" Then
            TotalRXFiles = TotalRXFiles + 1
        End If
    Next fl
End Function

Private Sub SaveLG()
Const METHOD_NAME = "SaveLG"
On Error GoTo debugs
Dim CLN As New ADODB.Recordset
Dim strsqlCLN As String
Dim strsqlLG As String
Dim RecordSaved As String
Dim Count As Integer
Dim startTrans As Boolean
Dim VerifyMBM As Boolean
Dim X As Boolean
Dim tmpmaternityamt As Double

VerifyMBM = False
   
   
    
    If Trim(cboCLNCode) = "" Then
        MsgBox "Please select a Clinic", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    If Trim(cmbGLType) = "" Then
        MsgBox "Please select Type of GL", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    If Trim(txtMBMNo) = "" Then
        MsgBox "Please select a Patient", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMNo.SetFocus
        Exit Sub
    End If
    If Trim(txtPatient) = "" Then
        MsgBox "Please select a Patient Name", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    If Trim(txtICNo) = "" Then
        MsgBox "Please select Patient's IC Number", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    If Trim(txtDoctor) = "" Then
        MsgBox "Please Enter the Treating Doctor's Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtDoctor.SetFocus
        Exit Sub
    End If
    If Trim(txtGLAmt) = "" And cmbGLType.Text <> "Pre-Employee Checkup" And cmbGLType.Text <> "Inpatient Delivery Private" And cmbGLType.Text <> "Inpatient Delivery Private-Complications" And cmbGLType.Text <> "Inpatient Delivery (GOV)" Then
        MsgBox "Please Enter the GL Amount", vbOKOnly + vbCritical, DialogBoxTitle
        txtGLAmt.SetFocus
        Exit Sub
    End If
    
    If lblMBMStatus <> "A" Then
        MsgBox "Member is not active.", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    If Trim(txtDate) = "__/__/____" Then
        MsgBox "Please enter the Treatment Date.", vbOKOnly + vbCritical, DialogBoxTitle
        txtDate.SetFocus
        Exit Sub
    ElseIf IsDate(Trim(txtDate)) = False Then
        MsgBox "Invalid Date Format, please re-enter the Treatment Date.", vbOKOnly + vbCritical, DialogBoxTitle
        txtDate.SetFocus
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Trim(txtDiagnosis) = "" Then
        MsgBox "Please Enter the Diagnosis", vbOKOnly + vbCritical, DialogBoxTitle
        txtDiagnosis.SetFocus
        Exit Sub
    End If
    'Treatment Date checking
    If mbmtmpexpirydate < CDate(txtDate.Text) Then
        MsgBox "Treatment Date cannot be Greater Than Member Expiry date..!", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    
    Count = InStr(1, cboCLNCode, "[")
    If Count <> 0 Then
        Count = Count - 1
    Else
        Count = Len(cboCLNCode)
    End If
                Clinic = Right(cboCLNCode, Len(cboCLNCode) - InStr(Trim(cboCLNCode), "["))
                Clinic = Trim(Clinic)
                ''Clinic = Mid(Clinic, 1, 4)
                Clinic = Mid(Trim(Clinic), 1, Len(Clinic) - 1)
                
    Letter = "L"
    TreatmentDate = Format(txtDate, "yyyy/mm/dd")
    '''''Validate LG Template
    '''#Start-Commented by Murali
'    TemplateName = ""
'    TmpRemarks = ""
'    SuppTempName = ""
'    SuppInd = ""
'    Call ValidateLG1
    '''#End-Commented by Murali
    Dim tmpMBMNum() As String
    Dim Pos As Integer
    Dim CovID As String
    CovID = "00"
    Pos = InStr(txtMBMNo, "-")
    
    If Pos > 0 Then
        tmpMBMNum = Split(txtMBMNo, "-")
        CovID = tmpMBMNum(1)
    End If
    
    If CovID <> "00" Then
        If SuppInd = "Y" Then
            TemplateName = SuppTempName
        End If
    End If
    
    If Len(TemplateName) Then
    RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    
    If RecordSaved = vbYes Then
        
        tmpPymtStatus = False
        Call PaymentInfo
        If tmpPymtStatus = True Then
            ''MsgBox "THE CASE CANNOT AMEND ALREADY PAYMENT DONE !!", vbOKOnly + vbCritical
            ''tmpPymtStatus = False
            Exit Sub
        End If
        
        
        If EditFlag = False And tmpPymtStatus = False Then
            If EditFlag = False Then
               MsgBox "Kindly Press Edit..Then Continue", vbOKOnly + vbCritical, DialogBoxTitle
               Exit Sub
            End If
            If Mid(cboIndicator, 1, 1) = "N" Then
                
            ElseIf Mid(cboIndicator, 1, 1) = "F" Then
                If TempLG = "" Then
                    MsgBox "Please Select the LG No. Which You Required to Follow Up", vbOKOnly + vbCritical, DialogBoxTitle
                    SSTab.Tab = 1
                    Exit Sub
                End If
            End If
            
            startTrans = True
            MediConnCISDB.beginTrans
            MediConnCISWS.beginTrans
            medixmasterconn.beginTrans
            
            
            strsqlLG = ""
            strsqlLG = "EXEC SaveLG2016 '" & ChkStr(LGNo) & "', '" & ChkStr(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1)) & "', " & _
                        "'" & Logon & "', '" & Format(TreatmentDate, "yyyy/mm/dd") & "','" & ChkStr(txtDoctor) & "', '" & Clinic & "', " & _
                        "'" & Right(txtMBMNo, Len(txtMBMNo) - InStr(txtMBMNo, "-")) & "', '" & txtRemarks & "', " & _
                        "'" & Logon & "', '" & txtTotalBill & "', '" & txtRBExcess & "', '" & txtOtherExcess & "', " & _
                        "'" & txtRBCoPay & "', '" & txtLmtExceeded & "', '" & txtTotalExcess & "', '" & txtGLAmt & "','" & Trim(txtDiagnosis) & "','" & Trim(txtInternalRemarks) & "','" & cmbGLType & "'"
            MediConnCISDB.Execute strsqlLG
            strsqlLG = ""
            strsqlLG = "Update TblRecord Set IsLGissued =1 , IsRejected=0,CaseProgress='LG Issued' where CaseNo='" + LGNo + "'"
            MediConnCISDB.Execute strsqlLG
        'New Code to insert values into claim table
        strsqlLG = ""
        strsqlLG = "Update Claim  Set ClaimOpenDate='" + CStr(Format(Date, "yyyy/mm/dd")) + "',ClaimOpenBy='" + Logon + "', ClaimCloseDate='" + CStr(Format(Date, "yyyy/mm/dd")) + "',ClaimClosedBy='" + Logon + "'," & _
                    " CLMTotalActual ='" + txtGLAmt + "',CLMFinalApprove='" + txtGLAmt + "',CLMPaidBYMedix2C='" + txtGLAmt + "',CLMTotalApproved='" + txtGLAmt + "',CLMBordxAmt='" + txtGLAmt + "' From claim  where  ClaimBordxStatus=0 and CASNumber='" + LGNo + "'"
        MediConnCISWS.Execute strsqlLG
        
        strsqlLG = ""
        strsqlLG = "Update CISCas  Set  CasDelayReason='" + Trim(cmdDelayreason) + "', DateVisit='" + CStr(Format(txtDate, "yyyy/mm/dd")) + "' , CASStatus='O',Claimability='A' From CISCas  where CASNumber='" + LGNo + "'"
        MediConnCISDB.Execute strsqlLG
        
        strsqlLG = ""
        strsqlLG = "Update ClaimInvoice  Set INVAmount='" + txtGLAmt + "',INVBillAmt='" + txtGLAmt + "' From ClaimInvoice  where InvBordxed=0 and CasNumber='" + LGNo + "'"
        MediConnCISWS.Execute strsqlLG
            
            
            MediConnCISDB.commitTrans
            MediConnCISWS.commitTrans
            medixmasterconn.commitTrans
            startTrans = False
            
            lblLG = LGNo
            MsgBox "Record Saved....", vbOKOnly + vbExclamation
        ElseIf tmpPymtStatus = False Then
            startTrans = True
            MediConnCISDB.beginTrans
            MediConnCISWS.beginTrans
            
            
            LGNo = lstLG.SelectedItem.Text
            tmpmaternityamt = 0
            txtmaternitynormal = IIf(Len(Trim(txtmaternitynormal)), Trim(txtmaternitynormal), 0)
            txtCaesarean = IIf(Len(Trim(txtCaesarean)), Trim(txtCaesarean), 0)
            txtComplications = IIf(Len(Trim(txtComplications)), Trim(txtComplications), 0)
            
            If cmbGLType.Text = "Inpatient Delivery Private" Then
                If txtmaternitynormal >= txtCaesarean Then
                    tmpmaternityamt = txtCaesarean
                ElseIf txtmaternitynormal <= txtCaesarean Then
                    tmpmaternityamt = txtmaternitynormal
                End If
            ElseIf cmbGLType.Text = "Inpatient Delivery Private-Complications" Then
                tmpmaternityamt = IIf(Len(txtComplications), txtComplications, 0)
            End If
            
            strsqlLG = ""
            If cmbGLType.Text = "Inpatient Delivery Private" Or cmbGLType.Text = "Inpatient Delivery Private-Complications" Then
                strsqlLG = "Update LG  Set LGDate='" + CStr(Format(TreatmentDate, "yyyy/mm/dd")) + "',LGIssuedBy='" + Logon + "', CLNName='" + CStr(Clinic) + "',LGRemarks='" + txtRemarks + "'," & _
                            " Diagnosis ='" + IIf(Len(Trim(txtDiagnosis)), Trim(txtDiagnosis), "") + "',GLAmt='" + Format(tmpmaternityamt, "#,##0.00") + "', InternalRemarks='" + txtInternalRemarks + "',DoctorName='" + ChkStr(txtDoctor) + "',GLType='" + cmbGLType + "', LGIssueDate = GETDATE() From LG  where LGNo='" + LGNo + "'"
            Else
                strsqlLG = "Update LG  Set LGDate='" + CStr(Format(TreatmentDate, "yyyy/mm/dd")) + "',LGIssuedBy='" + Logon + "', CLNName='" + CStr(Clinic) + "',LGRemarks='" + txtRemarks + "'," & _
                            " Diagnosis ='" + IIf(Len(Trim(txtDiagnosis)), Trim(txtDiagnosis), "") + "',GLAmt='" + txtGLAmt + "', InternalRemarks='" + txtInternalRemarks + "',DoctorName='" + ChkStr(txtDoctor) + "',GLType='" + cmbGLType + "', LGIssueDate = GETDATE() From LG  where LGNo='" + LGNo + "'"
            End If
            MediConnCISDB.Execute strsqlLG
            
            strsqlLG = ""
            If cmbGLType.Text = "Inpatient Delivery Private" Or cmbGLType.Text = "Inpatient Delivery Private-Complications" Then
                strsqlLG = "Update Claim  Set ClaimOpenDate='" + CStr(Format(Date, "yyyy/mm/dd")) + "',ClaimOpenBy='" + Logon + "', ClaimCloseDate='" + CStr(Format(Date, "yyyy/mm/dd")) + "',ClaimClosedBy='" + Logon + "'," & _
                            " CLMTotalActual ='" + Format(tmpmaternityamt, "#,##0.00") + "',CLMFinalApprove='" + Format(tmpmaternityamt, "#,##0.00") + "',CLMPaidBYMedix2C='" + Format(tmpmaternityamt, "#,##0.00") + "',CLMTotalApproved='" + Format(tmpmaternityamt, "#,##0.00") + "',CLMBordxAmt='" + Format(tmpmaternityamt, "#,##0.00") + "' From claim  where ClaimBordxStatus=0 and CASNumber='" + LGNo + "'"
            Else
                strsqlLG = "Update Claim  Set ClaimOpenDate='" + CStr(Format(Date, "yyyy/mm/dd")) + "',ClaimOpenBy='" + Logon + "', ClaimCloseDate='" + CStr(Format(Date, "yyyy/mm/dd")) + "',ClaimClosedBy='" + Logon + "'," & _
                            " CLMTotalActual ='" + txtGLAmt + "',CLMFinalApprove='" + txtGLAmt + "',CLMPaidBYMedix2C='" + txtGLAmt + "',CLMTotalApproved='" + txtGLAmt + "',CLMBordxAmt='" + txtGLAmt + "' From claim  where ClaimBordxStatus=0 and CASNumber='" + LGNo + "'"
            End If
            MediConnCISWS.Execute strsqlLG
            
            strsqlLG = ""
            If tmpIsPending = True And (cmbGLType.Text = "Inpatient Delivery Private" Or cmbGLType.Text = "Inpatient Delivery Private-Complications" Or cmbGLType.Text = "Inpatient Delivery (GOV)") Then
                strsqlLG = "Update CISDB..CISCas  Set CasDelayReason='" + Trim(cmdDelayreason) + "',Pending='N',CASTreatReason1='302',PendingReminderDate=NULL, CASRepAction=CASRepAction+'|'+convert(nvarchar,getdate(),103) +' '+ convert(varchar, getdate(), 108)+'-',DateVisit='" + CStr(Format(txtDate, "yyyy/mm/dd")) + "',Claimability='A' From CISCas  where CASNumber='" + LGNo + "'"
            ElseIf tmpIsPending = True Then
                strsqlLG = "Update CISDB..CISCas  Set CasDelayReason='" + Trim(cmdDelayreason) + "',Pending='N',CASTreatReason1=NULL,PendingReminderDate=NULL, CASRepAction=CASRepAction+'|'+convert(nvarchar,getdate(),103) +' '+ convert(varchar, getdate(), 108)+'-',DateVisit='" + CStr(Format(txtDate, "yyyy/mm/dd")) + "',Claimability='A' From CISCas  where CASNumber='" + LGNo + "'"
            ElseIf cmbGLType.Text = "Inpatient Delivery Private" Or cmbGLType.Text = "Inpatient Delivery Private-Complications" Or cmbGLType.Text = "Inpatient Delivery (GOV)" Then
                strsqlLG = "Update CISDB..CISCas  Set CasDelayReason='" + Trim(cmdDelayreason) + "',DateVisit='" + CStr(Format(txtDate, "yyyy/mm/dd")) + "',CASTreatReason1='302',Claimability='A' From CISCas  where CASNumber='" + LGNo + "'"
            Else
                strsqlLG = "Update CISDB..CISCas  Set CasDelayReason='" + Trim(cmdDelayreason) + "',DateVisit='" + CStr(Format(txtDate, "yyyy/mm/dd")) + "',CASTreatReason1=NULL,Claimability='A' From CISCas  where CASNumber='" + LGNo + "'"
            End If
            MediConnCISDB.Execute strsqlLG
            
            strsqlLG = ""
            strsqlLG = "Update TblRecord Set IsLGissued =1 , IsRejected=0,CaseProgress='LG Issued' where CaseNo='" + LGNo + "'"
            MediConnCISDB.Execute strsqlLG
            
            strsqlLG = ""
            strsqlLG = "Update ClaimInvoice  Set INVAmount='" + txtGLAmt + "',INVBillAmt='" + txtGLAmt + "' From ClaimInvoice  where InvBordxed=0 and CasNumber='" + LGNo + "'"
            MediConnCISWS.Execute strsqlLG
            
            strsqlLG = ""
            If cmbGLType.Text = "Inpatient Delivery Private" Or cmbGLType.Text = "Inpatient Delivery Private-Complications" Then
                strsqlLG = "Update ClaimInvoiceActualNew set CLMMaternity='" + Format(tmpmaternityamt, "#,##0.00") + "' where CASNumber='" + LGNo + "'"
            Else
                strsqlLG = "Update ClaimInvoiceActualNew set CLMMaternity='" + Format(0, "#,##0.00") + "' where CASNumber='" + LGNo + "'"
            End If
            MediConnCISWS.Execute strsqlLG
            
            
            MediConnCISDB.commitTrans
            MediConnCISWS.commitTrans
            startTrans = False
            
            MsgBox "Record Updated....", vbOKOnly + vbExclamation
            cmdLG.Caption = "&LG"
            LGNo = lblLG
            Call funcwordNew
        End If
        Call LGListing
        If tmpEmailLG <> "1" And tmpPymtStatus = False Then
                            Dim strsqlPLN As String
                            Dim PLN As New ADODB.Recordset
                            Dim Cnt As Integer
                            Dim Mailtype As String
                            strsqlPLN = ""
                            strsqlPLN = "Select Count(*) as Cnt,Mailtype from HISMaintenance..Tbl_CISEmailMaster where coalesce(Mailtype,'')<>'E' and coalesce(IsCIS,'')='Y' and  Grpcompany='" + Trim(Replace(txtGrpCompany, "'", "''")) + "' Group by Mailtype"
                            PLN.Open strsqlPLN, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
                            If Not PLN.EOF Then
                                Cnt = IIf(Len(PLN!Cnt), CInt(PLN!Cnt), 0)
                                Mailtype = PLN!Mailtype
                            End If
                            PLN.Close
                            Set PLN = Nothing
                            If Cnt >= 1 Then
                                Dim subsidary As String
                                Dim CostCenter As String
                                Dim Department As String
                                Dim sArr() As String
                                sArr = Split(txtMBMNo, "-")
                                txtMBMNo = sArr(0)
                                subsidary = ""
                                CostCenter = ""
                                Department = ""
                                strsqlPLN = ""
                                strsqlPLN = "Select top 1  MBMGroupCompany, Subsidiary , MBMCostCenter , MBMDepartment   from MBMOthers where MBMNumber ='" & Trim(txtMBMNo) & "'"
                                PLN.Open strsqlPLN, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                                If Not PLN.EOF Then
                                    subsidary = IIf(Len(PLN!Subsidiary), (PLN!Subsidiary), "")
                                    CostCenter = IIf(Len(PLN!MBMCostCenter), (PLN!MBMCostCenter), "")
                                    Department = IIf(Len(PLN!MBMDepartment), (PLN!MBMDepartment), "")
                                End If
                                PLN.Close
                                Set PLN = Nothing
                                
                            If Mailtype <> "G" And Cnt > 1 Then
                                'tmpMBMNum
                                Dim Email As String
                                Dim BCC As String
                                Dim CC As String
                                
                                strsqlPLN = ""
                                If Trim(Mailtype) = Trim("G") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsCIS='Y' and Grpcompany='" + Trim(Replace(txtGrpCompany, "'", "''")) + "'"
                                ElseIf Trim(Mailtype) = Trim("S") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsCIS='Y' and Grpcompany='" + Trim(Replace(txtGrpCompany, "'", "''")) + "' and  Subsidary='" + Trim(Replace(subsidary, "'", "''")) + "'"
                                ElseIf Trim(Mailtype) = Trim("D") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsCIS='Y' and Grpcompany='" + Trim(Replace(txtGrpCompany, "'", "''")) + "' and Department='" + Trim(Replace(Department, "'", "''")) + "'"
                                ElseIf Trim(Mailtype) = Trim("C") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsCIS='Y' and Grpcompany='" + Trim(Replace(txtGrpCompany, "'", "''")) + "' and Costcenter='" + Trim(Replace(CostCenter, "'", "''")) + "'"
                                End If
                                
                                PLN.Open strsqlPLN, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
                                If Not PLN.EOF Then
                                    Email = IIf(Len(PLN!EmailAddress), PLN!EmailAddress, "")
                                    BCC = IIf(Len(PLN!BCC), PLN!BCC, "")
                                    CC = IIf(Len(PLN!CC), PLN!CC, "")
                                End If
                                PLN.Close
                                Set PLN = Nothing
                                If Len(Email) Then
                                Call MailAdmInitialLG(Email, BCC, CC, LGNo, subsidary)
                                End If
                                Else
                                strsqlPLN = ""
                                If Trim(Mailtype) = Trim("G") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsCIS='Y' and Grpcompany='" + Trim(Replace(txtGrpCompany, "'", "''")) + "'"
                                ElseIf Trim(Mailtype) = Trim("S") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsCIS='Y' and Grpcompany='" + Trim(Replace(txtGrpCompany, "'", "''")) + "' and  Subsidary='" + Replace(subsidary, "'", "''") + "'"
                                ElseIf Trim(Mailtype) = Trim("D") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsCIS='Y' and Grpcompany='" + Trim(Replace(txtGrpCompany, "'", "''")) + "' and Department='" + Replace(Department, "'", "''") + "'"
                                ElseIf Trim(Mailtype) = Trim("C") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsCIS='Y' and Grpcompany='" + Trim(Replace(txtGrpCompany, "'", "''")) + "' and Costcenter='" + Replace(CostCenter, "'", "''") + "'"
                                End If
                                
                                PLN.Open strsqlPLN, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
                                If Not PLN.EOF Then
                                    Email = IIf(Len(PLN!EmailAddress), PLN!EmailAddress, "")
                                    BCC = IIf(Len(PLN!BCC), PLN!BCC, "")
                                    CC = IIf(Len(PLN!CC), PLN!CC, "")
                                End If
                                PLN.Close
                                Set PLN = Nothing
                                If Len(Email) Then
                                Call MailAdmInitialLG(Email, BCC, CC, LGNo, subsidary)
                                End If
                                End If
                
                End If
        End If
            'Call funcword
        lblLG = ""
        EditFlag = False
        AddFlag = False
        cmdEdit.Enabled = False
        cmdLG.Enabled = False
        cmdPreview.Enabled = False
        cboCLNCode.Enabled = False
        chkCLN.Enabled = False
        txtDoctor.Enabled = False
        txtDate.Enabled = False
        txtRemarks.Enabled = False
        txtDiagnosis.Enabled = False
        cboIndicator.Enabled = False
    End If
    
    Else
        MsgBox ("Please Check NO GLTemplate is Assigned for Plan Code : " + Trim(lblPLNCode) + ". Please Check LG Maintainace Table")
    End If
    
debugs:
If startTrans = True Then
    MediConnCISDB.RollbackTrans
    MediConnCISWS.RollbackTrans
End If

If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub ValidateLG1()
Const METHOD_NAME = "ValidateLG1"
On Error GoTo debugs

Dim SqlQry As String
Dim AnnSpecLimitrst2 As New ADODB.Recordset
Dim tmpGLFormat As String
Dim VipInd1 As String
Dim tmpclncode As String

    TemplateName = ""
    SuppTempName = ""
    SuppInd = ""
    tmpCLNName = ""
    tmpclnadd1 = ""
    tmpclnadd2 = ""
    tmpCLNState = ""
    
    tmpclncode = ""
    
    
    If Len(cboCLNCode) Then
    
    tmpclncode = Right(cboCLNCode, Len(cboCLNCode) - InStr(Trim(cboCLNCode), "["))
    tmpclncode = Trim(tmpclncode)
    ''tmpclncode = Mid(tmpclncode, 1, 4)
    tmpclncode = Mid(Trim(tmpclncode), 1, Len(tmpclncode) - 1)
    
    SqlQry = "Select Top 1 CLNGLFormat,coalesce(CLNName,'') As CLNName,clnadd1,(coalesce(clnadd2,'') +','+ coalesce(clnadd3,'')) as clnadd2,(coalesce(CLNState,'')+','+coalesce(CLNPostCode,'')) As CLNState from CLN where CLNCode = '" & tmpclncode & "' "
    AnnSpecLimitrst2.Open SqlQry, medixmasterconn, adOpenForwardOnly, adLockOptimistic
    Do Until AnnSpecLimitrst2.EOF
            tmpGLFormat = IIf(Len(AnnSpecLimitrst2!CLNGLFormat), Trim(AnnSpecLimitrst2!CLNGLFormat), "")
            tmpCLNName = IIf(Len(AnnSpecLimitrst2!CLNName), AnnSpecLimitrst2!CLNName, "")
            tmpclnadd1 = IIf(Len(AnnSpecLimitrst2!CLNAdd1), AnnSpecLimitrst2!CLNAdd1, "")
            tmpclnadd2 = IIf(Len(AnnSpecLimitrst2!CLNAdd2), AnnSpecLimitrst2!CLNAdd2, "")
            tmpCLNState = IIf(Len(AnnSpecLimitrst2!CLNState), AnnSpecLimitrst2!CLNState, "")
        AnnSpecLimitrst2.MoveNext
    Loop
    AnnSpecLimitrst2.Close
    Set AnnSpecLimitrst2 = Nothing
    
    If tmpGLFormat = "M" Then
        tmpGLFormat = tmpGLFormat
    Else
        tmpGLFormat = "H"
    End If


        Dim Balcmd As New ADODB.Command
        Dim BalPrm1, BalPrm2, BalPrm3, BalPrm4, BalPrm5, BalPrm6, BalPrm7 As New ADODB.Parameter
        Dim BalRes As New ADODB.Recordset
        Set Balcmd.ActiveConnection = MediConnCISDB
        Balcmd.CommandText = "Sp_Gettemplatedetails_2"
        Balcmd.CommandType = adCmdStoredProc
        Balcmd.CommandTimeout = 9000
        
        Set BalPrm1 = Balcmd.CreateParameter("CompanyName", adVarChar, adParamInput, 2000, Trim(txtGrpCompany))
        Balcmd.Parameters.Append BalPrm1
        
        Set BalPrm2 = Balcmd.CreateParameter("plncode", adVarChar, adParamInput, 2000, Trim(lblPLNCode))
        Balcmd.Parameters.Append BalPrm2
        
        Set BalPrm3 = Balcmd.CreateParameter("GLFormat", adVarChar, adParamInput, 2000, Trim(tmpGLFormat))
        Balcmd.Parameters.Append BalPrm3
        
        Set BalPrm4 = Balcmd.CreateParameter("GLType", adVarChar, adParamInput, 2000, Trim(cmbGLType))
        Balcmd.Parameters.Append BalPrm4
        
        Set BalPrm5 = Balcmd.CreateParameter("IC", adVarChar, adParamInput, 2000, Trim(Replace(tmpicno, "-", "")))
        Balcmd.Parameters.Append BalPrm5
        
        Set BalPrm6 = Balcmd.CreateParameter("PEMC", adVarChar, adParamInput, 2000, Trim(tmpMBMCategory))
        Balcmd.Parameters.Append BalPrm6
        
        Set BalPrm7 = Balcmd.CreateParameter("caseno", adVarChar, adParamInput, 2000, Trim(LGNum))
        Balcmd.Parameters.Append BalPrm7

        BalRes.CursorLocation = adUseClient
        BalRes.CursorType = adOpenForwardOnly
        BalRes.CacheSize = 100
        Set BalRes = Balcmd.Execute
        
        If BalRes.EOF = False Then
            TemplateName = Trim(IIf(Len(BalRes!GLTemplate), BalRes!GLTemplate, ""))
            SuppTempName = Trim(IIf(Len(BalRes!DepTemplate), BalRes!DepTemplate, ""))
            SuppInd = Trim(IIf(Len(BalRes!DepInd), BalRes!DepInd, ""))
            VipInd = Trim(IIf(Len(BalRes!VipInd), BalRes!VipInd, ""))
            TmpRemarks = Trim(IIf(Len(BalRes!TmpRemarks), BalRes!TmpRemarks, ""))
            If txtRemarks <> TmpRemarks And Trim(tmpGLtype) <> Trim(cmbGLType) Then
                txtRemarks = TmpRemarks
            ElseIf Trim(tmpGLtype) = Trim(cmbGLType) And Len(tmpglpreviousremark) Then
                txtRemarks = tmpglpreviousremark
            Else
                txtRemarks = TmpRemarks
            End If
        Else
            txtRemarks = ""
        End If
      End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub MBMlisting()
Dim Record As ListItem
        Set Record = lstMBM.ListItems.Add(, , txtMBMNo)
        If Len(txtPatient) Then
            Record.SubItems(1) = txtPatient
        End If
        If Len(txtICNo) Then
            Record.SubItems(2) = txtICNo
        End If
        lblCount = 1
End Sub



Private Sub List1_Click()

On Error Resume Next
Dim lRet    As Long
Dim sFile   As String
    
  '  PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & ""
   
    crt PictureFileName & "\" & List1.List(List1.ListIndex)
        sFile = PictureFileName & "\" & List1.List(List1.ListIndex)
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)


End Sub

Private Sub lstCaseList_DblClick()
Const METHOD_NAME = "lstCaseList_DblClick"
On Error GoTo debugs

    If lstCaseList.ListItems.Count > 0 Then
        frmInvoiceAdj.txtCaseNo = Trim(lstCaseList.SelectedItem)
        Call frmInvoiceAdj.CmdGo_Click
        Load frmInvoiceAdj
        frmInvoiceAdj.Show
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstLG_Click()
Const METHOD_NAME = "lstLG_Click"
On Error GoTo debugs
Dim rst As New ADODB.Recordset
Dim StrSql As String
tmpGLtype = ""
tmpglpreviousremark = ""
''''''''''''''''''''''''''''''''''''''''''
Dim AnnSpecLimitcmd As New ADODB.Command
Dim AnnSpecLimitPrm1 As New ADODB.Parameter
Dim AnnSpecLimitPrm2 As New ADODB.Parameter
Dim AnnSpecLimitrst2 As New ADODB.Recordset
'''''''''''''''''''''''''''''''''''''''''''
    TempLG = ""
    tmpLGCasDate = ""
    tmpLGCasTime = ""
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    cmbDecline = ""
    
        'SRI START:check for pending case
            Call CheckForPending
            If tmpIsPending = True Then
                MsgBox "The case is still in pending state." & vbNewLine & "If you proceed to issue GL, the pending status is turned back to 'Accepted' and the selected pending reasons will be cleared.", vbInformation, "Case Pending"
            End If
        'SRI END:check for pending case
    
        If Len(lstLG.SelectedItem.SubItems(15)) And lstLG.ListItems.Count <> 0 Then
            MsgBox lstLG.SelectedItem.SubItems(15), vbOKOnly + vbInformation, DialogBoxTitle
        End If
    
    If lstLG.ListItems.Count <> 0 Then
        'cmdEdit.Enabled = True
        If TempMBM = False Then
            StrSql = " Select * From View_PrintLG Where LGNo = '" & lstLG.SelectedItem.Text & "'"
            rst.Open StrSql, MediConnCISDB, adOpenDynamic, adLockOptimistic
            Do Until rst.EOF
                If rst!coverid = rst!MBMCOVID Then
                    txtMBMNo = rst!MBMNumber & "-" & rst!MBMCOVID
                    txtICNo = IIf(Len(rst!MBMIcBcPp), rst!MBMIcBcPp, "")
                    txtPatient = IIf(Len(rst!MBMName), rst!MBMName, "")
                    'tmpMBMName = rst!MBMName
                    tmpglpreviousremark = IIf(Len(rst!LGRemarks), rst!LGRemarks, "")
                    txtRemarks = IIf(Len(rst!LGRemarks), rst!LGRemarks, "")
                    txtInternalRemarks = IIf(Len(rst!InternalRemarks), rst!InternalRemarks, "")
                    tmpGLtype = IIf(Len(rst!GLType), rst!GLType, "")
                    cmbGLType = IIf(Len(rst!GLType), rst!GLType, "")
                    
                    If Len(rst!DeclineReason) Then
                        cmbDecline = IIf(Len(rst!DeclineReason), rst!DeclineReason, "")
                    Else
                        cmbDecline.Clear
                         Call BindDecline
                    End If
                    Exit Do
                ElseIf rst!coverid = rst!MBMCCoverID Then
                    txtMBMNo = rst!MBMNumber & "-" & rst!MBMCCoverID
                    txtICNo = IIf(Len(rst!MBMCICBCPPNo), rst!MBMCICBCPPNo, "")
                    txtPatient = IIf(Len(rst!MBMCName), rst!MBMCName, "")
                    'tmpMBMName = rst!MBMName
                    'Test code For GL Remarks
                    txtRemarks = IIf(Len(rst!LGRemarks), rst!LGRemarks, "")
                    tmpglpreviousremark = IIf(Len(rst!LGRemarks), rst!LGRemarks, "")
                    txtInternalRemarks = IIf(Len(rst!InternalRemarks), rst!InternalRemarks, "")
                    tmpGLtype = IIf(Len(rst!GLType), rst!GLType, "")
                    cmbGLType = IIf(Len(rst!GLType), rst!GLType, "")
                    If Len(rst!DeclineReason) Then
                        cmbDecline = IIf(Len(rst!DeclineReason), rst!DeclineReason, "")
                    Else
                        cmbDecline.Clear
                        Call BindDecline
                    End If
                    Exit Do
                End If
                
                
                
                rst.MoveNext
            Loop
            rst.Close
            Set rst = Nothing
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Set AnnSpecLimitcmd.ActiveConnection = MediConnCISWS
            AnnSpecLimitcmd.CommandText = "GetLGAnnLimitSpecLimit"
            AnnSpecLimitcmd.CommandType = adCmdStoredProc
            AnnSpecLimitcmd.CommandTimeout = 9000
            Set AnnSpecLimitPrm1 = AnnSpecLimitcmd.CreateParameter("SchMBMNumber", adVarChar, adParamInput, 2000, Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1)))
            AnnSpecLimitcmd.Parameters.Append AnnSpecLimitPrm1
            Set AnnSpecLimitPrm2 = AnnSpecLimitcmd.CreateParameter("CoverId", adVarChar, adParamInput, 2000, Right(txtMBMNo, Len(txtMBMNo) - InStr(txtMBMNo, "-")))
            AnnSpecLimitcmd.Parameters.Append AnnSpecLimitPrm2
            AnnSpecLimitrst2.CursorLocation = adUseClient
            AnnSpecLimitrst2.CursorType = adOpenForwardOnly
            AnnSpecLimitrst2.CacheSize = 100
            Set AnnSpecLimitrst2 = AnnSpecLimitcmd.Execute
            
            Do While Not AnnSpecLimitrst2.EOF
                With AnnSpecLimitrst2
                    lblAnnLmt = IIf(IsNumeric(Trim(AnnSpecLimitrst2!OriAnnLimit)), Format(Trim(AnnSpecLimitrst2!OriAnnLimit), "#,##0.00"), Trim(AnnSpecLimitrst2!OriAnnLimit))
                    
                     If ASOInd = "Y" Then
                        lblAccumAmt = UtilisedAmt
                     Else
                        lblAccumAmt = IIf(IsNumeric(Trim(AnnSpecLimitrst2!GUtilized)), Format(Trim(AnnSpecLimitrst2!GUtilized), "#,##0.00"), Trim(AnnSpecLimitrst2!GUtilized))
                     End If
                    
                    lblBalLmt = IIf(IsNumeric(Trim(AnnSpecLimitrst2!GBalance)), Format(Trim(AnnSpecLimitrst2!GBalance), "#,##0.00"), Trim(AnnSpecLimitrst2!GBalance))
                    lblSpecAnn = IIf(IsNumeric(Trim(AnnSpecLimitrst2!OriSpecAnnLimit)), Format(Trim(AnnSpecLimitrst2!OriSpecAnnLimit), "#,##0.00"), Trim(AnnSpecLimitrst2!OriSpecAnnLimit))
                    lblSpecAcc = IIf(IsNumeric(Trim(AnnSpecLimitrst2!SUtilized)), Format(Trim(AnnSpecLimitrst2!SUtilized), "#,##0.00"), Trim(AnnSpecLimitrst2!SUtilized))
                    lblSpecBal = IIf(IsNumeric(Trim(AnnSpecLimitrst2!SBalance)), Format(Trim(AnnSpecLimitrst2!SBalance), "#,##0.00"), Trim(AnnSpecLimitrst2!SBalance))
                    lblTopUp = IIf(IsNumeric(Trim(AnnSpecLimitrst2!TopUpAmt)), Format(Trim(AnnSpecLimitrst2!TopUpAmt), "#,##0.00"), Trim(AnnSpecLimitrst2!TopUpAmt))
                    lblNoAnnLimtDed = IIf(IsNumeric(Trim(AnnSpecLimitrst2!GNoLimitDedTotal)), Format(Trim(AnnSpecLimitrst2!GNoLimitDedTotal), "#,##0.00"), Trim(AnnSpecLimitrst2!GNoLimitDedTotal))
                    lblNoSpecLimtDed = IIf(IsNumeric(Trim(AnnSpecLimitrst2!SNoLimitDedTotal)), Format(Trim(AnnSpecLimitrst2!SNoLimitDedTotal), "#,##0.00"), Trim(AnnSpecLimitrst2!SNoLimitDedTotal))
                    lblPLNCode = IIf(IsNumeric(Trim(AnnSpecLimitrst2!PLNCode)), Format(Trim(AnnSpecLimitrst2!PLNCode), "###"), Trim(AnnSpecLimitrst2!PLNCode))
                    
                    If Len(Trim(AnnSpecLimitrst2!LGIssuedSuppInd)) Then
                        lblLGIssuedInd = Trim(AnnSpecLimitrst2!LGIssuedSuppInd)
                    Else
                        lblLGIssuedInd = "N"
                    End If
                    If Len(Trim(AnnSpecLimitrst2!LGIssuedInd)) Then
                        lblLGIssuedInd = Trim(AnnSpecLimitrst2!LGIssuedInd)
                    Else
                        lblLGIssuedInd = "N"
                    End If
                    If Len(Trim(AnnSpecLimitrst2!LGIssuedAfterLimitExInd)) Then
                        lblLGIssuedAfterLimitExInd = Trim(AnnSpecLimitrst2!LGIssuedAfterLimitExInd)
                    Else
                        lblLGIssuedAfterLimitExInd = "N"
                    End If
                    If Len(Trim(AnnSpecLimitrst2!LGUsingLimit)) Then
                        lblUsingLimit = Trim(AnnSpecLimitrst2!LGUsingLimit)
                    Else
                        lblUsingLimit = "N/A"
                    End If
                    If Len(Trim(AnnSpecLimitrst2!LGUsingLimitDesc)) Then
                        lblUsingLimitDesc = Trim(AnnSpecLimitrst2!LGUsingLimitDesc)
                    Else
                        lblUsingLimitDesc = "N/A"
                    End If
                    If Trim(AnnSpecLimitrst2!error) <> "N/A" Then
                        MsgBox AnnSpecLimitrst2!error, vbOKOnly + vbCritical
                    End If
                End With
                AnnSpecLimitrst2.MoveNext
            Loop
            
            AnnSpecLimitrst2.Close
            Set AnnSpecLimitrst2 = Nothing
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
        LGNum = lstLG.SelectedItem.Text
        TempLG = lstLG.SelectedItem.Text
        lblLG = lstLG.SelectedItem.Text
        txtDoctor = lstLG.SelectedItem.SubItems(4)
        If Len(lstLG.SelectedItem.SubItems(2)) Then
            txtDate = lstLG.SelectedItem.SubItems(2)
        Else
            txtDate = Format(DateTime.Now, "DD/MM/YYYY")
        End If
        
        If Len(lstLG.SelectedItem.SubItems(11)) Then
            tmpglpreviousremark = lstLG.SelectedItem.SubItems(11)
            txtRemarks = lstLG.SelectedItem.SubItems(11)
        Else
            tmpglpreviousremark = TmpRemarks
            txtRemarks = TmpRemarks
        End If
        
        txtInternalRemarks = lstLG.SelectedItem.SubItems(14)
        
        cboCLNCode = Trim(lstLG.SelectedItem.SubItems(9)) & "[" & Trim(lstLG.SelectedItem.SubItems(5)) & "]"
        
        'txtRemarks = lstLG.SelectedItem.SubItems(6)
        txtDiagnosis = lstLG.SelectedItem.SubItems(7)
        
        If Len(lstLG.SelectedItem.SubItems(13)) Then
           tmpGLtype = lstLG.SelectedItem.SubItems(13)
           cmbGLType.Text = lstLG.SelectedItem.SubItems(13)
        Else
            'cmbGLType = ""
        End If
        
        cmdDelayreason = lstLG.SelectedItem.SubItems(16)
        tmpLGCasDate = lstLG.SelectedItem.SubItems(2)
        tmpLGCasTime = lstLG.SelectedItem.SubItems(17)
        
        
        If cmbGLType.Text = "Inpatient Delivery Private" Then
            txtmaternitynormal = lstLG.SelectedItem.SubItems(6)
        ElseIf cmbGLType.Text = "Inpatient Delivery Private-Complications" Then
            txtComplications = lstLG.SelectedItem.SubItems(6)
        Else
            txtGLAmt = lstLG.SelectedItem.SubItems(6)
        End If
        
        tmpGLAmt_Chk_update = IIf(IsNumeric(lstLG.SelectedItem.SubItems(6)), lstLG.SelectedItem.SubItems(6), 0)
        'tmpGLFormat = lstLG.SelectedItem.SubItems(10)
        If AddFlag = True Then
            If Mid(cboIndicator, 1, 1) = "F" Then
                cboIndicator.ListIndex = 1
            Else
                cboIndicator.ListIndex = 0
            End If
        Else
            'If Mid(lstLG.SelectedItem.Text, Len(lstLG.SelectedItem.Text), 1) = 1 Then
            If Mid(cboIndicator, 1, 1) = "F" Then
                cboIndicator.ListIndex = 0
            Else
                cboIndicator.ListIndex = 1
            End If
        End If
    Else
        lstLG.Enabled = False
    End If
    List1.Clear
    Call CheckIsPharmacy(lstLG.SelectedItem.Text)
    If tmpMBMCategory = "PEMC" Then
        Call PEMCdata(LGNum)
    Else
        tmpPatientName = ""
        tmpicno = ""
    End If
    If AddFlag = False And Mid(cboIndicator, 1, 1) <> "F" Then
        cboCLNCode.Enabled = False
        chkCLN.Enabled = False
        txtDoctor.Enabled = False
        txtDate.Enabled = False
        txtRemarks.Enabled = False
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        txtDiagnosis.Enabled = False
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        cboIndicator.Enabled = False
     Else
         'Code to add call docs
         LGNum = lstLG.SelectedItem.Text
         List1.Clear
         Call addTheFilesClaims(LGNum)
         
         'Get LG Template Name
         Call ValidateLG1
        
        'If Len(TemplateName) Then
        '     txtRemarks = IIf(Len(TmpRemarks), TmpRemarks, "")
        'Else
        '     txtRemarks = ""
        'End If
    End If
    
    'To disabled the Edit if the medication request order is pending in pharmacies side
    If IsPharmacy Then
        If IsMedRequestInProgress(lstLG.SelectedItem.Text) Then
            cmdDecline.Enabled = False
            cmdLG.Enabled = False
            cmdEdit.Enabled = False
            cmdCopyFile.Enabled = False
        Else
            cmdDecline.Enabled = True
            cmdLG.Enabled = True
            cmdEdit.Enabled = True
            If Not IsMasterCase Then
                cmdCopyFile.Enabled = True
            End If
        End If
    Else
        cmdEdit.Enabled = True
    End If
    Me.MousePointer = vbDefault
    Me.Enabled = True
debugs:
Me.MousePointer = vbDefault
Me.Enabled = True
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub PEMCdata(CasNumber As String)
Dim sql As String
Dim ado As New ADODB.Recordset
tmpPatientName = ""
tmpicno = ""
sql = "Select patientname,icno,* from cisdb..tblOfflineGL WHERE casno = '" & CasNumber & "' "
ado.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If Not ado.EOF Then
        tmpPatientName = IIf(Len(ado!PatientName), ado!PatientName, "")
        tmpicno = IIf(Len(ado!IcNo), Trim(Replace(ado!IcNo, "-", "")), "")
    End If
ado.Close
Set ado = Nothing
End Sub

Private Sub CheckIsPharmacy(CasNumber As String)
Dim objExec
Dim sql As String
Dim ado As New ADODB.Recordset

sql = "SELECT COUNT(*) AS TotalRecord FROM CISDB..LG INNER JOIN HISDB..CLN ON CLN.CLNCode = LG.CLNName WHERE LGNo = '" & CasNumber & "' AND CLN.CLNType = 'D'"
Set objExec = medixmasterconn.Execute(sql)
ado.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic

'Reset to default
IsMasterCase = False
IsPharmacy = False

If Not ado.EOF Then
    If objExec!totalrecord > 0 Then
        IsPharmacy = True
        ado.Close
        Set ado = Nothing
        sql = "SELECT MainCaseNo FROM CISDB..tblMedRequestFollowup WHERE FollowupCaseNo = '" & CasNumber & "'"
        Set objExec = medixmasterconn.Execute(sql)
        
        ado.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If Not ado.EOF Then
            If objExec!MainCaseNo <> "" Then
                CasNumber = objExec!MainCaseNo
                IsMasterCase = True
            End If
        End If
    End If
End If

ado.Close
Set ado = Nothing

End Sub

Private Function IsMedRequestInProgress(ByVal strCaseNo As String) As Boolean
Const METHOD_NAME = "IsMedRequestInProgress"
On Error GoTo debugs

Dim MRcmd As New ADODB.Command
Dim MRPrm1 As New ADODB.Parameter
Dim MRRes As New ADODB.Recordset
Dim sRetValue As Boolean

    IsMedRequestInProgress = False
    
    Set MRcmd.ActiveConnection = MediConnCISDB
    MRcmd.CommandText = "sp_HCMedRequestInProgress"
    MRcmd.CommandType = adCmdStoredProc
    MRcmd.CommandTimeout = 9000
    Set MRPrm1 = MRcmd.CreateParameter("CaseNo", adVarChar, adParamInput, 100, strCaseNo)
    MRcmd.Parameters.Append MRPrm1
    MRcmd.Parameters.Append MRcmd.CreateParameter("RetOut", adBoolean, adParamInputOutput, 1, 0)
    MRRes.CursorLocation = adUseClient
    MRRes.CursorType = adOpenForwardOnly
    MRRes.CacheSize = 100
    Set MRRes = MRcmd.Execute
    sRetValue = MRcmd.Parameters("RetOut")
    Set MRPrm1 = Nothing
    Set MRcmd = Nothing
    Set MRRes = Nothing
    
    IsMedRequestInProgress = sRetValue
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Sub addTheFilesClaims(CasNumber As String)
Dim B As String
'Dim PictureFileName As String
List1.Clear
PictureFileName = ""

If IsPharmacy Then
    PictureFileName = "\\Svr-doc\Outpatient\OtherScannedDoc\" & CasNumber
Else
    PictureFileName = "\\Svr-doc\Outpatient\Guarantee Letter\" & CasNumber
End If

B = Dir(PictureFileName & "\*.*")
If B <> "" Then
    If Right(B, 4) = ".jpg" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".tif" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".gif" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".JPG" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".TIF" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".GIF" Then
        List1.AddItem B
    End If
    
    If UCase(Right(B, 4)) = ".PDF" Then
        List1.AddItem B
    End If
    
    If UCase(Right(B, 4)) = ".BMP" Then
        List1.AddItem B
    End If

    If UCase(Right(B, 4)) = ".PNG" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".msg" Then
        List1.AddItem B
    End If

Else
    Exit Sub
End If
Do: DoEvents
    B = Dir
    If B <> "" Then
        If Right(B, 4) = ".jpg" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".tif" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".gif" Then
            List1.AddItem B
        End If
        
        If Right(B, 4) = ".JPG" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".TIF" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".GIF" Then
            List1.AddItem B
        End If
        
        If UCase(Right(B, 4)) = ".PDF" Then
            List1.AddItem B
        End If
        
        If UCase(Right(B, 4)) = ".BMP" Then
            List1.AddItem B
        End If
        
        If UCase(Right(B, 4)) = ".PNG" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".msg" Then
        List1.AddItem B
        End If
        
    Else
        Exit Do
    End If
Loop
End Sub

Private Sub lstLG_DblClick()
Const METHOD_NAME = "lstLG_DblClick"
On Error GoTo debugs
Dim tmpclncode As String
Dim strsqlCLN As String
Dim CLN As New ADODB.Recordset

    Letter = "R"
    strsqlCLN = ""
    tmpclncode = ""
    tmpclncode = Right(cboCLNCode, Len(cboCLNCode) - InStr(Trim(cboCLNCode), "["))
    tmpclncode = Trim(tmpclncode)
    ''tmpclncode = Mid(tmpclncode, 1, 4)
    tmpclncode = Mid(Trim(tmpclncode), 1, Len(tmpclncode) - 1)
    
    strsqlCLN = "Select Top 1 * from CLN where CLNCode = '" & tmpclncode & "' "
    CLN.Open strsqlCLN, medixmasterconn, adOpenForwardOnly, adLockOptimistic
    
    Do Until CLN.EOF
            tmpGLFormat = IIf(Len(CLN!CLNGLFormat), CLN!CLNGLFormat, "")
        CLN.MoveNext
    Loop
    CLN.Close
    Set CLN = Nothing
    
    Dim status As String
    status = Trim(lstLG.SelectedItem.SubItems(12))
    
    If Trim(status) = Trim("LG Issued") Then
        ''Call funcwordNew
    ElseIf Trim(status) = Trim("LG Declined") Then
        ''Call funcwordDec
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstLG_KeyDown(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "lstLG_KeyDown"
On Error GoTo debugs
Dim rst As New ADODB.Recordset
Dim StrSql As String
''''''''''''''''''''''''''''''''''''''''''
Dim AnnSpecLimitcmd As New ADODB.Command
Dim AnnSpecLimitPrm1 As New ADODB.Parameter
Dim AnnSpecLimitPrm2 As New ADODB.Parameter
Dim AnnSpecLimitrst2 As New ADODB.Recordset
'''''''''''''''''''''''''''''''''''''''''''
    
TempLG = ""
tmpLGCasDate = ""
tmpLGCasTime = ""

    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
        If Len(lstLG.SelectedItem.SubItems(15)) And lstLG.ListItems.Count <> 0 Then
            MsgBox lstLG.SelectedItem.SubItems(15), vbOKOnly + vbInformation, DialogBoxTitle
        End If

        If lstLG.ListItems.Count <> 0 Then
            'cmdEdit.Enabled = True
            If TempMBM = False Then
                StrSql = " Select * From View_PrintLG Where LGNo = '" & lstLG.SelectedItem.Text & "'"
                rst.Open StrSql, MediConnCISDB, adOpenDynamic, adLockOptimistic
                Do Until rst.EOF
                    If rst!coverid = rst!MBMCOVID Then
                        txtMBMNo = rst!MBMNumber & "-" & rst!MBMCOVID
                        txtICNo = rst!MBMIcBcPp
                        txtPatient = rst!MBMName
                        Exit Do
                    ElseIf rst!coverid = rst!MBMCCoverID Then
                        txtMBMNo = rst!MBMNumber & "-" & rst!MBMCCoverID
                        txtICNo = rst!MBMCICBCPPNo
                        txtPatient = rst!MBMCName
                        Exit Do
                    End If
                    rst.MoveNext
                Loop
                rst.Close
                Set rst = Nothing
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Set AnnSpecLimitcmd.ActiveConnection = MediConnCISWS
                AnnSpecLimitcmd.CommandText = "GetLGAnnLimitSpecLimit"
                AnnSpecLimitcmd.CommandType = adCmdStoredProc
                AnnSpecLimitcmd.CommandTimeout = 9000
                Set AnnSpecLimitPrm1 = AnnSpecLimitcmd.CreateParameter("SchMBMNumber", adVarChar, adParamInput, 2000, Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1)))
                AnnSpecLimitcmd.Parameters.Append AnnSpecLimitPrm1
                Set AnnSpecLimitPrm2 = AnnSpecLimitcmd.CreateParameter("CoverId", adVarChar, adParamInput, 2000, Right(txtMBMNo, Len(txtMBMNo) - InStr(txtMBMNo, "-")))
                AnnSpecLimitcmd.Parameters.Append AnnSpecLimitPrm2
                AnnSpecLimitrst2.CursorLocation = adUseClient
                AnnSpecLimitrst2.CursorType = adOpenForwardOnly
                AnnSpecLimitrst2.CacheSize = 100
                Set AnnSpecLimitrst2 = AnnSpecLimitcmd.Execute
                Do While Not AnnSpecLimitrst2.EOF
                    With AnnSpecLimitrst2
                        lblAnnLmt = IIf(IsNumeric(Trim(AnnSpecLimitrst2!OriAnnLimit)), Format(Trim(AnnSpecLimitrst2!OriAnnLimit), "#,##0.00"), Trim(AnnSpecLimitrst2!OriAnnLimit))
                        If ASOInd = "Y" Then
                           lblAccumAmt = UtilisedAmt
                        Else
                           lblAccumAmt = IIf(IsNumeric(Trim(AnnSpecLimitrst2!GUtilized)), Format(Trim(AnnSpecLimitrst2!GUtilized), "#,##0.00"), Trim(AnnSpecLimitrst2!GUtilized))
                        End If
                        lblBalLmt = IIf(IsNumeric(Trim(AnnSpecLimitrst2!GBalance)), Format(Trim(AnnSpecLimitrst2!GBalance), "#,##0.00"), Trim(AnnSpecLimitrst2!GBalance))
                        lblSpecAnn = IIf(IsNumeric(Trim(AnnSpecLimitrst2!OriSpecAnnLimit)), Format(Trim(AnnSpecLimitrst2!OriSpecAnnLimit), "#,##0.00"), Trim(AnnSpecLimitrst2!OriSpecAnnLimit))
                        lblSpecAcc = IIf(IsNumeric(Trim(AnnSpecLimitrst2!SUtilized)), Format(Trim(AnnSpecLimitrst2!SUtilized), "#,##0.00"), Trim(AnnSpecLimitrst2!SUtilized))
                        lblSpecBal = IIf(IsNumeric(Trim(AnnSpecLimitrst2!SBalance)), Format(Trim(AnnSpecLimitrst2!SBalance), "#,##0.00"), Trim(AnnSpecLimitrst2!SBalance))
                        lblTopUp = IIf(IsNumeric(Trim(AnnSpecLimitrst2!TopUpAmt)), Format(Trim(AnnSpecLimitrst2!TopUpAmt), "#,##0.00"), Trim(AnnSpecLimitrst2!TopUpAmt))
                        lblNoAnnLimtDed = IIf(IsNumeric(Trim(AnnSpecLimitrst2!GNoLimitDedTotal)), Format(Trim(AnnSpecLimitrst2!GNoLimitDedTotal), "#,##0.00"), Trim(AnnSpecLimitrst2!GNoLimitDedTotal))
                        lblNoSpecLimtDed = IIf(IsNumeric(Trim(AnnSpecLimitrst2!SNoLimitDedTotal)), Format(Trim(AnnSpecLimitrst2!SNoLimitDedTotal), "#,##0.00"), Trim(AnnSpecLimitrst2!SNoLimitDedTotal))
                        lblPLNCode = IIf(IsNumeric(Trim(AnnSpecLimitrst2!PLNCode)), Format(Trim(AnnSpecLimitrst2!PLNCode), "###"), Trim(AnnSpecLimitrst2!PLNCode))
                        If Len(Trim(AnnSpecLimitrst2!LGIssuedInd)) Then
                            lblLGIssuedInd = Trim(AnnSpecLimitrst2!LGIssuedInd)
                        Else
                            lblLGIssuedInd = "N"
                        End If
                        If Len(Trim(AnnSpecLimitrst2!LGIssuedAfterLimitExInd)) Then
                            lblLGIssuedAfterLimitExInd = Trim(AnnSpecLimitrst2!LGIssuedAfterLimitExInd)
                        Else
                            lblLGIssuedAfterLimitExInd = "N"
                        End If
                        If Len(Trim(AnnSpecLimitrst2!LGUsingLimit)) Then
                            lblUsingLimit = Trim(AnnSpecLimitrst2!LGUsingLimit)
                        Else
                            lblUsingLimit = "N/A"
                        End If
                        If Len(Trim(AnnSpecLimitrst2!LGUsingLimitDesc)) Then
                            lblUsingLimitDesc = Trim(AnnSpecLimitrst2!LGUsingLimitDesc)
                        Else
                            lblUsingLimitDesc = "N/A"
                        End If
                        If Trim(AnnSpecLimitrst2!error) <> "N/A" Then
                            MsgBox AnnSpecLimitrst2!error, vbOKOnly + vbCritical
                        End If
                        
                    End With
                    AnnSpecLimitrst2.MoveNext
                Loop
                AnnSpecLimitrst2.Close
                Set AnnSpecLimitrst2 = Nothing
            End If
            
            TempLG = lstLG.SelectedItem.Text
            lblLG = lstLG.SelectedItem.Text
            txtDoctor = lstLG.SelectedItem.SubItems(4)
            txtDate = IIf(Len(lstLG.SelectedItem.SubItems(2)), lstLG.SelectedItem.SubItems(2), "")
            cboCLNCode = lstLG.SelectedItem.SubItems(5)
            txtRemarks = lstLG.SelectedItem.SubItems(12)
            txtDiagnosis = lstLG.SelectedItem.SubItems(7)
            cmdDelayreason = lstLG.SelectedItem.SubItems(16)
            tmpLGCasDate = lstLG.SelectedItem.SubItems(2)
            tmpLGCasTime = lstLG.SelectedItem.SubItems(17)
                        
            If AddFlag = True Then
                If Mid(cboIndicator, 1, 1) = "F" Then
                    cboIndicator.ListIndex = 1
                Else
                    cboIndicator.ListIndex = 0
                End If
            Else
                If Mid(lstLG.SelectedItem.Text, Len(lstLG.SelectedItem.Text), 1) = 1 Then
                    cboIndicator.ListIndex = 0
                Else
                    cboIndicator.ListIndex = 1
                End If
            End If
        Else
            lstLG.Enabled = False
        End If
        
        If AddFlag = False And Mid(cboIndicator, 1, 1) <> "F" Then
            cboCLNCode.Enabled = False
            chkCLN.Enabled = False
            txtDoctor.Enabled = False
            txtDate.Enabled = False
            txtRemarks.Enabled = False
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            txtDiagnosis.Enabled = False
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            cboIndicator.Enabled = False
        End If
    
    If IsPharmacy Then
        If IsMedRequestInProgress(lstLG.SelectedItem.Text) Then
            cmdDecline.Enabled = False
            cmdLG.Enabled = False
            cmdEdit.Enabled = False
            cmdCopyFile.Enabled = False
        Else
            cmdDecline.Enabled = True
            cmdLG.Enabled = True
            cmdEdit.Enabled = True
            If Not IsMasterCase Then
                cmdCopyFile.Enabled = True
            End If
        End If
    Else
        cmdEdit.Enabled = True
    End If
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub



Private Sub lstLG_KeyUp(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "lstLG_KeyUp"
On Error GoTo debugs
Dim rst As New ADODB.Recordset
Dim StrSql As String
''''''''''''''''''''''''''''''''''''''''''
Dim AnnSpecLimitcmd As New ADODB.Command
Dim AnnSpecLimitPrm1 As New ADODB.Parameter
Dim AnnSpecLimitPrm2 As New ADODB.Parameter
Dim AnnSpecLimitrst2 As New ADODB.Recordset
'''''''''''''''''''''''''''''''''''''''''''
    
TempLG = ""
tmpLGCasDate = ""
tmpLGCasTime = ""

  '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        
        If Len(lstLG.SelectedItem.SubItems(15)) And lstLG.ListItems.Count <> 0 Then
            MsgBox lstLG.SelectedItem.SubItems(15), vbOKOnly + vbInformation, DialogBoxTitle
        End If

        If lstLG.ListItems.Count <> 0 Then
            'cmdEdit.Enabled = True
            If TempMBM = False Then
                StrSql = " Select * From View_PrintLG Where LGNo = '" & lstLG.SelectedItem.Text & "'"
                rst.Open StrSql, MediConnCISDB, adOpenDynamic, adLockOptimistic
                Do Until rst.EOF
                    If rst!coverid = rst!MBMCOVID Then
                        txtMBMNo = rst!MBMNumber & "-" & rst!MBMCOVID
                        txtICNo = rst!MBMIcBcPp
                        txtPatient = rst!MBMName
                        Exit Do
                    ElseIf rst!coverid = rst!MBMCCoverID Then
                        txtMBMNo = rst!MBMNumber & "-" & rst!MBMCCoverID
                        txtICNo = rst!MBMCICBCPPNo
                        txtPatient = rst!MBMCName
                        Exit Do
                    End If
                    rst.MoveNext
                Loop
                rst.Close
                Set rst = Nothing
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Set AnnSpecLimitcmd.ActiveConnection = MediConnCISWS
                AnnSpecLimitcmd.CommandText = "GetLGAnnLimitSpecLimit"
                AnnSpecLimitcmd.CommandType = adCmdStoredProc
                AnnSpecLimitcmd.CommandTimeout = 9000
                Set AnnSpecLimitPrm1 = AnnSpecLimitcmd.CreateParameter("SchMBMNumber", adVarChar, adParamInput, 2000, Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1)))
                AnnSpecLimitcmd.Parameters.Append AnnSpecLimitPrm1
                Set AnnSpecLimitPrm2 = AnnSpecLimitcmd.CreateParameter("CoverId", adVarChar, adParamInput, 2000, Right(txtMBMNo, Len(txtMBMNo) - InStr(txtMBMNo, "-")))
                AnnSpecLimitcmd.Parameters.Append AnnSpecLimitPrm2
                AnnSpecLimitrst2.CursorLocation = adUseClient
                AnnSpecLimitrst2.CursorType = adOpenForwardOnly
                AnnSpecLimitrst2.CacheSize = 100
                Set AnnSpecLimitrst2 = AnnSpecLimitcmd.Execute
                
                Do While Not AnnSpecLimitrst2.EOF
                    With AnnSpecLimitrst2
                        lblAnnLmt = IIf(IsNumeric(Trim(AnnSpecLimitrst2!OriAnnLimit)), Format(Trim(AnnSpecLimitrst2!OriAnnLimit), "#,##0.00"), Trim(AnnSpecLimitrst2!OriAnnLimit))
                        If ASOInd = "Y" Then
                            lblAccumAmt = UtilisedAmt
                        Else
                            lblAccumAmt = IIf(IsNumeric(Trim(AnnSpecLimitrst2!GUtilized)), Format(Trim(AnnSpecLimitrst2!GUtilized), "#,##0.00"), Trim(AnnSpecLimitrst2!GUtilized))
                        End If
                        lblBalLmt = IIf(IsNumeric(Trim(AnnSpecLimitrst2!GBalance)), Format(Trim(AnnSpecLimitrst2!GBalance), "#,##0.00"), Trim(AnnSpecLimitrst2!GBalance))
                        lblSpecAnn = IIf(IsNumeric(Trim(AnnSpecLimitrst2!OriSpecAnnLimit)), Format(Trim(AnnSpecLimitrst2!OriSpecAnnLimit), "#,##0.00"), Trim(AnnSpecLimitrst2!OriSpecAnnLimit))
                        lblSpecAcc = IIf(IsNumeric(Trim(AnnSpecLimitrst2!SUtilized)), Format(Trim(AnnSpecLimitrst2!SUtilized), "#,##0.00"), Trim(AnnSpecLimitrst2!SUtilized))
                        lblSpecBal = IIf(IsNumeric(Trim(AnnSpecLimitrst2!SBalance)), Format(Trim(AnnSpecLimitrst2!SBalance), "#,##0.00"), Trim(AnnSpecLimitrst2!SBalance))
                        lblTopUp = IIf(IsNumeric(Trim(AnnSpecLimitrst2!TopUpAmt)), Format(Trim(AnnSpecLimitrst2!TopUpAmt), "#,##0.00"), Trim(AnnSpecLimitrst2!TopUpAmt))
                        lblNoAnnLimtDed = IIf(IsNumeric(Trim(AnnSpecLimitrst2!GNoLimitDedTotal)), Format(Trim(AnnSpecLimitrst2!GNoLimitDedTotal), "#,##0.00"), Trim(AnnSpecLimitrst2!GNoLimitDedTotal))
                        lblNoSpecLimtDed = IIf(IsNumeric(Trim(AnnSpecLimitrst2!SNoLimitDedTotal)), Format(Trim(AnnSpecLimitrst2!SNoLimitDedTotal), "#,##0.00"), Trim(AnnSpecLimitrst2!SNoLimitDedTotal))
                        lblPLNCode = IIf(IsNumeric(Trim(AnnSpecLimitrst2!PLNCode)), Format(Trim(AnnSpecLimitrst2!PLNCode), "###"), Trim(AnnSpecLimitrst2!PLNCode))
                        
                        If Len(Trim(AnnSpecLimitrst2!LGIssuedInd)) Then
                            lblLGIssuedInd = Trim(AnnSpecLimitrst2!LGIssuedInd)
                        Else
                            lblLGIssuedInd = "N"
                        End If
                        
                        If Len(Trim(AnnSpecLimitrst2!LGIssuedAfterLimitExInd)) Then
                            lblLGIssuedAfterLimitExInd = Trim(AnnSpecLimitrst2!LGIssuedAfterLimitExInd)
                        Else
                            lblLGIssuedAfterLimitExInd = "N"
                        End If
                        
                        If Len(Trim(AnnSpecLimitrst2!LGUsingLimit)) Then
                            lblUsingLimit = Trim(AnnSpecLimitrst2!LGUsingLimit)
                        Else
                            lblUsingLimit = "N/A"
                        End If
                        
                        If Len(Trim(AnnSpecLimitrst2!LGUsingLimitDesc)) Then
                            lblUsingLimitDesc = Trim(AnnSpecLimitrst2!LGUsingLimitDesc)
                        Else
                            lblUsingLimitDesc = "N/A"
                        End If
                        
                        If Trim(AnnSpecLimitrst2!error) <> "N/A" Then
                            MsgBox AnnSpecLimitrst2!error, vbOKOnly + vbCritical
                        End If
                        
                    End With
                    AnnSpecLimitrst2.MoveNext
                Loop
                
                AnnSpecLimitrst2.Close
                Set AnnSpecLimitrst2 = Nothing
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            End If
            
            TempLG = lstLG.SelectedItem.Text
            lblLG = lstLG.SelectedItem.Text
            txtDoctor = lstLG.SelectedItem.SubItems(4)
            txtDate = lstLG.SelectedItem.SubItems(2)
            cboCLNCode = lstLG.SelectedItem.SubItems(5)
            txtRemarks = lstLG.SelectedItem.SubItems(12)
            ''''''''''''''''''''''''''''''''''''''''''''''
            txtDiagnosis = lstLG.SelectedItem.SubItems(7)
            cmdDelayreason = lstLG.SelectedItem.SubItems(16)
            tmpLGCasDate = lstLG.SelectedItem.SubItems(2)
            tmpLGCasTime = lstLG.SelectedItem.SubItems(17)
            ''''''''''''''''''''''''''''''''''''''''''''''
            
            If AddFlag = True Then
                If Mid(cboIndicator, 1, 1) = "F" Then
                    cboIndicator.ListIndex = 1
                Else
                    cboIndicator.ListIndex = 0
                End If
            Else
                If Mid(lstLG.SelectedItem.Text, Len(lstLG.SelectedItem.Text), 1) = 1 Then
                    cboIndicator.ListIndex = 0
                Else
                    cboIndicator.ListIndex = 1
                End If
            End If
        Else
            lstLG.Enabled = False
        End If
        
        If AddFlag = False And Mid(cboIndicator, 1, 1) <> "F" Then
            cboCLNCode.Enabled = False
            chkCLN.Enabled = False
            txtDoctor.Enabled = False
            txtDate.Enabled = False
            txtRemarks.Enabled = False
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            txtDiagnosis.Enabled = False
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            cboIndicator.Enabled = False
        End If
    
    If IsPharmacy Then
        If IsMedRequestInProgress(lstLG.SelectedItem.Text) Then
            cmdDecline.Enabled = False
            cmdLG.Enabled = False
            cmdEdit.Enabled = False
            cmdCopyFile.Enabled = False
        Else
            cmdDecline.Enabled = True
            cmdLG.Enabled = True
            cmdEdit.Enabled = True
            If Not IsMasterCase Then
                cmdCopyFile.Enabled = True
            End If
        End If
    Else
        cmdEdit.Enabled = True
    End If
    
   ' MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' MediConnCISWS.Close
   ' Set MediConnCISWS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub lstMBM_Click()
    Call CheckMBMLimit
    Dim IcNo As String
    IcNo = txtICNo.Text
    Dim MBMNo As String
    MBMNo = txtMBMNo.Text
    Call ValidateLG(IcNo, MBMNo)
    If lblMBMStatus <> "A" Then
           MsgBox "Member is not active.", vbOKOnly + vbCritical, DialogBoxTitle
           cmdAdd.Enabled = False
           cmdLG.Enabled = False
           Exit Sub
    Else
           cmdAdd.Enabled = True
           cmdLG.Enabled = True
           Exit Sub
    End If
End Sub

Private Sub lstMBM_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CheckMBMLimit
End Sub

Private Sub lstMBM_KeyUp(KeyCode As Integer, Shift As Integer)
    Call CheckMBMLimit
End Sub
Private Sub CheckMBMLimit()
Const METHOD_NAME = "CheckMBMLimit"
On Error GoTo debugs
Dim AnnSpecLimitcmd As New ADODB.Command
Dim AnnSpecLimitPrm1 As New ADODB.Parameter
Dim AnnSpecLimitPrm2 As New ADODB.Parameter
Dim AnnSpecLimitrst2 As New ADODB.Recordset
Dim TopupRec As New ADODB.Recordset
Dim StrSql As String
Dim PLN As New ADODB.Recordset
Dim tmpMaternityInd As String
Dim tmpclearMBMNO As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim strAppend As String
Dim tmpLimitStatus As String
Dim MBMNo As String
Dim CovID As String
Dim sArr() As String
Dim dblMeternity As Double
Dim dblDental As Double
Dim OtherDed As Double
Dim thyroid As Double
Dim SuppAnnLmt As String
Dim BordxAmt As Double
Dim isPrinDepGPShared As Boolean

SuppAnnLmt = 0
    If lstMBM.ListItems.Count = 0 Then
        TempMBM = False
    Else
        TempMBM = True
    End If
    lblCount = 0
    If lstMBM.ListItems.Count <> 0 Then
        txtMBMNo = lstMBM.SelectedItem.Text
        RecordArray() = Split(txtMBMNo, "-")
        tmpCoverID = RecordArray(1)
        'tmpMBMName = lstMBM.SelectedItem.SubItems(5)
        tmpclearMBMNO = txtMBMNo
        txtMBMNo = ""
        txtPatient = ""
        txtICNo = ""
        txtRemarks = ""
        txtDiagnosis = ""
        txtDoctor = ""
        txtDate = "__/__/____"
        txtMBMNo.Enabled = False
        txtPatient.Enabled = False
        txtICNo.Enabled = False
        cboLetterHead.Enabled = False
        txtRemarks.Enabled = False
        txtDiagnosis.Enabled = False
        cboCLNCode.Enabled = False
        txtDoctor.Enabled = False
        txtDate.Enabled = False
        cboIndicator.Enabled = False
        txtGrpCompany.Enabled = False
        chkCLN.Enabled = False
        txtDoctor.Clear
        cboLetterHead.Clear
        cboCLNCode.Clear
        cboIndicator.Clear
        chkCLN.Value = 0
        Call ComboListing
        lblAnnLmt = ""
        lblAccumAmt = ""
        lblBalLmt = ""
        lblSpecAnn = ""
        lblSpecAcc = ""
        lblSpecBal = ""
        lblTopUp = ""
        lblsptopup = ""
        lblNoAnnLimtDed = ""
        lblNoSpecLimtDed = ""
        lblPLNCode = ""
        lblLGIssuedInd = ""
        lblUsingLimitDesc = ""
        lblUsingLimit = ""
        lblLGIssuedAfterLimitExInd = ""
        lblCount = ""
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        lstLG.ListItems.Clear
        txtMBMNo = tmpclearMBMNO
        txtPatient = lstMBM.SelectedItem.SubItems(1)
        txtICNo = lstMBM.SelectedItem.SubItems(2)
        Dim SCovId As String
        SCovId = Right(txtMBMNo, Len(txtMBMNo) - InStr(txtMBMNo, "-"))
        Set AnnSpecLimitcmd.ActiveConnection = MediConnCISWS
        AnnSpecLimitcmd.CommandText = "GetLGAnnLimitSpecLimit"
        AnnSpecLimitcmd.CommandType = adCmdStoredProc
        AnnSpecLimitcmd.CommandTimeout = 9000
        Set AnnSpecLimitPrm1 = AnnSpecLimitcmd.CreateParameter("SchMBMNumber", adVarChar, adParamInput, 2000, Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1)))
        AnnSpecLimitcmd.Parameters.Append AnnSpecLimitPrm1
        Set AnnSpecLimitPrm2 = AnnSpecLimitcmd.CreateParameter("CoverId", adVarChar, adParamInput, 2000, Right(txtMBMNo, Len(txtMBMNo) - InStr(txtMBMNo, "-")))
        AnnSpecLimitcmd.Parameters.Append AnnSpecLimitPrm2
        AnnSpecLimitrst2.CursorLocation = adUseClient
        AnnSpecLimitrst2.CursorType = adOpenForwardOnly
        AnnSpecLimitrst2.CacheSize = 100
        Set AnnSpecLimitrst2 = AnnSpecLimitcmd.Execute
        Do While Not AnnSpecLimitrst2.EOF
            With AnnSpecLimitrst2
                TmpEffDate = AnnSpecLimitrst2!PolEffDate
                SuppAnnLmt = IIf(IsNumeric(Trim(AnnSpecLimitrst2!SuppLimit)), Format(Trim(AnnSpecLimitrst2!SuppLimit), "#,##0.00"), Trim(AnnSpecLimitrst2!SuppLimit))
                lblAnnLmt = IIf(IsNumeric(Trim(AnnSpecLimitrst2!OriAnnLimit)), Format(Trim(AnnSpecLimitrst2!OriAnnLimit), "#,##0.00"), Trim(AnnSpecLimitrst2!OriAnnLimit))
                If ASOInd = "Y" Then
                    lblAccumAmt = UtilisedAmt
                Else
                    lblAccumAmt = IIf(IsNumeric(Trim(AnnSpecLimitrst2!GUtilized)), Format(Trim(AnnSpecLimitrst2!GUtilized), "#,##0.00"), Trim(AnnSpecLimitrst2!GUtilized))
                End If
                
                lblCount = IIf(IsNumeric(Trim(AnnSpecLimitrst2!TopUpAmt)), Format(Trim(AnnSpecLimitrst2!TopUpAmt), "#,##0.00"), 0)
                lblSpecAnn = IIf(IsNumeric(Trim(AnnSpecLimitrst2!OriSpecAnnLimit)), Format(Trim(AnnSpecLimitrst2!OriSpecAnnLimit), "#,##0.00"), Trim(AnnSpecLimitrst2!OriSpecAnnLimit))
                lblSpecAcc = IIf(IsNumeric(Trim(AnnSpecLimitrst2!SUtilized)), Format(Trim(AnnSpecLimitrst2!SUtilized), "#,##0.00"), Trim(AnnSpecLimitrst2!SUtilized))
                lblSpecBal = IIf(IsNumeric(Trim(AnnSpecLimitrst2!SBalance)), Format(Trim(AnnSpecLimitrst2!SBalance), "#,##0.00"), Trim(AnnSpecLimitrst2!SBalance))
                lblTopUp = IIf(IsNumeric(Trim(AnnSpecLimitrst2!TopUpAmt)), Format(Trim(AnnSpecLimitrst2!TopUpAmt), "#,##0.00"), Trim(AnnSpecLimitrst2!TopUpAmt))
                'lblNoAnnLimtDed = IIf(IsNumeric(Trim(AnnSpecLimitrst2!GNoLimitDedTotal)), Format(Trim(AnnSpecLimitrst2!GNoLimitDedTotal), "#,##0.00"), Trim(AnnSpecLimitrst2!GNoLimitDedTotal))
                lblNoSpecLimtDed = IIf(IsNumeric(Trim(AnnSpecLimitrst2!SNoLimitDedTotal)), Format(Trim(AnnSpecLimitrst2!SNoLimitDedTotal), "#,##0.00"), Trim(AnnSpecLimitrst2!SNoLimitDedTotal))
                'lblPLNCode = IIf(IsNumeric(Trim(AnnSpecLimitrst2!PLNCode)), Format(Trim(AnnSpecLimitrst2!PLNCode), "###"), Trim(AnnSpecLimitrst2!PLNCode))
                lblPLNCode = AnnSpecLimitrst2!PLNCode
                tmpPLNGLNO = IIf(Len(AnnSpecLimitrst2!PLNGLNO), AnnSpecLimitrst2!PLNGLNO, "")
                
                tmpSpecSuppLimitStatus = IIf(Len(AnnSpecLimitrst2!SpecSuppLimitStatus), Trim(AnnSpecLimitrst2!SpecSuppLimitStatus), "")
                If tmpCoverID <> "00" Then
                    If Trim(AnnSpecLimitrst2!LGIssuedSuppInd) = "N" Then
                        lblLGIssuedInd = "N"
                    Else
                        lblLGIssuedInd = "Y"
                    End If
                Else
                    If Len(Trim(AnnSpecLimitrst2!LGIssuedInd)) Then
                        lblLGIssuedInd = Trim(AnnSpecLimitrst2!LGIssuedInd)
                    Else
                        lblLGIssuedInd = "N"
                    End If
                End If
                
                If Len(Trim(AnnSpecLimitrst2!LGIssuedAfterLimitExInd)) Then
                    lblLGIssuedAfterLimitExInd = Trim(AnnSpecLimitrst2!LGIssuedAfterLimitExInd)
                Else
                    lblLGIssuedAfterLimitExInd = "N"
                End If
                
                If Len(Trim(AnnSpecLimitrst2!LGUsingLimit)) Then
                    lblUsingLimit = Trim(AnnSpecLimitrst2!LGUsingLimit)
                Else
                    lblUsingLimit = "N/A"
                End If
                
                If Len(Trim(AnnSpecLimitrst2!LGUsingLimitDesc)) Then
                    lblUsingLimitDesc = Trim(AnnSpecLimitrst2!LGUsingLimitDesc)
                Else
                    lblUsingLimitDesc = "N/A"
                End If
                
               ' If Trim(AnnSpecLimitrst2!Error) <> "N/A" Then
               '     MsgBox AnnSpecLimitrst2!Error, vbOKOnly + vbCritical
               ' End If
                
                If AnnSpecLimitrst2!HRApproval = "Y" Then
                    If MsgBox(" All MO/SP GL Request Must be Come From HR.", vbExclamation + vbOKCancel, DialogBoxTitle) = vbCancel Then
                        Exit Sub
                    End If
                End If
            End With
            AnnSpecLimitrst2.MoveNext
        Loop
        
        AnnSpecLimitrst2.Close
        Set AnnSpecLimitrst2 = Nothing
        lblsptopup = 0
        TopupRec.Open "SELECT Top 1 SUM(coalesce(TopupAmt,0)) as sumOfTopupAmt,Sum(coalesce(SP_TopupAmt,0)) AS SP_TopupAmt from CISDB..MBMCISTopupHistory where MBMNumber='" & Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1)) & "' and MBMCovID='" & Right(txtMBMNo, Len(txtMBMNo) - InStr(txtMBMNo, "-")) & "' ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            With TopupRec
                If TopupRec.EOF = False Then
                    lblsptopup = Format(CDbl(lblsptopup) + IIf(IsNumeric(Trim(TopupRec!SP_TopupAmt)), Format(Trim(TopupRec!SP_TopupAmt), "#,##0.00"), 0), "#,##0.00")
                Else
                    lblsptopup = "0"
                End If
            End With
        TopupRec.Close
        Set TopupRec = Nothing
        '------Calculate Limt Start-------
        Dim SpecLmtRec As New ADODB.Recordset
        Dim tmpAnnLmt As String
        Dim tmpVisitLmt As String
        Dim tmpSpecLmt As String
        Dim tmpSpecAnnLmt As String
        Dim tmpSpecLimitStatus As String
        Dim tmpSpecAccLmt As String
        Dim tmpDentalInd As String
            
        strAppend = ""
        tmpMaternityInd = ""
            
        AnnSpecLimitrst2.Open "SELECT Top 1 SuppLimitStatus FROM AnnualLimit WHERE PLNCode='" & Trim(lblPLNCode) & "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        If AnnSpecLimitrst2.EOF = False Then
            'lblLimitStatus = IIf(Len(AnnSpecLimitrst2!SuppLimitStatus), AnnSpecLimitrst2!SuppLimitStatus, "A")
            tmpLimitStatus = IIf(Len(AnnSpecLimitrst2!SuppLimitStatus), AnnSpecLimitrst2!SuppLimitStatus, "A")
        End If
        AnnSpecLimitrst2.Close
        Set AnnSpecLimitrst2 = Nothing
            
        If SCovId = "00" Then
            Call CallLimit("P", tmpAnnLmt, tmpVisitLmt, tmpSpecLmt, tmpLimitStatus)
        ElseIf SCovId <> "00" Then
            Call CallLimit("S", tmpAnnLmt, tmpVisitLmt, tmpSpecLmt, tmpLimitStatus)
        End If
        
        lblNoAnnLimtDed = 0
        If Len(txtMBMNo) Then
            If InStr(1, txtMBMNo, "-") Then
               sArr = Split(txtMBMNo, "-")
               MBMNo = sArr(0)
               CovID = sArr(1)
            Else
               MBMNo = txtMBMNo
               CovID = "00"
            End If
        End If
                    
    If tmpLimitStatus = "E" And CovID = "00" Then
        tmpAnnLmt = "FULL"
    End If
    
    If tmpLimitStatus = "E" And CovID <> "00" Then
        lblAnnLmt = CStr(SuppAnnLmt)
    Else
        lblAnnLmt = lblAnnLmt
    End If
    
    SpecLmtRec.Open "SELECT Top 1 PLNMaternityInd,PDentalStatus, COALESCE(IsPrinDepGPShared,0) AS IsPrinDepGPShared FROM PLN WHERE PLNCode='" & Trim(lblPLNCode) & _
    "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    If SpecLmtRec.EOF = False Then
        tmpMaternityInd = IIf(Len(SpecLmtRec!PLNMaternityInd), SpecLmtRec!PLNMaternityInd, "")
        tmpDentalInd = IIf(Len(SpecLmtRec!PDentalStatus), SpecLmtRec!PDentalStatus, "")
        isPrinDepGPShared = SpecLmtRec!isPrinDepGPShared
    End If
    SpecLmtRec.Close
    Set SpecLmtRec = Nothing
    
    If Trim(lblSpecAnn) = "N/A" Then
        If Len(txtMBMNo) Then
            strAppend = ""
            If isPrinDepGPShared Then
                strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' "
            Else
                If tmpLimitStatus = "A" Then
                    strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "'"
                ElseIf tmpLimitStatus = "B" Then
                    If CovID = "00" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID = '00' "
                    Else
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID <> '00' "
                    End If
                ElseIf tmpLimitStatus = "C" Or tmpLimitStatus = "" Then
                    strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                ElseIf tmpLimitStatus = "D" Then
                    If CovID = "00" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID = '00' "
                    Else
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID <> '00' "
                    End If
                ElseIf tmpLimitStatus = "E" Then
                    strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                End If
            End If
        End If

        lblAccumAmt = 0
        lblAccumAmt = 0

        If Trim(strAppend) = "" Then
            'tmpMBMFFS = 0
            'tmpMBMCLMAmt = 0
        Else
            Set cmd.ActiveConnection = MediConnCISDB
            cmd.CommandText = "CIS_FFS2013"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 9000
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
                
                BordxAmt = 0
                If Len(rst2!CLMBordxAmt) Then
                   If IsNumeric(rst2!CLMBordxAmt) Then
                       BordxAmt = CDbl(rst2!CLMBordxAmt)
                   Else
                       BordxAmt = 0
                   End If
                Else
                       BordxAmt = 0
                End If
                If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                    lblAccumAmt = CDbl(lblAccumAmt) + BordxAmt - (dblMeternity + OtherDed + thyroid + dblDental)
                    lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (dblMeternity + OtherDed + thyroid + dblDental)
                ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                    lblAccumAmt = CDbl(lblAccumAmt) + BordxAmt - (dblMeternity + OtherDed + thyroid)
                    lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (dblMeternity + OtherDed + thyroid)
                ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                    lblAccumAmt = CDbl(lblAccumAmt) + BordxAmt - (OtherDed + thyroid + dblDental)
                    lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (OtherDed + thyroid + dblDental)
                ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                      lblAccumAmt = CDbl(lblAccumAmt) + BordxAmt - (OtherDed + thyroid)
                    lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (OtherDed + thyroid)
                End If
                    
                If ASOInd = "Y" Then
                    lblAccumAmt = UtilisedAmt
                End If
                rst2.MoveNext
            Loop
            rst2.Close
            Set rst2 = Nothing
            Set cmd = Nothing
            Set Prm1 = Nothing
        End If
        lblAccumAmt = Format(lblAccumAmt, "#,###.00")
        lblSpecAcc = "N/A"
        lblSpecBal = "N/A"
    Else
        lblAccumAmt = 0
        lblNoAnnLimtDed = 0
        tmpSpecAccLmt = 0
           
        If Len(txtMBMNo) Then
            If Len(txtMBMNo) Then
                If InStr(1, txtMBMNo, "-") Then
                    sArr = Split(txtMBMNo, "-")
                    MBMNo = sArr(0)
                    CovID = sArr(1)
                Else
                    MBMNo = txtMBMNo
                    CovID = "00"
                End If
            End If

            strAppend = ""
            If isPrinDepGPShared Then
                strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' "
            Else
                If tmpLimitStatus = "A" Then
                    If tmpSpecSuppLimitStatus = "C" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                    Else
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "'"
                    End If
                    'strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(MBMNo) & "'"
                ElseIf tmpLimitStatus = "B" Then
                    If CovID = "00" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID = '00' "
                    Else
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID <> '00' "
                    End If
                ElseIf tmpLimitStatus = "C" Or tmpLimitStatus = "" Then
                    strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                ElseIf tmpLimitStatus = "D" Then
                    If CovID = "00" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID = '00' "
                    Else
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID <> '00' "
                    End If
                ElseIf tmpLimitStatus = "E" Then
                    'strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                    strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                End If
            End If
        End If
            
        If Trim(strAppend) = "" Then
           ' tmpMBMFFS = 0
           ' tmpMBMCLMAmt = 0
        Else
            Set cmd.ActiveConnection = MediConnCISDB
            cmd.CommandText = "CIS_FFS2013"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 9000
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append Prm1
            rst2.CursorLocation = adUseClient
            rst2.CursorType = adOpenForwardOnly
            rst2.CacheSize = 100
            Set rst2 = cmd.Execute
            
            Do While Not rst2.EOF
                dblMeternity = 0
                dblDental = 0
                OtherDed = 0
                thyroid = 0
                
                If IsNumeric(rst2!CLMMaternity) Then
                    dblMeternity = CDbl(rst2!CLMMaternity)
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
                
                BordxAmt = 0
                If Len(rst2!CLMBordxAmt) Then
                   If IsNumeric(rst2!CLMBordxAmt) Then
                       BordxAmt = CDbl(rst2!CLMBordxAmt)
                   Else
                       BordxAmt = 0
                   End If
                Else
                       BordxAmt = 0
                End If
                    
                If rst2!PLNGPSPShared = "Y" Then
                    If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                        lblAccumAmt = CDbl(lblAccumAmt) + BordxAmt - (dblMeternity + OtherDed + thyroid + dblDental)
                        lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (dblMeternity + OtherDed + thyroid + dblDental)
                    ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                        lblAccumAmt = CDbl(lblAccumAmt) + BordxAmt - (dblMeternity + OtherDed + thyroid)
                        lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (dblMeternity + OtherDed + thyroid)
                    ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                        lblAccumAmt = CDbl(lblAccumAmt) + BordxAmt - (OtherDed + thyroid + dblDental)
                        lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (OtherDed + thyroid + dblDental)
                    ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                        lblAccumAmt = CDbl(lblAccumAmt) + BordxAmt - (OtherDed + thyroid)
                        lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (OtherDed + thyroid)
                    End If
                        
                    If ASOInd = "Y" Then
                        lblAccumAmt = UtilisedAmt
                    End If
                Else
                    If rst2!CLNType <> "S" Then
                        If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                            lblAccumAmt = CDbl(lblAccumAmt) + BordxAmt - (dblMeternity + OtherDed + thyroid + dblDental)
                            lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (dblMeternity + OtherDed + thyroid + dblDental)
                        ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                            lblAccumAmt = CDbl(lblAccumAmt) + BordxAmt - (dblMeternity + OtherDed + thyroid)
                            lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (dblMeternity + OtherDed + thyroid)
                        ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                            lblAccumAmt = CDbl(lblAccumAmt) + BordxAmt - (OtherDed + thyroid + dblDental)
                            lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (OtherDed + thyroid + dblDental)
                        ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                            lblAccumAmt = CDbl(lblAccumAmt) + BordxAmt - (OtherDed + thyroid)
                            lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (OtherDed + thyroid)
                        End If
                        
                        If ASOInd = "Y" Then
                          lblAccumAmt = UtilisedAmt
                        End If
                    ElseIf rst2!CLNType = "S" Then
                        Dim bAccLmt As Boolean
                        bAccLmt = False
                        If (tmpLimitStatus = "A" And tmpSpecSuppLimitStatus = "C" And rst2!coverid = CovID) Or (tmpLimitStatus = "A" And tmpSpecSuppLimitStatus <> "C") Then
                            bAccLmt = True
                        ElseIf (tmpLimitStatus = "B" And CovID = "00" And rst2!coverid = "00") Or (tmpLimitStatus = "B" And CovID <> "00" And rst2!coverid <> "00") Then
                           bAccLmt = True
                        ElseIf (tmpLimitStatus = "C" Or tmpLimitStatus = "") And rst2!coverid = CovID Then
                            bAccLmt = True
                        ElseIf (tmpLimitStatus = "D" And CovID = "00" And rst2!coverid = "00") Or (tmpLimitStatus = "D" And CovID <> "00" And rst2!coverid <> "00") Then
                            bAccLmt = True
                        ElseIf tmpLimitStatus = "E" And rst2!coverid = CovID Then
                            bAccLmt = True
                        End If
                        If bAccLmt = True Then
                            If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                                tmpSpecAccLmt = CDbl(tmpSpecAccLmt) + BordxAmt - (dblMeternity + OtherDed + thyroid + dblDental)
                                lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (dblMeternity + OtherDed + thyroid + dblDental)
                            ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                                tmpSpecAccLmt = CDbl(tmpSpecAccLmt) + BordxAmt - (dblMeternity + OtherDed + thyroid)
                                lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (dblMeternity + OtherDed + thyroid)
                            ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                                tmpSpecAccLmt = CDbl(tmpSpecAccLmt) + BordxAmt - (OtherDed + thyroid + dblDental)
                                lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (OtherDed + thyroid + dblDental)
                            ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                                tmpSpecAccLmt = CDbl(tmpSpecAccLmt) + BordxAmt - (OtherDed + thyroid)
                                lblNoAnnLimtDed = CDbl(lblNoAnnLimtDed) + (OtherDed + thyroid)
                            End If
                        End If
                        If ASOInd = "Y" Then
                          lblAccumAmt = UtilisedAmt
                          lblAnnLmt = tmpAnnLmt
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
    End If
            
    '------Calculate Limt Ends-------
    '------Assigining  Calculated Values-------
    If ASOInd = "Y" Then
      lblAccumAmt = UtilisedAmt
      lblAnnLmt = tmpAnnLmt
    End If
    lblAccumAmt = Format(lblAccumAmt, "#,##0.00")
    If IsNumeric(lblAnnLmt) Then
        If IsNumeric(lblTopUp) Then
            lblBalLmt = Format(((CDbl(lblAnnLmt) + CDbl(lblTopUp)) - CDbl(lblAccumAmt)), "#,##0.00")
        Else
            lblBalLmt = Format(((CDbl(lblAnnLmt)) - CDbl(lblAccumAmt)), "#,##0.00")
        End If
    Else
        lblAnnLmt = lblAnnLmt
        lblBalLmt = lblAnnLmt
    End If
          
    tmpSpecAccLmt = Format(tmpSpecAccLmt, "#,##0.00")
    If IsNumeric(tmpSpecLmt) Then
        lblSpecAnn = Format(tmpSpecLmt, "#,##0.00")
    Else
        lblSpecAnn = tmpSpecLmt
    End If
    If IsNumeric(lblSpecAnn) Then
        lblSpecAcc = Format(tmpSpecAccLmt, "#,##0.00")
        
        If IsNumeric(tmpSpecLmt) Then
            If IsNumeric(tmpSpecAccLmt) Then
                lblSpecBal = Format(CDbl(tmpSpecLmt) + CDbl(lblsptopup) - CDbl(tmpSpecAccLmt), "#,##0.00")
            End If
        Else
            lblSpecBal = "FULL"
        End If
    Else
        If Trim(lblSpecAnn) = "FULL" Then
            lblSpecAcc = Format(tmpSpecAccLmt, "#,##0.00")
            lblSpecBal = "FULL"
        ElseIf Trim(lblSpecAnn) = "N/A" Then
            lblSpecAcc = "N/A"
            lblSpecBal = "N/A"
        Else
            lblSpecAnn = "N/A"
            lblSpecAcc = "N/A"
            lblSpecBal = "N/A"
        End If
    End If
            
    If tmpLimitStatus = "C" And CovID = "00" Then
        If IsNumeric(lblAnnLmt) Then
            If CDbl(lblAnnLmt) = 0 Then
                lblBalLmt = "FULL"
            End If
        ElseIf lblAnnLmt = "FULL" Then
            lblBalLmt = "FULL"
        End If
    ElseIf tmpLimitStatus = "E" And CovID = "00" Then
        If IsNumeric(lblAnnLmt) Then
            If CDbl(lblAnnLmt) = 0 Then
                lblBalLmt = "FULL"
            End If
        ElseIf lblAnnLmt = "FULL" Then
            lblBalLmt = "FULL"
        End If
    End If
    If isPrinDepGPShared Then
        lblAnnLmt = "N/A"
        lblBalLmt = "N/A"
        lblAccumAmt = "N/A"
    Else
        lblAnnLmt = Format(lblAnnLmt, "#,##0.00")
        lblBalLmt = Format(lblBalLmt, "#,##0.00")
        lblAccumAmt = Format(lblAccumAmt, "#,##0.00")
    End If
    
    If Trim(txtGrpCompany) = "MERCHANTRADE ASIA SDN BHD" And IsNumeric(lblAnnLmt) And Trim(lblPLNCode) = "M230" Then
        Dim SETMERCHANTRADE As New ADODB.Recordset
        Dim cmd1 As New ADODB.Command
        Dim Prmmer As ADODB.Parameter
        
        Set cmd1.ActiveConnection = MediConnCISDB
        ''spMT_temp
        cmd1.CommandText = "spMERCHANTRADE"
        cmd1.CommandType = adCmdStoredProc
        cmd1.CommandTimeout = 1500
        Set Prmmer = cmd1.CreateParameter("mbmno", adVarChar, adParamInput, 2000, Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1)))
        cmd1.Parameters.Append Prmmer
        SETMERCHANTRADE.CursorLocation = adUseClient
        SETMERCHANTRADE.CursorType = adOpenForwardOnly
        SETMERCHANTRADE.CacheSize = 100
        Set SETMERCHANTRADE = cmd1.Execute
        If SETMERCHANTRADE.EOF = False Then
            If SETMERCHANTRADE!Company = "MERCHANTRADE ASIA SDN BHD" Then
                lblBalLmt = IIf(Len(SETMERCHANTRADE!BalLimit), Format(SETMERCHANTRADE!BalLimit, "#,##0.00"), "")
            End If
        End If
        SETMERCHANTRADE.Close
        Set SETMERCHANTRADE = Nothing
        MsgBox "Current Month Balance Limit For Plan Code M230. Amount RM " & Format((CDbl(lblBalLmt)), "#,##0.00") & " ", vbOKOnly + vbInformation, DialogBoxTitle
    End If
    '------Assigining  Calculated Values-------
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub CallLimit(tmpMBMTYpe As String, tmpAnnLmt As String, tmpVisitLmt As String, tmpSpecLmt As String, tmpLimitStatus As String)
Const METHOD_NAME = "CallLimit"
On Error GoTo debugs
    Dim sArr() As String
    Dim MBMNo As String
    Dim CovID As String
    
    Dim ASOqry As String
    Dim ASORS As New ADODB.Recordset
    Dim ASOlimit As Double
    Dim IPLimit As Double
    Dim OPLimit As Double
    Dim IPUtilization As Double
    Dim OPUtilization As Double
    Dim IsCombined As String
    
    If ASOInd = "Y" Then
        If Len(txtMBMNo) Then
           If InStr(1, txtMBMNo, "-") Then
              sArr = Split(txtMBMNo, "-")
              MBMNo = sArr(0)
              CovID = sArr(1)
           Else
              MBMNo = txtMBMNo
              CovID = "00"
           End If
        End If
        
        ASOqry = ""
        ASOqry = "Select MBMASOLimit,MBMIPLimit,MBMOPLimit,MBMIPUtilized,MBMOPUtilized ,IsIPOPCombined  from Tbl_ASOMemberDetails  where MBMNumber= '" + MBMNo + "' and MBMCoverID='" + CovID + "'"
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
            tmpAnnLmt = ASOlimit
            UtilisedAmt = IPUtilization + OPUtilization
        Else
            tmpAnnLmt = OPLimit
            UtilisedAmt = OPUtilization
        End If
    Else
        tmpAnnLmt = ""
        tmpAnnLmt = Format(GetAnnualLimit("H", lblPLNCode, "N", Format(TmpEffDate), tmpMBMTYpe, "DC", tmpLimitStatus), "#,##0.00")
    End If
   
    tmpSpecLmt = ""
    tmpSpecLmt = Format(GetSpecLimit("H", lblPLNCode, "N", Format(TmpEffDate), tmpMBMTYpe, "DC"), "#,##0.00")

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub funcword()
Const METHOD_NAME = "funcword"
On Error GoTo debugs
On Error GoTo openerror
Dim LetterHead As String
Dim AnnSpecLimitrst2 As New ADODB.Recordset
Dim SqlQry As String
Dim X As Boolean
Dim MBMCNUmber As String

    If Len(txtMBMNo) Then
        If InStr(txtMBMNo, "-") > 0 Then
           MBMCNUmber = Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1))
        Else
           MBMCNUmber = Trim(txtMBMNo)
        End If
    End If
  
  
 SqlQry = "Select MBMGroupCompany,productname from HISDB..MBM inner join HISDB..PLN on PLN.PLNCode=MBM.PLNCode inner join MBMOthers ON MBMOthers.MBMNumber=MBM.MBMNumber where MBMGroupCompany='" + txtGrpCompany + "' and mbm.MBMNumber = '" + MBMCNUmber + "'"
 AnnSpecLimitrst2.Open SqlQry, medixmasterconn, adOpenForwardOnly, adLockReadOnly
 If AnnSpecLimitrst2.EOF = False Then
 tmpProdName = Trim(IIf(Len(AnnSpecLimitrst2!ProductName), AnnSpecLimitrst2!ProductName, ""))
 End If
 
    
    If cboLetterHead = "" Then
        MsgBox "Please Select a Letter Head", vbOKOnly + vbCritical, DialogBoxTitle
        cboLetterHead.SetFocus
        Exit Sub
    Else
        If tmpGLFormat = "H" And Trim(txtGrpCompany) = Trim("MALAYSIA CONVENTION AND EXHIBITION BUREAU") Then
             If Len(tmpProdName) And Trim(tmpProdName) = Trim("PLAN 1") Then
                LetterHead = "UHCGuarantee Letter2_MCEB_Plan1.dot"
             ElseIf Len(tmpProdName) And Trim(tmpProdName) <> Trim("PLAN 1") Then
                LetterHead = "UHCGuarantee Letter2_MCEB_AllPlans.dot"
             End If
        ElseIf tmpGLFormat = "M" And Trim(txtGrpCompany) = Trim("MALAYSIA CONVENTION AND EXHIBITION BUREAU") Then
             If Len(tmpProdName) And Trim(tmpProdName) = Trim("PLAN 1") Then
                LetterHead = "UHCGuarantee Letter2_MCEB_Plan1.dot"
             ElseIf Len(tmpProdName) And Trim(tmpProdName) <> Trim("PLAN 1") Then
                LetterHead = "UHCGuarantee Letter2_MCEB_AllPlans.dot"
             End If
        ElseIf tmpGLFormat = "E" And Trim(txtGrpCompany) = Trim("MALAYSIA CONVENTION AND EXHIBITION BUREAU") Then
             If Len(tmpProdName) And Trim(tmpProdName) = Trim("PLAN 1") Then
                LetterHead = "UHCGuarantee Letter2_MCEB_Plan1.dot"
             ElseIf Len(tmpProdName) And Trim(tmpProdName) <> Trim("PLAN 1") Then
                LetterHead = "UHCGuarantee Letter2_MCEB_AllPlans.dot"
             End If
        ElseIf tmpGLFormat = "G" And Trim(txtGrpCompany) = Trim("MALAYSIA CONVENTION AND EXHIBITION BUREAU") Then
             If Len(tmpProdName) And Trim(tmpProdName) = Trim("PLAN 1") Then
                LetterHead = "UHCGuarantee Letter2_MCEB_Plan1.dot"
             ElseIf Len(tmpProdName) And Trim(tmpProdName) <> Trim("PLAN 1") Then
                LetterHead = "UHCGuarantee Letter2_MCEB_AllPlans.dot"
             End If
        ElseIf tmpGLFormat = "H" And Trim(txtGrpCompany) = Trim("INTERNATIONAL RESCUE COMMITTEE") Then
            LetterHead = "UHCGuarantee Letter2_IRC.dot"
        ElseIf tmpGLFormat = "E" And Trim(txtGrpCompany) = Trim("INTERNATIONAL RESCUE COMMITTEE") Then
            LetterHead = "UHCGuarantee Letter2_IRC.dot"
        ElseIf tmpGLFormat = "M" And Trim(txtGrpCompany) = Trim("INTERNATIONAL RESCUE COMMITTEE") Then
            LetterHead = "UHCGuarantee Letter2_IRC.dot"
        ElseIf tmpGLFormat = "G" And Trim(txtGrpCompany) = Trim("INTERNATIONAL RESCUE COMMITTEE") Then
            LetterHead = "UHCGuarantee Letter2_IRC.dot"
        ElseIf tmpGLFormat = "H" And Trim(txtGrpCompany) = Trim("PELABUHAN TANJUNG PELEPAS SDN BHD") Then
            LetterHead = "UHCGuarantee Letter2_PTP.dot"
        ElseIf tmpGLFormat = "E" And Trim(txtGrpCompany) = Trim("PELABUHAN TANJUNG PELEPAS SDN BHD") Then
            LetterHead = "UHCGuarantee Letter2_PTP.dot"
        ElseIf tmpGLFormat = "M" And Trim(txtGrpCompany) = Trim("PELABUHAN TANJUNG PELEPAS SDN BHD") Then
            LetterHead = "UHCGuarantee Letter2_PTP.dot"
        ElseIf tmpGLFormat = "G" And Trim(txtGrpCompany) = Trim("PELABUHAN TANJUNG PELEPAS SDN BHD") Then
            LetterHead = "UHCGuarantee Letter2_PTP.dot"
        ElseIf tmpGLFormat = "H" And Trim(txtGrpCompany) = Trim("FELDA GLOBAL VENTURES HOLDINGS BERHAD") Then
            LetterHead = "MEDGuarantee Letter_FGVHH.dot"
        ElseIf tmpGLFormat = "E" And Trim(txtGrpCompany) = Trim("FELDA GLOBAL VENTURES HOLDINGS BERHAD") Then
            LetterHead = "MEDGuarantee Letter_FGVH.dot"
        ElseIf tmpGLFormat = "M" And Trim(txtGrpCompany) = Trim("FELDA GLOBAL VENTURES HOLDINGS BERHAD") Then
            LetterHead = "MEDGuarantee Letter_FGVH.dot"
        ElseIf tmpGLFormat = "G" And Trim(txtGrpCompany) = Trim("FELDA GLOBAL VENTURES HOLDINGS BERHAD") Then
            LetterHead = "MEDGuarantee Letter_FGVHH.dot"
        ElseIf tmpGLFormat = "H" And Trim(txtGrpCompany) = Trim("PRINCE COURT MEDICAL CENTRE SDN BHD") Then
            LetterHead = "UHCGuarantee Letter_PCMC.dot"
        ElseIf tmpGLFormat = "E" And Trim(txtGrpCompany) = Trim("PRINCE COURT MEDICAL CENTRE SDN BHD") Then
            LetterHead = "UHCGuarantee Letter_PCMC.dot"
        ElseIf tmpGLFormat = "M" And Trim(txtGrpCompany) = Trim("PRINCE COURT MEDICAL CENTRE SDN BHD") Then
            LetterHead = "UHCGuarantee Letter_PCMC.dot"
        ElseIf tmpGLFormat = "G" And Trim(txtGrpCompany) = Trim("PRINCE COURT MEDICAL CENTRE SDN BHD") Then
            LetterHead = "UHCGuarantee Letter_PCMC.dot"
        ElseIf tmpGLFormat = "E" Then
            LetterHead = "GLMaybankEtiqaOP.dot"
        ElseIf tmpGLFormat = "M" Then
            LetterHead = "MEDGuarantee Letter2.dot"
        ElseIf tmpGLFormat = "H" Then
            LetterHead = "UHCGuarantee Letter2.dot"
        ElseIf tmpGLFormat = "G" Then
            LetterHead = "UHCGuarantee Letter2.dot"
            AnnSpecLimitrst2.Close
           Set AnnSpecLimitrst2 = Nothing
        End If
    End If

    If Letter = "L" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\CIS\LG\" & LetterHead)
        End If
    ElseIf Letter = "R" Then
        Set objword = Nothing
        LGNo = lblLG
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                Set objdoc = objword.Documents.Add(LocCardPath & "\CIS\LG\" & LetterHead)
        ElseIf Not objdoc Is Nothing Then
              If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
                vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
              ElseIf vbCancel Then
                Exit Sub
              End If
                Set objdoc = Nothing
                Set objdoc = objword.Documents.Add(LocCardPath & "Template\" & LetterHead & "\")
        End If
    End If
    
    If bookmark(LetterHead) = True Then
        Me.Show
        objdoc.SaveAs2 Filename:="\\Svr-doc\ScannedDocOP\" & LGNo & ".doc", FileFormat:=Word.wdFormatDocument
        
        Dim path As String
        path = "\\Svr-doc\ScannedDoc\ScannedDocOP\LG\" & LGNo
        X = CreatePath(path)
        objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
        objword.ActiveDocument.SaveAs Filename:=path & "\" & Trim(LGNo) & ".pdf", FileFormat:=Word.wdFormatPDF
        
    End If
    
        If txtTotalExcess <> "" And txtTotalExcess <> "0" Then
            LetterHead = "ExcessLetter.dot"
            Set objdoc = objword.Documents.Add(LocCardPath & "\CIS\LG\" & LetterHead)
        End If
   
    Call CopyFile(LGNo)
    Exit Sub
    
openerror:
    Me.Show
    If Err.Number = "5356" Then
        MsgBox Err.Description & "Please close the previous document first using Word or save it to other name", vbExclamation + vbOKOnly, DialogBoxTitle
    ElseIf Err.Number = "462" Then
        Set objword = New Word.Application
        objword.Visible = True
    ElseIf Err.Number = "4198" Then
        Err.Clear
    Else
        MsgBox Err.Description & Err.Number, vbExclamation + vbOKOnly, DialogBoxTitle
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Public Function CreatePath(NewPath) As Boolean
    'Add a trailing slash if none
    If Right(NewPath, 1) <> "\" Then
        NewPath = NewPath & "\"
    End If
    
    'Call API
    If MakeSureDirectoryPathExists(NewPath) <> 0 Then
        'No errors, return True
        CreatePath = True
    End If

End Function

Private Sub funcwordNew()
On Error GoTo openerror
Dim LetterHead As String
Dim AnnSpecLimitrst2 As New ADODB.Recordset
Dim SqlQry As String
Dim X As Boolean
Dim MBMCNUmber As String
Dim tmpclncode As String
Dim oRng As Word.Range
Dim oShape As Word.InlineShape
Dim oChart As Object

 
    Dim tmpMBMNum() As String
    Dim Pos As Integer
    Dim CovID As String
    CovID = "00"
    Pos = InStr(txtMBMNo, "-")
    
    If Pos > 0 Then
        tmpMBMNum = Split(txtMBMNo, "-")
        CovID = tmpMBMNum(1)
    End If
    
    'If CovID <> "00" Then
     '   If SuppInd = "Y" Then
      '      TemplateName = SuppTempName
       ' End If
    'End If
    
    If tmpCoverID_template <> "00" Then
        If SuppInd = "Y" Then
            TemplateName = SuppTempName
        End If
    End If
    

    If TemplateName = "" Then
        MsgBox ("Please Check LG Maintainace. Unable to find LG Template for Plan Code:" + lblPLNCode + " and Clinic GL Format : " + tmpGLFormat)
    Else
        LetterHead = TemplateName
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = False
            Set objdoc = objword.Documents.Add("\\Svr-doc\Template\CIS LG\" & LetterHead)
        End If
    End If
  
  If bookmark1(LetterHead) = True Then
        ''############################
        ''Additional Page code start
            Dim WordApp As Word.Application
            Dim BaseDoc As Word.Document
            Set WordApp = New Word.Application
            WordApp.Visible = False
            If Mid(txtMBMNo, 1, 2) = "AH" Then
                Set BaseDoc = WordApp.Documents.Add("\\Svr-doc\Template\CIS LG\" & "AH-AdditionalPage.docx")
            ElseIf (Trim(txtGrpCompany) = "SARAWAK ENERGY BERHAD" Or Trim(txtGrpCompany) = "SARAWAK ENERGY BERHAD - RETIREE") And (cmbGLType.Text = "Fitness to Work- Group B" Or cmbGLType.Text = "Fitness to Work- Group A" Or cmbGLType.Text = "Health Screening" Or cmbGLType.Text = "Pre-Employee Checkup") Then
                Set BaseDoc = WordApp.Documents.Add("\\Svr-doc\Template\CIS LG\" & "SEB Health Assessment Form.docx")
            Else
                Set BaseDoc = WordApp.Documents.Add("\\Svr-doc\Template\CIS LG\" & "AdditionalPage.docx")
            End If
            BaseDoc.Repaginate
            WordApp.Selection.EndKey (6)
            WordApp.Selection.HomeKey 6, 1
            WordApp.Selection.Copy
            DoEvents
            objword.Selection.EndKey (6)
            objword.Selection.InsertBreak
            objword.Selection.Paste
        ''Additional Page code end
        ''############################
        
        Me.Show
        'Dim path As String
        path = "\\Svr-doc\Outpatient\Guarantee Letter\" & LGNo
        X = CreatePath(path)
        FinalPath = ""
        FinalPath = path & "\L" & Trim(LGNo) & ".pdf"
        objword.ActiveDocument.SaveAs Filename:=FinalPath, FileFormat:=Word.wdFormatPDF
        MsgBox ("LG Generated Successfully")
        
        objword.Documents.Close (wdDoNotSaveChanges)
        objword.Quit (wdDoNotSaveChanges)
        Set objword = Nothing
        
        WordApp.Documents.Close (wdDoNotSaveChanges)
        WordApp.Quit (wdDoNotSaveChanges)
        Set WordApp = Nothing
        
        Shell "C:\WINDOWS\explorer.exe """ & path & "L" & Trim(LGNo) & ".pdf" & "", vbNormalFocus
    End If
        If txtTotalExcess <> "" And txtTotalExcess <> "0" Then
            LetterHead = "ExcessLetter.dot"
            Set objdoc = objword.Documents.Add(LocCardPath & "\CIS\LG\" & LetterHead)
        End If
    Exit Sub
openerror:
    Me.Show
    If Err.Number = "5356" Then
        MsgBox Err.Description & "Please close the previous document first using Word or save it to other name", vbExclamation + vbOKOnly, DialogBoxTitle
    ElseIf Err.Number = "462" Then
        Set objword = New Word.Application
        objword.Visible = True
    ElseIf Err.Number = "4198" Then
        Err.Clear
    Else
        MsgBox Err.Description & Err.Number, vbExclamation + vbOKOnly, DialogBoxTitle
    End If
End Sub
Private Function AHbookmark(Optional strAppend As String) As Integer
Dim AHCLN As New ADODB.Recordset
Dim strAppendCase As String
Dim strAppend1 As String
Dim cmd1 As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
AHbookmark = False

    strAppendCase = "26"
    If Len(LGNum) Then
        strAppend1 = " AND CISCas.CASNumber = '" & Trim(LGNum) & "'"
'    Else
'        StrAppend1 = " AND CISCas.CASNumber = '" & Trim(LGNo) & "'"
    End If
    Set cmd1.ActiveConnection = medixmasterconn
    cmd1.CommandText = "Case_Adjustment"
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandTimeout = 900
    Set Prm1 = cmd1.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend1)
    cmd1.Parameters.Append Prm1
    Set prm2 = cmd1.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd1.Parameters.Append prm2
    AHCLN.CursorLocation = adUseClient
    AHCLN.CursorType = adOpenForwardOnly
    AHCLN.CacheSize = 100
    Set AHCLN = cmd1.Execute
    
    
    'PATIENT_NAME
    If objdoc.Bookmarks.Exists("PATIENT_NAME") = True Then
        objdoc.Bookmarks("PATIENT_NAME").Range.Text = IIf(Len(AHCLN!PatientName), AHCLN!PatientName, "")
    End If
    If objdoc.Bookmarks.Exists("PATIENT_NAME1") = True Then
        objdoc.Bookmarks("PATIENT_NAME1").Range.Text = IIf(Len(AHCLN!PatientName), AHCLN!PatientName, "")
    End If
    If objdoc.Bookmarks.Exists("PATIENT_NAME2") = True Then
        objdoc.Bookmarks("PATIENT_NAME2").Range.Text = IIf(Len(AHCLN!PatientName), AHCLN!PatientName, "")
    End If
    'PATIENT_ICNO
    If objdoc.Bookmarks.Exists("PATIENT_ICNO") = True Then
        objdoc.Bookmarks("PATIENT_ICNO").Range.Text = IIf(Len(AHCLN!PatientIC), AHCLN!PatientIC, "")
    End If
    If objdoc.Bookmarks.Exists("PATIENT_ICNO1") = True Then
        objdoc.Bookmarks("PATIENT_ICNO1").Range.Text = IIf(Len(AHCLN!PatientIC), AHCLN!PatientIC, "")
    End If
    If objdoc.Bookmarks.Exists("PATIENT_ICNO2") = True Then
        objdoc.Bookmarks("PATIENT_ICNO2").Range.Text = IIf(Len(AHCLN!PatientIC), AHCLN!PatientIC, "")
    End If
    'EMP_NAME
    If objdoc.Bookmarks.Exists("EMP_NAME") = True Then
        objdoc.Bookmarks("EMP_NAME").Range.Text = IIf(Len(AHCLN!Principalname), AHCLN!Principalname, "")
    End If
    If objdoc.Bookmarks.Exists("EMP_NAME1") = True Then
        objdoc.Bookmarks("EMP_NAME1").Range.Text = IIf(Len(AHCLN!Principalname), AHCLN!Principalname, "")
    End If
    If objdoc.Bookmarks.Exists("EMP_NAME2") = True Then
        objdoc.Bookmarks("EMP_NAME2").Range.Text = IIf(Len(AHCLN!Principalname), AHCLN!Principalname, "")
    End If
    'EMPLOYEE_ICNO
     If objdoc.Bookmarks.Exists("EMPLOYEE_ICNO") = True Then
        objdoc.Bookmarks("EMPLOYEE_ICNO").Range.Text = IIf(Len(AHCLN!Principalic), AHCLN!Principalic, "")
    End If
    If objdoc.Bookmarks.Exists("EMPLOYEE_ICNO1") = True Then
        objdoc.Bookmarks("EMPLOYEE_ICNO1").Range.Text = IIf(Len(AHCLN!Principalic), AHCLN!Principalic, "")
    End If
    If objdoc.Bookmarks.Exists("EMPLOYEE_ICNO2") = True Then
        objdoc.Bookmarks("EMPLOYEE_ICNO2").Range.Text = IIf(Len(AHCLN!Principalic), AHCLN!Principalic, "")
    End If
    'MEMBERSHIP_NO
     If objdoc.Bookmarks.Exists("MEMBERSHIP_NO") = True Then
        objdoc.Bookmarks("MEMBERSHIP_NO").Range.Text = IIf(Len(AHCLN!MemberNumber), AHCLN!MemberNumber, "")
    End If
    If objdoc.Bookmarks.Exists("MEMBERSHIP_NO1") = True Then
        objdoc.Bookmarks("MEMBERSHIP_NO1").Range.Text = IIf(Len(AHCLN!MemberNumber), AHCLN!MemberNumber, "")
    End If
    If objdoc.Bookmarks.Exists("MEMBERSHIP_NO2") = True Then
        objdoc.Bookmarks("MEMBERSHIP_NO2").Range.Text = IIf(Len(AHCLN!MemberNumber), AHCLN!MemberNumber, "")
    End If
    'PAYOR_NAME
    If objdoc.Bookmarks.Exists("PAYOR_NAME") = True Then
        objdoc.Bookmarks("PAYOR_NAME").Range.Text = IIf(Len(AHCLN!CompanyName), AHCLN!CompanyName, "")
    End If
    If objdoc.Bookmarks.Exists("PAYOR_NAME1") = True Then
        objdoc.Bookmarks("PAYOR_NAME1").Range.Text = IIf(Len(AHCLN!CompanyName), AHCLN!CompanyName, "")
    End If
    If objdoc.Bookmarks.Exists("PAYOR_NAME2") = True Then
        objdoc.Bookmarks("PAYOR_NAME2").Range.Text = IIf(Len(AHCLN!CompanyName), AHCLN!CompanyName, "")
    End If
    'CLAIM_ID
    If objdoc.Bookmarks.Exists("CLAIM_ID") = True Then
        objdoc.Bookmarks("CLAIM_ID").Range.Text = IIf(Len(AHCLN!CaseNo), AHCLN!CaseNo, "")
    End If
    If objdoc.Bookmarks.Exists("CLAIM_ID1") = True Then
        objdoc.Bookmarks("CLAIM_ID2").Range.Text = IIf(Len(AHCLN!CaseNo), AHCLN!CaseNo, "")
    End If
    If objdoc.Bookmarks.Exists("CLAIM_ID2") = True Then
        objdoc.Bookmarks("CLAIM_ID2").Range.Text = IIf(Len(AHCLN!CaseNo), AHCLN!CaseNo, "")
    End If
     If objdoc.Bookmarks.Exists("CLAIM_ID3") = True Then
        objdoc.Bookmarks("CLAIM_ID3").Range.Text = IIf(Len(AHCLN!CaseNo), AHCLN!CaseNo, "")
    End If
    If objdoc.Bookmarks.Exists("CLAIM_ID4") = True Then
        objdoc.Bookmarks("CLAIM_ID4").Range.Text = IIf(Len(AHCLN!CaseNo), AHCLN!CaseNo, "")
    End If
    If objdoc.Bookmarks.Exists("CLAIM_ID5") = True Then
        objdoc.Bookmarks("CLAIM_ID5").Range.Text = IIf(Len(AHCLN!CaseNo), AHCLN!CaseNo, "")
    End If
    'PRINT_DATE
    If objdoc.Bookmarks.Exists("PRINT_DATE") = True Then
        objdoc.Bookmarks("PRINT_DATE").Range.Text = CStr(Format(DateTime.Now, "DD/MM/YYYY"))
    End If
    If objdoc.Bookmarks.Exists("PRINT_DATE1") = True Then
        objdoc.Bookmarks("PRINT_DATE1").Range.Text = CStr(Format(DateTime.Now, "DD/MM/YYYY"))
    End If
    If objdoc.Bookmarks.Exists("PRINT_DATE2") = True Then
        objdoc.Bookmarks("PRINT_DATE2").Range.Text = CStr(Format(DateTime.Now, "DD/MM/YYYY"))
    End If
    'PRINT_TIME
    If objdoc.Bookmarks.Exists("PRINT_TIME") = True Then
        objdoc.Bookmarks("PRINT_TIME").Range.Text = Format(Time, "hh:mm")
    End If
    If objdoc.Bookmarks.Exists("PRINT_TIME1") = True Then
        objdoc.Bookmarks("PRINT_TIME1").Range.Text = Format(Time, "hh:mm")
    End If
    If objdoc.Bookmarks.Exists("PRINT_TIME2") = True Then
        objdoc.Bookmarks("PRINT_TIME2").Range.Text = Format(Time, "hh:mm")
    End If
    'POLICY_NO
     If objdoc.Bookmarks.Exists("POLICY_NO") = True Then
        objdoc.Bookmarks("POLICY_NO").Range.Text = IIf(Len(AHCLN!Policyno), AHCLN!Policyno, "")
    End If
    If objdoc.Bookmarks.Exists("POLICY_NO1") = True Then
        objdoc.Bookmarks("POLICY_NO1").Range.Text = IIf(Len(AHCLN!Policyno), AHCLN!Policyno, "")
    End If
    If objdoc.Bookmarks.Exists("POLICY_NO2") = True Then
        objdoc.Bookmarks("POLICY_NO2").Range.Text = IIf(Len(AHCLN!Policyno), AHCLN!Policyno, "")
    End If
    'RELATION
    If objdoc.Bookmarks.Exists("RELATION") = True Then
        objdoc.Bookmarks("RELATION").Range.Text = IIf(Len(AHCLN!Relation), AHCLN!Relation, "")
    End If
    If objdoc.Bookmarks.Exists("RELATION1") = True Then
        objdoc.Bookmarks("RELATION1").Range.Text = IIf(Len(AHCLN!Relation), AHCLN!Relation, "")
    End If
    If objdoc.Bookmarks.Exists("RELATION2") = True Then
        objdoc.Bookmarks("RELATION2").Range.Text = IIf(Len(AHCLN!Relation), AHCLN!Relation, "")
    End If
    'CARD_NO
    If objdoc.Bookmarks.Exists("CARD_NO") = True Then
        objdoc.Bookmarks("CARD_NO").Range.Text = "-"
    End If
    If objdoc.Bookmarks.Exists("CARD_NO1") = True Then
        objdoc.Bookmarks("CARD_NO1").Range.Text = "-"
    End If
    If objdoc.Bookmarks.Exists("CARD_NO2") = True Then
        objdoc.Bookmarks("CARD_NO2").Range.Text = "-"
    End If
    'PROVIDER_NAME
    If objdoc.Bookmarks.Exists("PROVIDER_NAME") = True Then
        objdoc.Bookmarks("PROVIDER_NAME").Range.Text = IIf(Len(AHCLN!CLNName), AHCLN!CLNName, "")
    End If
    If objdoc.Bookmarks.Exists("PROVIDER_NAME1") = True Then
        objdoc.Bookmarks("PROVIDER_NAME1").Range.Text = IIf(Len(AHCLN!CLNName), AHCLN!CLNName, "")
    End If
    If objdoc.Bookmarks.Exists("PROVIDER_NAME2") = True Then
        objdoc.Bookmarks("PROVIDER_NAME2").Range.Text = IIf(Len(AHCLN!CLNName), AHCLN!CLNName, "")
    End If
    'PHYSICIAN_NAME
     If objdoc.Bookmarks.Exists("PHYSICIAN_NAME") = True Then
        objdoc.Bookmarks("PHYSICIAN_NAME").Range.Text = txtDoctor
    End If
    If objdoc.Bookmarks.Exists("PHYSICIAN_NAME1") = True Then
        objdoc.Bookmarks("PHYSICIAN_NAME1").Range.Text = txtDoctor
    End If
    If objdoc.Bookmarks.Exists("PHYSICIAN_NAME2") = True Then
        objdoc.Bookmarks("PHYSICIAN_NAME2").Range.Text = txtDoctor
    End If
    'DIAGNOSIS
     If objdoc.Bookmarks.Exists("DIAGNOSIS") = True Then
        objdoc.Bookmarks("DIAGNOSIS").Range.Text = IIf(Len(txtDiagnosis), txtDiagnosis, "")
    End If
    If objdoc.Bookmarks.Exists("DIAGNOSIS1") = True Then
        objdoc.Bookmarks("DIAGNOSIS1").Range.Text = IIf(Len(txtDiagnosis), txtDiagnosis, "")
    End If
    If objdoc.Bookmarks.Exists("DIAGNOSIS2") = True Then
        objdoc.Bookmarks("DIAGNOSIS2").Range.Text = IIf(Len(txtDiagnosis), txtDiagnosis, "")
    End If
    'PROVIDER_TELNO
    If objdoc.Bookmarks.Exists("PROVIDER_TELNO") = True Then
        objdoc.Bookmarks("PROVIDER_TELNO").Range.Text = IIf(Len(AHCLN!CLNPhone), AHCLN!CLNPhone, "")
    End If
    If objdoc.Bookmarks.Exists("PROVIDER_TELNO1") = True Then
        objdoc.Bookmarks("PROVIDER_TELNO1").Range.Text = IIf(Len(AHCLN!CLNPhone), AHCLN!CLNPhone, "")
    End If
    If objdoc.Bookmarks.Exists("PROVIDER_TELNO2") = True Then
        objdoc.Bookmarks("PROVIDER_TELNO2").Range.Text = IIf(Len(AHCLN!CLNPhone), AHCLN!CLNPhone, "")
    End If
    'PROVIDER_FAXNO
    If objdoc.Bookmarks.Exists("PROVIDER_FAXNO") = True Then
        objdoc.Bookmarks("PROVIDER_FAXNO").Range.Text = IIf(Len(AHCLN!CLNFax), AHCLN!CLNFax, "")
    End If
    If objdoc.Bookmarks.Exists("PROVIDER_FAXNO1") = True Then
        objdoc.Bookmarks("PROVIDER_FAXNO1").Range.Text = IIf(Len(AHCLN!CLNFax), AHCLN!CLNFax, "")
    End If
    If objdoc.Bookmarks.Exists("PROVIDER_FAXNO2") = True Then
        objdoc.Bookmarks("PROVIDER_FAXNO2").Range.Text = IIf(Len(AHCLN!CLNFax), AHCLN!CLNFax, "")
    End If
    'ADMISSION_DATE
    If objdoc.Bookmarks.Exists("ADMISSION_DATE") = True Then
        objdoc.Bookmarks("ADMISSION_DATE").Range.Text = IIf(Len(AHCLN!visitdate), AHCLN!visitdate, "")
    End If
    If objdoc.Bookmarks.Exists("ADMISSION_DATE1") = True Then
        objdoc.Bookmarks("ADMISSION_DATE1").Range.Text = IIf(Len(AHCLN!visitdate), AHCLN!visitdate, "")
    End If
    If objdoc.Bookmarks.Exists("ADMISSION_DATE2") = True Then
        objdoc.Bookmarks("ADMISSION_DATE2").Range.Text = IIf(Len(AHCLN!visitdate), AHCLN!visitdate, "")
    End If
    'AG_AMOUNT
    If objdoc.Bookmarks.Exists("AG_AMOUNT") = True Then
        objdoc.Bookmarks("AG_AMOUNT").Range.Text = IIf(Len(txtGLAmt), txtGLAmt, "")
    End If
    
    'PROVIDER_ADDRESS
     If objdoc.Bookmarks.Exists("PROVIDER_ADDRESS") = True Then
        objdoc.Bookmarks("PROVIDER_ADDRESS").Range.Text = IIf(Len(AHCLN!CLNAdd1), AHCLN!CLNAdd1, "")
    End If
    'ATTENTION_TO
    If objdoc.Bookmarks.Exists("ATTENTION_TO") = True Then
        objdoc.Bookmarks("ATTENTION_TO").Range.Text = ""
    End If
    'AUTHORIZATION_CODE
     If objdoc.Bookmarks.Exists("AUTHORIZATION_CODE") = True Then
        objdoc.Bookmarks("AUTHORIZATION_CODE").Range.Text = "NA"
    End If
    'DECLINED_DESCRIPTION
     If objdoc.Bookmarks.Exists("DECLINED_DESCRIPTION") = True Then
        objdoc.Bookmarks("DECLINED_DESCRIPTION").Range.Text = Trim(cmbDecline)
    End If
    'PUBLIC_REMARK
    If objdoc.Bookmarks.Exists("PUBLIC_REMARK") = True Then
        objdoc.Bookmarks("PUBLIC_REMARK").Range.Text = IIf(Len(txtInternalRemarks), txtInternalRemarks, "")
    End If
    
    AHbookmark = True
    AHCLN.Close
Set AHCLN = Nothing
End Function


Private Function bookmark1(Optional strAppend As String) As Integer
bookmark1 = False
If objdoc.Bookmarks.Exists("patient") = True And tmpMBMCategory = "PEMC" Then
    objdoc.Bookmarks("patient").Range.Text = Trim(tmpPatientName)
ElseIf objdoc.Bookmarks.Exists("patient") = True Then
    objdoc.Bookmarks("patient").Range.Text = StrConv(txtPatient, vbProperCase)
End If

If objdoc.Bookmarks.Exists("prepatient") = True And tmpMBMCategory = "PEMC" Then
    objdoc.Bookmarks("prepatient").Range.Text = Trim(tmpPatientName)
End If

If objdoc.Bookmarks.Exists("LGno") = True Then
    objdoc.Bookmarks("LGno").Range.Text = LGNo + "-" + Logon
End If
If objdoc.Bookmarks.Exists("LGID") = True Then
   objdoc.Bookmarks("LGID").Range.Text = ""
End If
If objdoc.Bookmarks.Exists("mbmnumber") = True Then
    objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNo
End If
If objdoc.Bookmarks.Exists("LGDate") = True Then
    objdoc.Bookmarks("LGDate").Range.Text = CStr(Format(DateTime.Now, "DD/MM/YYYY"))
End If
If objdoc.Bookmarks.Exists("EmpName") = True Then
    objdoc.Bookmarks("EmpName").Range.Text = tmpMBMName
End If

If objdoc.Bookmarks.Exists("dateofadmission") = True Then
    objdoc.Bookmarks("dateofadmission").Range.Text = txtDate
End If

If objdoc.Bookmarks.Exists("GrpCompany") = True Then
    objdoc.Bookmarks("GrpCompany").Range.Text = txtGrpCompany
End If

If objdoc.Bookmarks.Exists("hospital") = True Then
    objdoc.Bookmarks("hospital").Range.Text = cboCLNCode
End If

If objdoc.Bookmarks.Exists("Fax") = True Then
    objdoc.Bookmarks("Fax").Range.Text = IIf(Len(txtFaxNumber), txtFaxNumber, "")
End If

If objdoc.Bookmarks.Exists("Diagnosis") = True Then
    objdoc.Bookmarks("Diagnosis").Range.Text = txtDiagnosis
End If

If objdoc.Bookmarks.Exists("AttDr") = True Then
    objdoc.Bookmarks("AttDr").Range.Text = txtDoctor
End If

If objdoc.Bookmarks.Exists("Remark") = True Then
    objdoc.Bookmarks("Remark").Range.Text = txtRemarks
End If

If objdoc.Bookmarks.Exists("GLAmt") = True Then
 objdoc.Bookmarks("GLAmt").Range.Text = txtGLAmt
End If

If objdoc.Bookmarks.Exists("ICNum") = True And tmpMBMCategory = "PEMC" Then
 objdoc.Bookmarks("ICNum").Range.Text = tmpicno
ElseIf objdoc.Bookmarks.Exists("ICNum") = True Then
 objdoc.Bookmarks("ICNum").Range.Text = txtICNo
End If

If objdoc.Bookmarks.Exists("maternitynormal") = True And IsNumeric(txtmaternitynormal) < 0 Then
    objdoc.Bookmarks("maternitynormal").Range.Text = txtmaternitynormal
End If

If objdoc.Bookmarks.Exists("maternityCaesarean") = True And IsNumeric(txtCaesarean) < 0 Then
 objdoc.Bookmarks("maternityCaesarean").Range.Text = txtCaesarean
End If

If objdoc.Bookmarks.Exists("maternityComplications") = True And IsNumeric(txtComplications) < 0 Then
 objdoc.Bookmarks("maternityComplications").Range.Text = txtComplications
End If

If objdoc.Bookmarks.Exists("govtmpCLNName") = True Then
 objdoc.Bookmarks("govtmpCLNName").Range.Text = tmpCLNName
End If

If objdoc.Bookmarks.Exists("govtmpclnadd1") = True Then
 objdoc.Bookmarks("govtmpclnadd1").Range.Text = tmpclnadd1
End If

If objdoc.Bookmarks.Exists("govtmpclnadd2") = True Then
 objdoc.Bookmarks("govtmpclnadd2").Range.Text = tmpclnadd2
End If

If objdoc.Bookmarks.Exists("govtmpCLNState") = True Then
 objdoc.Bookmarks("govtmpCLNState").Range.Text = tmpCLNState
End If

If objdoc.Bookmarks.Exists("Subsidiary") = True And Len(tmpSubsidiary) Then
 objdoc.Bookmarks("Subsidiary").Range.Text = Trim(tmpSubsidiary)
End If

If objdoc.Bookmarks.Exists("hosfulladdress") = True Then
    Dim strclnfulladd As String
    strclnfulladd = strclnfulladd + tmpCLNName
    If tmpclnadd1 <> "" Then
        strclnfulladd = strclnfulladd + vbNewLine + ", " + tmpclnadd1
    End If
    If tmpclnadd2 <> "" Then
        strclnfulladd = strclnfulladd + vbNewLine + ", " + tmpclnadd2
    End If
    If tmpCLNState <> "" Then
        strclnfulladd = strclnfulladd + vbNewLine + ", " + tmpCLNState
    End If
    objdoc.Bookmarks("hosfulladdress").Range.Text = Trim(strclnfulladd)
End If

bookmark1 = True
End Function

Private Function bookmarkDecline(Optional strAppend As String) As Integer
bookmarkDecline = False
objdoc.Bookmarks("AdmissionDate").Range.Text = txtDate
objdoc.Bookmarks("CasNumber").Range.Text = LGNum
objdoc.Bookmarks("Date").Range.Text = CStr(Format(DateTime.Now, "DD/MM/YYYY"))
objdoc.Bookmarks("DeclineReason").Range.Text = cmbDecline
objdoc.Bookmarks("Employee").Range.Text = tmpMBMName
objdoc.Bookmarks("Hospital").Range.Text = cboCLNCode
objdoc.Bookmarks("ICNumber").Range.Text = txtICNo
objdoc.Bookmarks("IssuedBy").Range.Text = Logon
objdoc.Bookmarks("MBMNumber").Range.Text = txtMBMNo
objdoc.Bookmarks("patientname").Range.Text = StrConv(txtPatient, vbProperCase)
objdoc.Bookmarks("TreatmentDt").Range.Text = txtDate
bookmarkDecline = True
End Function

Private Function bookmark(Optional strAppend As String) As Integer
Dim CLN As New ADODB.Recordset
Dim CurrentDate, MyDate, MyDay As String
Dim StrSql As String
Dim Count As Integer
Dim tmpUpdWording1 As String
Dim tmpUpdWording2 As String
Dim tmptxtAmtWording As String
tmptxtAmtWording = ""
Dim wordArray() As String
Erase wordArray
Dim Y As Integer
Dim lenword As Integer
lenword = 0
    Count = InStr(1, cboCLNCode, "[")
    If Count <> 0 Then
        Count = Count - 1
    Else
        Count = Len(cboCLNCode)
    End If
    
    If Letter = "L" Or Letter = "R" Then
     bookmark = False
              If Trim(cboCLNCode) <> "" Then
                  objdoc.Bookmarks("CLNName").Range.Text = StrConv(Trim(Mid(cboCLNCode, 1, Count)), vbUpperCase)
              Else
                  objdoc.Bookmarks("CLNName").Range.Text = " Please Select a Clinic"
              End If
                  
              If Trim(LGNo) <> "" Then
                  objdoc.Bookmarks("LGNo").Range.Text = LGNo
              End If
              
              If Trim(txtMBMNo) <> "" Then
                  objdoc.Bookmarks("MBMNo").Range.Text = txtMBMNo
              End If
              
              If Trim(txtPatient) <> "" Then
                  objdoc.Bookmarks("Patient").Range.Text = StrConv(txtPatient, vbProperCase)
              End If
              
              If Trim(txtICNo) <> "" Then
                  objdoc.Bookmarks("ICNo").Range.Text = txtICNo
              End If
              
              If Trim(txtDoctor) <> "" Then
                  objdoc.Bookmarks("Doctor").Range.Text = StrConv(txtDoctor, vbProperCase)
              End If
              
              If Trim(cboIndicator) <> "" Then
                  objdoc.Bookmarks("ConsultType").Range.Text = Mid(cboIndicator, 5, Len(cboIndicator))
              End If
            
              If Trim(txtDiagnosis) <> "" Then
              
                  tmptxtAmtWording = Trim(txtDiagnosis)
                  tmpUpdWording1 = ""
                  tmpUpdWording2 = ""
                      
                  wordArray() = Split(tmptxtAmtWording, " ")
                  For Y = LBound(wordArray) To UBound(wordArray)
                      lenword = Len(tmpUpdWording1) + Len(tmpUpdWording2) + Len(wordArray(Y))
                      If lenword <= 60 Then
                          If Trim(tmpUpdWording1) = "" Then
                              tmpUpdWording1 = wordArray(Y)
                          Else
                              tmpUpdWording1 = tmpUpdWording1 + " " + wordArray(Y)
                          End If
                          tmpUpdWording1 = Trim(tmpUpdWording1)
                      Else
                          If Trim(tmpUpdWording2) = "" Then
                              tmpUpdWording2 = wordArray(Y)
                          Else
                              tmpUpdWording2 = tmpUpdWording2 + " " + wordArray(Y)
                          End If
                          tmpUpdWording2 = Trim(tmpUpdWording2)
                      End If
                  Next Y
              
                  If Trim(tmpUpdWording1) <> "" Then
                      objdoc.Bookmarks("Remarks").Range.Text = tmpUpdWording1
                  End If
                  
                  If Trim(tmpUpdWording2) <> "" Then
                      objdoc.Bookmarks("Remarks2").Range.Text = tmpUpdWording2
                  End If
              End If
              If Trim(txtGrpCompany) <> "" Then
                  objdoc.Bookmarks("GrpCompany").Range.Text = StrConv(txtGrpCompany, vbProperCase)
              End If
              
              MyDay = Format(txtDate, "dddd")
              MyDate = Format(txtDate, "dd mmmm yyyy")
              CurrentDate = Format(Date, "dd mmmm yyyy")
              
              If Len(MyDate) <> "" Then
                  objdoc.Bookmarks("Date1").Range.Text = MyDay & ", " & MyDate
              End If
              
              If Len(CurrentDate) <> "" Then
                  objdoc.Bookmarks("Date").Range.Text = CurrentDate
              End If
              
              If cboCLNCode <> "" Then
                  
                  StrSql = "Exec SearchCLN"
                  CLN.Open StrSql, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                  
                  Do Until CLN.EOF
                      With CLN
                          If StrComp(Trim(Mid(cboCLNCode, Count + 2, Len(cboCLNCode) - (Count + 2))), Trim(!CLNCode)) = 0 Then
                              If Trim(!CLNAdd1) <> "" Then
                                  objdoc.Bookmarks("Add1").Range.Text = StrConv(!CLNAdd1, vbProperCase)
                              End If
                              If Trim(!CLNAdd2) <> "" Then
                                  objdoc.Bookmarks("Add2").Range.Text = StrConv(!CLNAdd2, vbProperCase)
                              End If
                              If Trim(!CLNAdd3) <> "" Then
                                  objdoc.Bookmarks("Add3").Range.Text = StrConv(!CLNAdd3, vbProperCase)
                              End If
                              If Not !CLNRepresentative = "" Then
                                  objdoc.Bookmarks("ContactPerson").Range.Text = !CLNRepresentative
                              Else
                                  objdoc.Bookmarks("ContactPerson").Range.Text = " Please enter contact person's name "
                              End If
                              If Not !CLNFax = "" Then
                                  objdoc.Bookmarks("FaxNo").Range.Text = !CLNFax
                              End If
                              'Exit Function
                          End If
                          .MoveNext
                      End With
                  Loop
                  CLN.Close
                  Set CLN = Nothing
              End If
              
              
    
    
    
    
'        bookmark = False
'        'If tmpGLFormat = "E" Then
'          objdoc.Bookmarks("Date").Range.Text = Date
'              If Trim(LGNo) <> "" Then
'                  objdoc.Bookmarks("LGNo").Range.Text = LGNo & "O"
'              End If
'
'              If Trim(cboCLNCode) <> "" Then
'                  objdoc.Bookmarks("CLNName").Range.Text = StrConv(Trim(Mid(cboCLNCode, 1, Count)), vbUpperCase)
'              Else
'                  objdoc.Bookmarks("CLNName").Range.Text = " Please Select a Clinic"
'              End If
'
'              If Trim(txtPatient) <> "" Then
'                  objdoc.Bookmarks("Patient").Range.Text = StrConv(txtPatient, vbProperCase)
'              End If
'
'              If Trim(txtICNo) <> "" Then
'                  objdoc.Bookmarks("ICNo").Range.Text = txtICNo
'              End If
'
'              'objdoc.Bookmarks("ContactPerson").Range.Text = tmpMBMName
'
'              objdoc.Bookmarks("Date1").Range.Text = txtDate
'
'              If Trim(txtMBMNo) <> "" Then
'                  objdoc.Bookmarks("MBMNo").Range.Text = txtMBMNo
'              End If
''              If Trim(txtDoctor) <> "" Then
'                  objdoc.Bookmarks("Doctor").Range.Text = StrConv(txtDoctor, vbProperCase)
'              End If
'
'              'If Trim(cboIndicator) <> "" Then
'              '    objdoc.Bookmarks("ConsultType").Range.Text = Mid(cboIndicator, 5, Len(cboIndicator))
'              'End If
'
'              If Trim(txtDiagnosis) <> "" Then
'
'                  tmptxtAmtWording = Trim(txtDiagnosis)
'                  tmpUpdWording1 = ""
'                  tmpUpdWording2 = ""
'
'                  wordArray() = Split(tmptxtAmtWording, " ")
'                  For Y = LBound(wordArray) To UBound(wordArray)
'                      lenword = Len(tmpUpdWording1) + Len(tmpUpdWording2) + Len(wordArray(Y))
'                      If lenword <= 60 Then
'                          If Trim(tmpUpdWording1) = "" Then
'                              tmpUpdWording1 = wordArray(Y)
'                          Else
'                              tmpUpdWording1 = tmpUpdWording1 + " " + wordArray(Y)
'                          End If
'                          tmpUpdWording1 = Trim(tmpUpdWording1)
'                      Else
'                          If Trim(tmpUpdWording2) = "" Then
'                              tmpUpdWording2 = wordArray(Y)
'                          Else
'                              tmpUpdWording2 = tmpUpdWording2 + " " + wordArray(Y)
'                          End If
'                          tmpUpdWording2 = Trim(tmpUpdWording2)
'                      End If
'                  Next Y
'
'                    objdoc.Bookmarks("Diagnosis").Range.Text = tmpUpdWording1 & tmpUpdWording2
'                  'If Trim(tmpUpdWording1) <> "" Then
'                      objdoc.Bookmarks("Remark").Range.Text = ""
'                  'End If
'
'                  'If Trim(tmpUpdWording2) <> "" Then
'                   '   objdoc.Bookmarks("Remarks2").Range.Text = tmpUpdWording2
'                  'End If
'              End If
'
'              'If lblGrpCompany <> "" Then
'              '    objdoc.Bookmarks("GrpCompany").Range.Text = StrConv(lblGrpCompany, vbProperCase)
'              'End If
'
'              If Trim(txtGrpCompany) <> "" Then
'                  objdoc.Bookmarks("GrpCompany").Range.Text = StrConv(txtGrpCompany, vbProperCase)
'              End If
'
'
'              If cboCLNCode <> "" Then
'
'                  StrSql = "Exec SearchCLN"
'                  CLN.Open StrSql, medixmasterconn, adOpenForwardOnly, adLockOptimistic
'
'                  Do Until CLN.EOF
'                      With CLN
'                          If StrComp(Trim(Mid(cboCLNCode, Count + 2, Len(cboCLNCode) - (Count + 2))), Trim(!CLNCode)) = 0 Then
'                         '     If Trim(!CLNAdd1) <> "" Then
'                         '         objdoc.Bookmarks("Add1").Range.Text = StrConv(!CLNAdd1, vbProperCase)
'                         '     End If
'                         '     If Trim(!CLNAdd2) <> "" Then
'                         '         objdoc.Bookmarks("Add2").Range.Text = StrConv(!CLNAdd2, vbProperCase)
'                         '     End If
'                         '     If Trim(!CLNAdd3) <> "" Then
'                         '         objdoc.Bookmarks("Add3").Range.Text = StrConv(!CLNAdd3, vbProperCase)
'                         '     End If
'                              'If Not !CLNRepresentative = "" Then
'                              '    objdoc.Bookmarks("ContactPerson").Range.Text = !CLNRepresentative
'                              'Else
'                              '    objdoc.Bookmarks("ContactPerson").Range.Text = " Please enter contact person's name "
'                              'End If
'                              If Not !CLNFax = "" Then
'                                  objdoc.Bookmarks("Fax").Range.Text = !CLNFax
'                              End If
'                              'Exit Function
'                          End If
'                          .MoveNext
'                      End With
'                  Loop
'                  CLN.Close
'                  Set CLN = Nothing
'              End If
        
        Else
        
              If Trim(cboCLNCode) <> "" Then
                  objdoc.Bookmarks("CLNName").Range.Text = StrConv(Trim(Mid(cboCLNCode, 1, Count)), vbUpperCase)
              Else
                  objdoc.Bookmarks("CLNName").Range.Text = " Please Select a Clinic"
              End If
                  
              If Trim(LGNo) <> "" Then
                  objdoc.Bookmarks("LGNo").Range.Text = LGNo
              End If
              
              If Trim(txtMBMNo) <> "" Then
                  objdoc.Bookmarks("MBMNo").Range.Text = txtMBMNo
              End If
              
              If Trim(txtPatient) <> "" Then
                  objdoc.Bookmarks("Patient").Range.Text = StrConv(txtPatient, vbProperCase)
              End If
              
              If Trim(txtICNo) <> "" Then
                  objdoc.Bookmarks("ICNo").Range.Text = txtICNo
              End If
              
              If Trim(txtDoctor) <> "" Then
                  objdoc.Bookmarks("Doctor").Range.Text = StrConv(txtDoctor, vbProperCase)
              End If
              
              If Trim(cboIndicator) <> "" Then
                  objdoc.Bookmarks("ConsultType").Range.Text = Mid(cboIndicator, 5, Len(cboIndicator))
              End If
              
              
            
              If Trim(txtDiagnosis) <> "" Then
              
                  tmptxtAmtWording = Trim(txtDiagnosis)
                  tmpUpdWording1 = ""
                  tmpUpdWording2 = ""
                      
                  wordArray() = Split(tmptxtAmtWording, " ")
                  For Y = LBound(wordArray) To UBound(wordArray)
                      lenword = Len(tmpUpdWording1) + Len(tmpUpdWording2) + Len(wordArray(Y))
                      If lenword <= 60 Then
                          If Trim(tmpUpdWording1) = "" Then
                              tmpUpdWording1 = wordArray(Y)
                          Else
                              tmpUpdWording1 = tmpUpdWording1 + " " + wordArray(Y)
                          End If
                          tmpUpdWording1 = Trim(tmpUpdWording1)
                      Else
                          If Trim(tmpUpdWording2) = "" Then
                              tmpUpdWording2 = wordArray(Y)
                          Else
                              tmpUpdWording2 = tmpUpdWording2 + " " + wordArray(Y)
                          End If
                          tmpUpdWording2 = Trim(tmpUpdWording2)
                      End If
                  Next Y
              
                  If Trim(tmpUpdWording1) <> "" Then
                      objdoc.Bookmarks("Remarks").Range.Text = tmpUpdWording1
                  End If
                  
                  If Trim(tmpUpdWording2) <> "" Then
                      objdoc.Bookmarks("Remarks2").Range.Text = tmpUpdWording2
                  End If
              End If
              
              'If lblGrpCompany <> "" Then
              '    objdoc.Bookmarks("GrpCompany").Range.Text = StrConv(lblGrpCompany, vbProperCase)
              'End If
              
              If Trim(txtGrpCompany) <> "" Then
                  objdoc.Bookmarks("GrpCompany").Range.Text = StrConv(txtGrpCompany, vbProperCase)
              End If
              
              MyDay = Format(txtDate, "dddd")
              MyDate = Format(txtDate, "dd mmmm yyyy")
              CurrentDate = Format(Date, "dd mmmm yyyy")
              
              If Len(MyDate) <> "" Then
                  objdoc.Bookmarks("Date1").Range.Text = MyDay & ", " & MyDate
              End If
              
              If Len(CurrentDate) <> "" Then
                  objdoc.Bookmarks("Date").Range.Text = CurrentDate
              End If
              
              If cboCLNCode <> "" Then
                  
                  StrSql = "Exec SearchCLN"
                  CLN.Open StrSql, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                  
                  Do Until CLN.EOF
                      With CLN
                          If StrComp(Trim(Mid(cboCLNCode, Count + 2, Len(cboCLNCode) - (Count + 2))), Trim(!CLNCode)) = 0 Then
                              If Trim(!CLNAdd1) <> "" Then
                                  objdoc.Bookmarks("Add1").Range.Text = StrConv(!CLNAdd1, vbProperCase)
                              End If
                              If Trim(!CLNAdd2) <> "" Then
                                  objdoc.Bookmarks("Add2").Range.Text = StrConv(!CLNAdd2, vbProperCase)
                              End If
                              If Trim(!CLNAdd3) <> "" Then
                                  objdoc.Bookmarks("Add3").Range.Text = StrConv(!CLNAdd3, vbProperCase)
                              End If
                              If Not !CLNRepresentative = "" Then
                                  objdoc.Bookmarks("ContactPerson").Range.Text = !CLNRepresentative
                              Else
                                  objdoc.Bookmarks("ContactPerson").Range.Text = " Please enter contact person's name "
                              End If
                              If Not !CLNFax = "" Then
                                  objdoc.Bookmarks("FaxNo").Range.Text = !CLNFax
                              End If
                              'Exit Function
                          End If
                          .MoveNext
                      End With
                  Loop
                  CLN.Close
                  Set CLN = Nothing
              End If
        
         '   If Mid(cboLetterHead, 1, 3) = "MED" And tmpStaffCat = "NE" And Right(txtMBMNo, Len(txtMBMNo) - InStr(txtMBMNo, "-")) <> "00" Then
         '       objdoc.Bookmarks("RBAmt").Range.Text = "RM120.00"
         '       objdoc.Bookmarks("GLAmt").Range.Text = "RM" & txtGLAmt
         '   End If
        'end If
    End If
    
    bookmark = True
End Function

Private Function bookmarkExcess(Optional strAppend As String) As Integer
Dim CLN As New ADODB.Recordset
Dim CurrentDate, MyDate, MyDay As String
Dim StrSql As String
Dim Count As Integer
Dim tmpUpdWording1 As String
Dim tmpUpdWording2 As String
Dim tmptxtAmtWording As String
tmptxtAmtWording = ""
Dim wordArray() As String
Erase wordArray
Dim Y As Integer
Dim lenword As Integer
lenword = 0
        bookmarkExcess = False
        CurrentDate = Format(Date, "dd/mm/yyyy")
        
        
        If Len(CurrentDate) <> "" Then
            objdoc.Bookmarks("ExcDate").Range.Text = CurrentDate
        End If
        
        If Len(MyDate) <> "" Then
            objdoc.Bookmarks("ExcDate").Range.Text = MyDay & ", " & MyDate
        End If
        
        If Trim(txtPatient) <> "" Then
            objdoc.Bookmarks("Patient").Range.Text = StrConv(txtPatient, vbProperCase)
        End If
        
        If Trim(txtMBMNo) <> "" Then
            objdoc.Bookmarks("MBMNo").Range.Text = txtMBMNo
        End If
        
        If Trim(txtGrpCompany) <> "" Then
            objdoc.Bookmarks("GrpCompany").Range.Text = StrConv(txtGrpCompany, vbProperCase)
        End If
    
        If Trim(cboCLNCode) <> "" Then
            objdoc.Bookmarks("CLNName").Range.Text = StrConv(Trim(Mid(cboCLNCode, 1, Count)), vbUpperCase)
        Else
            objdoc.Bookmarks("CLNName").Range.Text = " Please Select a Clinic"
        End If
            
        If Len(txtRBExcess) Then
            objdoc.Bookmarks("RB").Range.Text = txtRBExcess
        End If

        If Len(txtRBCoPay) Then
            objdoc.Bookmarks("CoPay").Range.Text = txtRBCoPay
        End If

        If Len(txtRBExcess) Then
            objdoc.Bookmarks("OthExcess").Range.Text = txtRBExcess
        End If
        
        If Len(txtLmtExceeded) Then
            objdoc.Bookmarks("LmtExceeded").Range.Text = txtLmtExceeded
        End If
        
        If Len(txtTotalExcess) Then
            objdoc.Bookmarks("TotalExcess").Range.Text = txtTotalExcess
        End If
    'End If
    bookmarkExcess = True
End Function

Private Sub lstPolicyWording_Click()
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String ''

'PictureFileName = "\\mars\Outpatient\Guarantee Letter\" & CasNumber '

    
    PictureFileName = SRVPath & "CLNPolicyWording\" & tmpPolicyWording & ""
    crt PictureFileName & "\" & lstPolicyWording.List(lstPolicyWording.ListIndex)
    sFile = PictureFileName & "\" & lstPolicyWording.List(lstPolicyWording.ListIndex)
    lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
      
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
Const METHOD_NAME = "lstSpecialistCases_DblClick"
On Error GoTo debugs

    If lstSpecialistCases.ListItems.Count > 0 Then
        frmInvoiceAdj.txtCaseNo = Trim(lstSpecialistCases.SelectedItem)
        Call frmInvoiceAdj.CmdGo_Click
        Load frmInvoiceAdj
        frmInvoiceAdj.Show
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
Dim recCount, lstCount As Integer
Dim StrSql As String
Dim i
    lblCount = ""
    recCount = 0
    If EditFlag = True Then
        lstMBM.Enabled = False
    Else
        lstMBM.Enabled = True
    End If
    Screen.MousePointer = vbHourglass
    Select Case SSTab.Tab
    Case 0
        lstCount = lstMBM.ListItems.Count
        TempLG = ""
        tmpLGCasDate = ""
        tmpLGCasTime = ""
        If txtMBMNo <> "" Then
            If lstMBM.ListItems.Count = 0 Then
                cmdSearch_Click
            End If
        End If
        For i = 1 To lstCount
            recCount = recCount + 1
        Next
        'lblCount = recCountsaw
    Case 1
        Call LGListing
        lstCount = lstLG.ListItems.Count
        For i = 1 To lstCount
            recCount = recCount + 1
        Next
        lblCount = recCount
    Case 2
        If Trim(tmpPolicyWording) <> "" Then
            Call LoadDoc
        End If
    Case 3
        If Trim(txtMBMNo) <> "" And lstCaseList.ListItems.Count = 0 Then
            Call ListCases
            Call ListSpecialistCases
        End If
    End Select
    Screen.MousePointer = Default
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
    'strsql = strsql & " AND MBMNumber = '" & Left(txtMBMNo, Len(txtMBMNo) - InStr(txtMBMNo, "-")) & "' "
    
    If InStr(txtMBMNo, "-") > 0 Then
        StrSql = StrSql & " AND MBMNumber = '" & Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1)) & "'" & _
                    " AND CoverID = '" & Right(txtMBMNo, Len(txtMBMNo) - InStr(txtMBMNo, "-")) & "'"
    ElseIf Len(txtMBMNo) Then
        StrSql = StrSql & " AND MBMNumber = '" & Trim(txtMBMNo) & "'"
    End If
    
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
                
                If IsNumeric(Trim(CAS!CLMBordxAmt)) Then
                lblTotalBordx = CDbl(lblTotalBordx) + CDbl(CAS!CLMBordxAmt)
                Else
                    lblTotalBordx = CDbl(lblTotalBordx) + CDbl(0)
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
    If InStr(txtMBMNo, "-") > 0 Then
        StrSql = StrSql & " AND MBMNumber = '" & Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1)) & "'" & _
                    " AND CoverID = '" & Right(txtMBMNo, Len(txtMBMNo) - InStr(txtMBMNo, "-")) & "'"
    ElseIf Len(txtMBMNo) Then
        StrSql = StrSql & " AND MBMNumber = '" & Trim(txtMBMNo) & "'"
    End If
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
                lblTotalSpeVisits = CDbl(lblTotalSpeVisits) + 1
                lblTotalSpeBill = CDbl(lblTotalSpeBill) + CDbl(CAS!CLMTotalActual)
                lblTotalSpecBordx = CDbl(lblTotalSpecBordx) + CDbl(CAS!CLMBordxAmt)

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
    
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub LGListing()
Const METHOD_NAME = "LGListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim rst2 As New ADODB.Recordset
Dim CLN As New ADODB.Recordset
Dim PRM As ADODB.Parameter
Dim strAppend As String
Dim strsqlCLN As String
Dim Record As ListItem
Dim tmpCLN As String
Dim tmpstring As String
Dim strsqlPLN As String

If Len(txtMBMNo) Or lstMBM.ListItems.Count <> 0 Then
    lstLG.ListItems.Clear
    lstLG.Enabled = True
    If InStr(txtMBMNo, "-") > 0 Then
    tmpstring = Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1))
        strAppend = " AND LG.MBMNumber = '" & Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1)) & "'" & _
                    " AND LG.CoverID = '" & Right(txtMBMNo, Len(txtMBMNo) - InStr(txtMBMNo, "-")) & "'"
    ElseIf Len(txtMBMNo) Then
        strAppend = " AND LG.MBMNumber = '" & Trim(txtMBMNo) & "'"
    ElseIf lstMBM.ListItems.Count <> 0 Then
        strAppend = " AND LG.MBMNumber = '" & Mid(lstMBM.SelectedItem.Text, 1, InStr(lstMBM.SelectedItem.Text, "-") - 1) & "'"
    End If
    
    Set cmd.ActiveConnection = MediConnCISDB
    
    'cmd.CommandText = "LGlst"
    cmd.CommandText = "LGlst2020"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 9000
    Set PRM = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append PRM
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    Do Until rst2.EOF
        Set Record = lstLG.ListItems.Add(, , rst2!LGNo)
        If Len(rst2!MBMNumber) Then
            Record.SubItems(1) = rst2!MBMNumber
        End If
        If Len(rst2!LGdate) Then
            Record.SubItems(2) = rst2!LGdate
        End If
        If Len(rst2!LGIssuedBy) Then
            Record.SubItems(3) = rst2!LGIssuedBy
        End If
        If Len(rst2!DoctorName) Then
            Record.SubItems(4) = rst2!DoctorName
        End If
        If Len(rst2!CLNName) Then
            Record.SubItems(5) = rst2!CLNName
        End If
        If Len(rst2!GLAmt) Then
            Record.SubItems(6) = rst2!GLAmt
        End If
        If Len(rst2!Diagnosis) Then
            Record.SubItems(7) = rst2!Diagnosis
        End If
        If Len(rst2!CLNNameDesc) Then
            Record.SubItems(9) = rst2!CLNNameDesc
        End If
        If Len(rst2!CLNGLFormat) Then
            Record.SubItems(10) = IIf(Len(rst2!CLNGLFormat), rst2!CLNGLFormat, "")
        End If
        If Len(rst2!CaseProgress) Then
            Record.SubItems(12) = IIf(Len(rst2!CaseProgress), rst2!CaseProgress, "")
        End If
        
        If Len(rst2!LGRemarks) Then
            Record.SubItems(11) = IIf(Len(rst2!LGRemarks), rst2!LGRemarks, "")
        End If
        
        If Len(rst2!GLType) Then
            Record.SubItems(13) = IIf(Len(rst2!GLType), rst2!GLType, "")
        End If
        
        If Len(rst2!InternalRemarks) Then
            Record.SubItems(14) = IIf(Len(rst2!InternalRemarks), rst2!InternalRemarks, "")
        End If
        
        If Len(rst2!Alertmsg) Then
            Record.SubItems(15) = IIf(Len(rst2!Alertmsg), rst2!Alertmsg, "")
        End If
        
        If Len(rst2!CasDelayReason) Then
            Record.SubItems(16) = IIf(Len(rst2!CasDelayReason), rst2!CasDelayReason, "")
        End If
        
        If Len(rst2!LGTIME) Then
            Record.SubItems(17) = rst2!LGTIME
        End If
        
        tmpEmailLG = IIf(Len(rst2!EmailInd), rst2!EmailInd, "")
        rst2.MoveNext
    Loop
    rst2.Close
    Set rst2 = Nothing
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Public Sub MailAdmInitialLG(ToAddress As String, BCC As String, CC As String, LGNo As String, Subsidiary As String)
Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM
Dim objMessage As Object
On Error GoTo debugs
Set objMessage = CreateObject("CDO.Message")
                                           
objMessage.Subject = "Notification of Outpatient LG Issued To : " & txtPatient & " , Membership Number :" & txtMBMNo
'objMessage.From = "adm@medix.com.my"
objMessage.From = "notifications@medix.com.my"
'objMessage.To = "Norlaila@maa.my;Farihan@maa.my;Zamri@maa.my"
objMessage.AddAttachment FinalPath
objMessage.To = ToAddress
objMessage.BCC = BCC
objMessage.CC = CC
objMessage.TextBody = "Dear Sir/Madam," & vbCrLf & vbCrLf & _
                            "LG ISSUED NOTIFICATION : " & vbCrLf & vbCrLf & _
                            "Please be advised that the above-mentioned member is an Outpatient treatment (Check-up) at below hospital." & vbCrLf & vbCrLf & _
                            "Below is the case details for your kind attention." & vbCrLf & vbCrLf & _
                            "Patient Name             :       " & txtPatient & vbCrLf & vbCrLf & _
                            "Employee Name            :       " & tmpMBMName & vbCrLf & vbCrLf & _
                            "IC Number                :       " & txtICNo & vbCrLf & vbCrLf & _
                            "Group Company            :       " & txtGrpCompany & vbCrLf & vbCrLf & _
                            "Subsidiary               :       " & Subsidiary & vbCrLf & vbCrLf & _
                            "LG Number                :       " & LGNo & vbCrLf & vbCrLf & _
                            "Treatment Date           :       " & txtDate & vbCrLf & vbCrLf & _
                            "Hospital                 :       " & cboCLNCode & vbCrLf & vbCrLf & _
                            " " & vbCrLf & vbCrLf & _
                            "Please contact us at 603-7884 1919 if you require any further clarification." & vbCrLf & vbCrLf & _
                            " " & vbCrLf & vbCrLf & _
                            "Thank you." & vbCrLf & vbCrLf & _
                            "Regards, " & vbCrLf & vbCrLf & _
                            "Health Connect (Malaysia) Sdn Bhd" & vbCrLf & vbCrLf
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.mediexpress.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "autoemail@mediexpress.com.my"
'objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "adm@medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Medix@123"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objMessage.Configuration.Fields.update
objMessage.Send
debugs:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub







Private Sub txtDate_LostFocus()
    If IsDate(txtDate) Then
    Else
        If txtDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDate.SetFocus
        End If
    End If
End Sub


Private Sub txtDoctor_LostFocus()
    If Trim(txtDoctor) = "" Then
        txtDoctor = "Not Specified"
    Else
        txtDoctor = ChkStr(txtDoctor)
    End If
End Sub


Private Sub lstLG_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstLG, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstLG, ColumnHeader.Index, ldtString, m_blnDirection(18)
    End Select

End Sub

Private Sub SaveCase()
Const METHOD_NAME = "SaveCase"
On Error GoTo debugs
'Dim CISCAS As New ADODB.Recordset
Dim CISSql As String
Dim tmpDiag1, tmpDiag2, tmpDiag3, tmpDiag4, tmpDiag5 As String
Dim NoOfDiag As Integer
Dim NoofMedi As Integer
Dim MCStat As Integer
Dim tmpRBAmt, tmpRBCoPay, tmpLmtExceed
Dim tmpINVMonth, tmpINVYear As String

    CISSql = ""
    MCStat = 0
    NoOfDiag = 0
    NoofMedi = 0
    tmpRBAmt = 0
    tmpRBCoPay = 0
    tmpLmtExceed = 0
        tmpCASnumber = NewCaseNo
        tmpMBMNumber = Trim(Mid(txtMBMNo, 1, InStr(txtMBMNo, "-") - 1))
        tmpCoverID = Right(txtMBMNo, Len(txtMBMNo) - InStr(txtMBMNo, "-"))
        tmpCASStatus = "O"
        tmpClaimability = "AA"
        tmpCaseType = "C" ' C - Clinical Case   D - Dental Case
        TmpPending = "N"
        tmpRegisterDate = Format(Date, "yyyy/mm/dd")
        tmpREgisterBy = Logon
        tmpLastUpdateDate = Format(Date, "yyyy/mm/dd")
        tmpLastUpdateBy = Logon
        tmpClinicCode = Trim(Clinic)
        tmpInvoiceNo = ""
        tmpInvoiceDate = ""
        tmpInvoiceStmtAmt = 0
        tmpDateVisit = Format(TreatmentDate, "yyyy/mm/dd")
        tmpMCDays = 0
        tmpOPBatchNo = ""
        tmpTimeVisit = ""
        tmpMedInjType1 = ""
        tmpMedInjType2 = ""
        tmpMedInjType3 = ""
        tmpMedInjType4 = ""
        tmpMedInjType5 = ""
        tmpMedInjDays1 = ""
        tmpMedInjDays2 = ""
        tmpMedInjDays3 = ""
        tmpMedInjDays4 = ""
        tmpMedInjDays5 = ""
        tmpMedInjQty1 = ""
        tmpMedInjQty2 = ""
        tmpMedInjQty3 = ""
        tmpMedInjQty4 = ""
        tmpMedInjQty5 = ""
        tmpMedInjType6 = ""
        tmpMedInjType7 = ""
        tmpMedInjDays6 = ""
        tmpMedInjDays7 = ""
        tmpMedInjQty6 = ""
        tmpMedInjQty7 = ""
        tmpInvoiceReceiptDate = ""
        tmpInvoiceReceiptMode = ""
        tmpCasRemark = txtDoctor & "/" & txtDiagnosis
        tmpCASTreatReason1 = ""
        tmpCASTreatReason2 = ""
        tmpCASTreatReason3 = ""
        tmpCASTreatReason4 = ""
        tmpCASTreatReason5 = ""
        tmpCASDiagnosis5 = ""
            MCStat = 0
        tmpMCStatus = MCStat
        tmpNoOfDiag = "0"
        tmpNoOfMedication = "0"
        CISSql = "EXEC Save_Case_2013 '" & tmpCASnumber & "' ,'" & tmpMBMNumber & "', '" & tmpCoverID & "', " & _
             "'" & tmpCASStatus & "', '" & tmpClaimability & "', '" & tmpCaseType & "', '" & TmpPending & "', " & _
             "'" & tmpRegisterDate & "', '" & tmpREgisterBy & "','" & tmpLastUpdateDate & "', '" & tmpLastUpdateBy & "', " & _
             "'" & tmpClinicCode & "', '" & tmpInvoiceNo & "','" & tmpInvoiceDate & "', " & tmpInvoiceStmtAmt & ", " & _
             "'" & tmpDateVisit & "', " & tmpMCDays & ", '" & tmpOPBatchNo & "','" & tmpTimeVisit & "', '" & tmpMedInjType1 & "', " & _
             "'" & tmpMedInjType2 & "', '" & tmpMedInjType3 & "', '" & tmpMedInjType4 & "', '" & tmpMedInjType5 & "', " & _
             "'" & tmpMedInjDays1 & "' , '" & tmpMedInjDays2 & "', '" & tmpMedInjDays3 & "', '" & tmpMedInjDays4 & "'," & _
             "'" & tmpMedInjDays5 & "', '" & tmpMedInjQty1 & "', '" & tmpMedInjQty2 & "', '" & tmpMedInjQty3 & "'," & _
             "'" & tmpMedInjQty4 & "', '" & tmpMedInjQty5 & "', " & _
             "'" & tmpMedInjType6 & "', '" & tmpMedInjType7 & "', '" & tmpMedInjDays6 & "', " & _
             "'" & tmpMedInjDays7 & "', '" & tmpMedInjQty6 & "', '" & tmpMedInjQty7 & "','" & tmpInvoiceReceiptDate & "', " & _
             "'" & tmpInvoiceReceiptMode & "', '" & tmpCasRemark & "', '" & tmpCASTreatReason1 & "', '" & tmpCASTreatReason2 & "'," & _
             "'" & tmpCASTreatReason3 & "', '" & tmpCASTreatReason4 & "', '" & tmpCASTreatReason5 & "'," & _
             "'" & tmpCASDiagnosis5 & "', '" & tmpMCStatus & "', " & tmpNoOfDiag & "," & tmpNoOfMedication & ", '" & tmpINVMonth & "', '" & tmpINVYear & "'"
        
    MediConnCISDB.Execute CISSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub SaveInvoice()
Const METHOD_NAME = "SaveInvoice"
On Error GoTo debugs
Dim ClaimInvoice  As New ADODB.Recordset
Dim SaveInvoice As String
        SaveInvoice = ""
        tmpCASnumber = NewCaseNo
        tmpINVEnteredDate = Format(Date, "yyyy/mm/dd")
        tmpINVEnteredBy = Logon
        tmpINVLastUpdateBy = Logon
        tmpINVLastUpdateDate = Format(Date, "yyyy/mm/dd")
        tmpBillNo = ""
        tmpBillDate = ""
        tmpINVAmount = txtGLAmt
        tmpINVPayTo = "C"
        tmpINVPayee = ""
        tmpINVBillAmt = 0
        SaveInvoice = ""
    SaveInvoice = "EXEC Save_Invoice '" & tmpCASnumber & "', '" & tmpINVEnteredDate & "', '" & tmpINVEnteredBy & "','" & tmpINVLastUpdateBy & "', '" & _
                  tmpINVLastUpdateDate & "', '" & tmpBillNo & "', '" & tmpBillDate & "'," & CDbl(tmpINVAmount) & ",'" & _
                  tmpINVPayTo & "', '" & tmpINVPayee & "'," & CDbl(tmpINVBillAmt) & ""
    
    MediConnCISWS.Execute SaveInvoice
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub saveActual()
Const METHOD_NAME = "saveActual"
On Error GoTo debugs
    Dim SaveActualSql As String
    Dim abcefg As Double
        SaveActualSql = ""
        tmpCaseNo = Trim(NewCaseNo)
        tmpCLMGPFeeBasicMed = 0
        tmpCLMGPFeeLTMed = 0
        tmpCLMGPFeeMedInj = 0
        tmpCLMSPFeeBasicMed = 0
        tmpCLMSPFeeLTMed = 0
        tmpCLMSPFeeMedInj = 0
        tmpCLMDental = 0
        tmpCLMMaternity = 0
        tmpCLMPRVCheckup = 0
        tmpCLMXRay = 0
        tmpCLMUltraSound = 0
        tmpCLMECG = 0
        tmpCLMDressing = 0
        tmpCLMNebulizer = 0
        tmpCLMID = 0
        tmpCLMTSSTO = 0
        tmpCLMMinorSurgery = 0
        tmpCLMPOthers = 0
        tmpCLMBloodDrawing = 0
        tmpCLMBloodGlucose = 0
        tmpCLMUrine = 0
        tmpCLMPapSmear = 0
        tmpCLMLOthers = 0
        tmpCLMDTPPolio = 0
        tmpCLMDTPPolioHIB = 0
        tmpCLMHepatitisB = 0
        tmpCLMMMR = 0
        tmpCLMIOthers = 0
        tmpCLMPOthersDesc = ""
        tmpCLMLOthersDesc = ""
        tmpCLMIOthersDesc = ""
        tmpCLMVision = 0
        tmpCLMTyphiad = 0
        tmpCLMOtherNoBalDed = 0

        SaveActualSql = "EXEC Save_ClaimActual '" & tmpCaseNo & "', " & tmpCLMGPFeeBasicMed & ", " & tmpCLMGPFeeLTMed & ", " & _
                        "" & tmpCLMGPFeeMedInj & ", " & tmpCLMSPFeeBasicMed & ", " & tmpCLMSPFeeLTMed & ", " & tmpCLMSPFeeMedInj & ", " & _
                        "" & tmpCLMDental & ", " & tmpCLMMaternity & "," & tmpCLMPRVCheckup & ", " & tmpCLMXRay & ", " & _
                        "" & tmpCLMUltraSound & ", " & tmpCLMECG & ", " & tmpCLMDressing & ", " & tmpCLMNebulizer & ", " & _
                        "" & tmpCLMID & ", " & tmpCLMTSSTO & ", " & tmpCLMMinorSurgery & ", " & _
                        "" & tmpCLMPOthers & ", " & tmpCLMBloodDrawing & ", " & tmpCLMBloodGlucose & ", " & tmpCLMUrine & ", " & _
                        "" & tmpCLMPapSmear & ", " & tmpCLMLOthers & ", " & tmpCLMDTPPolio & ", " & tmpCLMDTPPolioHIB & ", " & _
                        "" & tmpCLMHepatitisB & ", " & tmpCLMMMR & ", " & tmpCLMIOthers & ", " & _
                        "'" & tmpCLMPOthersDesc & "', '" & tmpCLMLOthersDesc & "', '" & tmpCLMIOthersDesc & "', " & tmpCLMVision & ", " & _
                        "" & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""

    MediConnCISWS.Execute SaveActualSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
 
End Sub

Private Sub saveActualOthers()
    Dim StrSql As String
    StrSql = ""
    
    tmpCASnumber = NewCaseNo
    tmpGPFee = 0
    tmpGPLTMed = 0
    tmpGPInj = 0
    tmpGPMed = 0
    tmpSPFee = 0
    tmpSPLTMed = 0
    tmpSPInj = 0
    tmpSPMed = 0
    StrSql = ""
    StrSql = "Exec SaveActualOthers '" & tmpCASnumber & "'," & CDbl(tmpGPFee) & "," & CDbl(tmpGPLTMed) & "," & CDbl(tmpGPInj) & "," & _
             CDbl(tmpGPMed) & "," & CDbl(tmpSPFee) & "," & CDbl(tmpSPLTMed) & "," & CDbl(tmpSPInj) & "," & CDbl(tmpSPMed) & ""
     
    MediConnCISWS.Execute StrSql
End Sub

Private Sub SaveINVClaim()
Const METHOD_NAME = "SaveINVClaim"
On Error GoTo debugs
Dim SaveINVClaimSql As String
        
        tmpLimitExceed = 0
        SaveINVClaimSql = ""
        tmpCASnumber = NewCaseNo
        tmpClaimOpenDate = Format(Date, "yyyy/mm/dd")
        tmpClaimOpenBy = Logon
        tmpClaimCloseDate = Format(Date, "yyyy/mm/dd")
        tmpClaimClosedBy = Logon
        tmpClaimLastUpdateDate = Format(Date, "yyyy/mm/dd")
        tmpClaimLastUpdateUser = Logon
        tmpCLMCoPayColl = 0
        tmpCLMCLNExcess = 0
        tmpCLMTotalActual = txtGLAmt
        tmpCLMFinalApprove = txtGLAmt
        tmpINVBillAmt = txtGLAmt
        tmpCLMExcess = 0
        tmpCLMBalanceLimit = 0
        tmpCLMPreviousBalance = 0
        tmpCLMPaidBYMedix2M = 0
        tmpCLMPaidBYMedix2C = txtGLAmt
        tmpCLMAdjAnnLmt = 0
        tmpClmNonGPFee = 0
        tmpNonGPFee = 0
        tmpCLMMCOProcStatus = ""
        tmpCLMMCOProcFee = 0
        tmpCLMPayableByMedix2C = 0
        tmpCLMProcMth = ""
        tmpCLMProcYR = ""
        tmpCLMCoPayCLNColl = 0
        tmpCLMBordxAmt = txtGLAmt
        tmpLimitExceed = 0
        tmpCLMTotalApproved = txtGLAmt
        SaveINVClaimSql = ""
        SaveINVClaimSql = "Exec SaveINVClaim_2012 '" & tmpCASnumber & "', '" & tmpClaimOpenDate & "', '" & _
                            tmpClaimOpenBy & "', '" & tmpClaimCloseDate & "', '" & tmpClaimClosedBy & "', '" & _
                            tmpClaimLastUpdateDate & "', '" & tmpClaimLastUpdateUser & "','0'," & _
                            CDbl(tmpCLMPayableByMedix2C) & "," & tmpCLMCLNExcess & ", " & CDbl(tmpCLMTotalActual) & ", " & _
                            CDbl(tmpCLMFinalApprove) & "," & tmpCLMExcess & "," & tmpCLMBalanceLimit & "," & _
                            CDbl(tmpCLMPreviousBalance) & "," & tmpCLMPaidBYMedix2M & "," & tmpCLMPaidBYMedix2C & "," & _
                            tmpCLMAdjAnnLmt & "," & tmpClmNonGPFee & "," & tmpCLMMCOProcFee & ",'" & _
                            tmpCLMMCOProcStatus & "','" & tmpCLMProcMth & "', '" & tmpCLMProcYR & "'," & tmpCLMCoPayCLNColl & "," & _
                            CDbl(tmpCLMBordxAmt) & "," & tmpLimitExceed & "," & CDbl(tmpCLMTotalApproved) & ""
    
    MediConnCISWS.Execute SaveINVClaimSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveINVCLNApp()
Const METHOD_NAME = "SaveINVCLNApp"
On Error GoTo debugs
'Dim CISExc As New ADODB.Recordset
Dim SaveINVCLNAppSql As String
        
        SaveINVCLNAppSql = ""
        tmpCASnumber = NewCaseNo
        tmpCLMGPFeeBasicMed = 0
        tmpCLMGPFeeMedInj = 0
        tmpCLMGPFeeLTMed = 0
        tmpCLMSPFeeBasicMed = 0
        tmpCLMSPFeeLTMed = 0
        tmpCLMSPFeeMedInj = 0
        tmpCLMXRay = 0
        tmpCLMUltraSound = 0
        tmpCLMECG = 0
        tmpCLMDressing = 0
        tmpCLMNebulizer = 0
        tmpCLMID = 0
        tmpCLMTSSTO = 0
        tmpCLMMinorSurgery = 0
        tmpCLMPOthers = 0
        tmpCLMBloodDrawing = 0
        tmpCLMBloodGlucose = 0
        tmpCLMUrine = 0
        tmpCLMPapSmear = 0
        tmpCLMLOthers = 0
        tmpCLMDTPPolio = 0
        tmpCLMDTPPolioHIB = 0
        tmpCLMHepatitisB = 0
        tmpCLMMMR = 0
        tmpCLMIOthers = 0
        tmpCLMDental = 0
        tmpCLMMaternity = 0
        tmpCLMPRVCheckup = 0
        tmpCLMVision = 0
        tmpCLMTyphiad = 0
        tmpCLMOtherNoBalDed = 0

    SaveINVCLNAppSql = "Exec SaveINVCLNApp '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
                        tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & tmpCLMSPFeeLTMed & "," & _
                        tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & tmpCLMUltraSound & "," & tmpCLMECG & "," & _
                        tmpCLMDressing & "," & tmpCLMNebulizer & "," & tmpCLMID & "," & tmpCLMTSSTO & "," & _
                        tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & tmpCLMBloodGlucose & "," & _
                        tmpCLMUrine & "," & tmpCLMPapSmear & "," & tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & _
                        tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & _
                        tmpCLMDental & "," & tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & _
                        tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
                        
    MediConnCISWS.Execute SaveINVCLNAppSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveINVCLNContract()
Const METHOD_NAME = "SaveINVCLNContract"
On Error GoTo debugs
'Dim CISCLNLmt As New ADODB.Recordset
Dim SaveINVCLNContractSql As String
        
       SaveINVCLNContractSql = ""
       tmpCASnumber = NewCaseNo
       tmpCLMGPFeeBasicMed = 0
       tmpCLMGPFeeMedInj = 0
       tmpCLMGPFeeLTMed = 0
       tmpCLMSPFeeBasicMed = 0
       tmpCLMSPFeeLTMed = 0
       tmpCLMSPFeeMedInj = 0
       tmpCLMXRay = 0
       tmpCLMUltraSound = 0
       tmpCLMECG = 0
       tmpCLMDressing = 0
       tmpCLMNebulizer = 0
       tmpCLMID = 0
       tmpCLMTSSTO = 0
       tmpCLMMinorSurgery = 0
       tmpCLMPOthers = 0
       tmpCLMBloodDrawing = 0
       tmpCLMBloodGlucose = 0
       tmpCLMUrine = 0
       tmpCLMPapSmear = 0
       tmpCLMLOthers = 0
       tmpCLMDTPPolio = 0
       tmpCLMDTPPolioHIB = 0
       tmpCLMHepatitisB = 0
       tmpCLMMMR = 0
       tmpCLMIOthers = 0
       tmpCLMDental = 0
       tmpCLMMaternity = 0
       tmpCLMPRVCheckup = 0
       tmpCLMVision = 0
       tmpCLMTyphiad = 0
       tmpCLMOtherNoBalDed = 0
    
        SaveINVCLNContractSql = "Exec SaveINVCLNContract '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
        tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & tmpCLMSPFeeLTMed & "," & _
        tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & _
        tmpCLMNebulizer & "," & tmpCLMID & "," & tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & _
        tmpCLMBloodDrawing & "," & tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & _
        tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & _
        tmpCLMMMR & "," & tmpCLMIOthers & "," & tmpCLMDental & "," & tmpCLMMaternity & "," & _
        tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
    
    MediConnCISWS.Execute SaveINVCLNContractSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub SaveINVActual()
Const METHOD_NAME = "SaveINVActual"
On Error GoTo debugs
'Dim CISAct As New ADODB.Recordset
Dim SaveINVActualSql As String
    
        SaveINVActualSql = ""
        
        tmpCLMNo = NewCaseNo
        tmpCLMGPFeeBasicMed = 0
        tmpCLMGPFeeMedInj = 0
        tmpCLMGPFeeLTMed = 0
        tmpCLMSPFeeBasicMed = 0
        tmpCLMSPFeeLTMed = 0
        tmpCLMSPFeeMedInj = 0
        tmpCLMXRay = 0
        tmpCLMUltraSound = 0
        tmpCLMECG = 0
        tmpCLMDressing = 0
        tmpCLMNebulizer = 0
        tmpCLMID = 0
        tmpCLMTSSTO = 0
        tmpCLMMinorSurgery = 0
        tmpCLMPOthers = 0
        tmpCLMBloodDrawing = 0
        tmpCLMBloodGlucose = 0
        tmpCLMUrine = 0
        tmpCLMPapSmear = 0
        tmpCLMLOthers = 0
        tmpCLMDTPPolio = 0
        tmpCLMDTPPolioHIB = 0
        tmpCLMHepatitisB = 0
        tmpCLMMMR = 0
        tmpCLMIOthers = 0
        tmpCLMDental = 0
        tmpCLMMaternity = 0
        tmpCLMPRVCheckup = 0
        tmpCLMVision = 0
        tmpCLMTyphiad = 0
        tmpCLMOtherNoBalDed = 0
        tmpCLMTyphiad = 0
        tmpCLMOtherNoBalDed = 0
        SaveINVActualSql = "Exec SaveINVActual '" & tmpCLMNo & "'," & tmpCLMGPFeeBasicMed & "," & tmpCLMGPFeeMedInj & "," & _
                            tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & _
                            tmpCLMXRay & "," & tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & tmpCLMNebulizer & "," & _
                            tmpCLMID & "," & tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & _
                            tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & tmpCLMLOthers & "," & _
                            tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & _
                            tmpCLMDental & "," & tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & _
                            tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
         Debug.Print SaveINVActualSql
         
        MediConnCISWS.Execute SaveINVActualSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveINVCLNExcess()
Const METHOD_NAME = "SaveINVCLNExcess"
On Error GoTo debugs
Dim SaveINVCLNExcessSql As String
        
        SaveINVCLNExcessSql = ""
        
        tmpCASnumber = NewCaseNo
        tmpCLMGPFeeBasicMed = 0
        tmpCLMGPFeeMedInj = 0
        tmpCLMGPFeeLTMed = 0
        tmpCLMSPFeeBasicMed = 0
        tmpCLMSPFeeLTMed = 0
        tmpCLMSPFeeMedInj = 0
        tmpCLMXRay = 0
        tmpCLMUltraSound = 0
        tmpCLMECG = 0
        tmpCLMDressing = 0
        tmpCLMNebulizer = 0
        tmpCLMID = 0
        tmpCLMTSSTO = 0
        tmpCLMMinorSurgery = 0
        tmpCLMPOthers = 0
        tmpCLMBloodDrawing = 0
        tmpCLMBloodGlucose = 0
        tmpCLMUrine = 0
        tmpCLMPapSmear = 0
        tmpCLMLOthers = 0
        tmpCLMDTPPolio = 0
        tmpCLMDTPPolioHIB = 0
        tmpCLMHepatitisB = 0
        tmpCLMMMR = 0
        tmpCLMIOthers = 0
        tmpCLMDental = 0
        tmpCLMMaternity = 0
        tmpCLMPRVCheckup = 0
        tmpCLMVision = 0
        tmpCLMTyphiad = 0
        tmpCLMOtherNoBalDed = 0

    SaveINVCLNExcessSql = "Exec SaveINVCLNExcess '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
                        tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & tmpCLMSPFeeLTMed & "," & _
                        tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & _
                        tmpCLMNebulizer & "," & tmpCLMID & "," & tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & _
                        tmpCLMBloodDrawing & "," & tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & tmpCLMLOthers & "," & _
                        tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & _
                        tmpCLMDental & "," & tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & _
                        tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
    MediConnCISWS.Execute SaveINVCLNExcessSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveINVMBMApp()
Const METHOD_NAME = "SaveINVMBMApp"
On Error GoTo debugs
'Dim CISAPP As New ADODB.Recordset
Dim SaveINVMBMAppSql As String
    
        SaveINVMBMAppSql = ""
        tmpCASnumber = NewCaseNo
        tmpCLMGPFeeBasicMed = 0
        tmpCLMGPFeeMedInj = 0
        tmpCLMGPFeeLTMed = 0
        tmpCLMSPFeeBasicMed = 0
        tmpCLMSPFeeLTMed = 0
        tmpCLMSPFeeMedInj = 0
        tmpCLMXRay = 0
        tmpCLMUltraSound = 0
        tmpCLMECG = 0
        tmpCLMDressing = 0
        tmpCLMNebulizer = 0
        tmpCLMID = 0
        tmpCLMTSSTO = 0
        tmpCLMMinorSurgery = 0
        tmpCLMPOthers = 0
        tmpCLMBloodDrawing = 0
        tmpCLMBloodGlucose = 0
        tmpCLMUrine = 0
        tmpCLMPapSmear = 0
        tmpCLMLOthers = 0
        tmpCLMDTPPolio = 0
        tmpCLMDTPPolioHIB = 0
        tmpCLMHepatitisB = 0
        tmpCLMMMR = 0
        tmpCLMIOthers = 0
        tmpCLMDental = 0
        tmpCLMMaternity = 0
        tmpCLMPRVCheckup = 0
        tmpCLMVision = 0
        tmpCLMTyphiad = 0
        tmpCLMOtherNoBalDed = 0
       

        SaveINVMBMAppSql = "Exec SaveINVMBMApp '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
                            tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & _
                            tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & _
                            tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & tmpCLMNebulizer & "," & _
                            tmpCLMID & "," & tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & _
                            tmpCLMBloodDrawing & "," & tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & _
                            tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & _
                            tmpCLMMMR & "," & tmpCLMIOthers & "," & tmpCLMDental & "," & tmpCLMMaternity & "," & _
                            tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
        MediConnCISWS.Execute SaveINVMBMAppSql

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveINVMBMExcess()
Const METHOD_NAME = "SaveINVMBMExcess"
On Error GoTo debugs
Dim SaveINVMBMExcessSql As String

        SaveINVMBMExcessSql = ""
        tmpCASnumber = NewCaseNo
        tmpCLMGPFeeBasicMed = 0
        tmpCLMGPFeeMedInj = 0
        tmpCLMGPFeeLTMed = 0
        tmpCLMSPFeeBasicMed = 0
        tmpCLMSPFeeLTMed = 0
        tmpCLMSPFeeMedInj = 0
        tmpCLMXRay = 0
        tmpCLMUltraSound = 0
        tmpCLMECG = 0
        tmpCLMDressing = 0
        tmpCLMNebulizer = 0
        tmpCLMID = 0
        tmpCLMTSSTO = 0
        tmpCLMMinorSurgery = 0
        tmpCLMPOthers = 0
        tmpCLMBloodDrawing = 0
        tmpCLMBloodGlucose = 0
        tmpCLMUrine = 0
        tmpCLMPapSmear = 0
        tmpCLMLOthers = 0
        tmpCLMDTPPolio = 0
        tmpCLMDTPPolioHIB = 0
        tmpCLMHepatitisB = 0
        tmpCLMMMR = 0
        tmpCLMIOthers = 0
        tmpCLMDental = 0
        tmpCLMMaternity = 0
        tmpCLMPRVCheckup = 0
        tmpCLMVision = 0
        tmpCLMTyphiad = 0
        tmpCLMOtherNoBalDed = 0


        SaveINVMBMExcessSql = "Exec SaveINVMBMExcess '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
                                tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & _
                                tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & tmpCLMUltraSound & "," & _
                                tmpCLMECG & "," & tmpCLMDressing & "," & tmpCLMNebulizer & "," & tmpCLMID & "," & _
                                tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & _
                                tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & _
                                tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & _
                                tmpCLMDental & "," & tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & _
                                tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
    
    MediConnCISWS.Execute SaveINVMBMExcessSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveINVMBMLmt()
Const METHOD_NAME = "SaveINVMBMLmt"
On Error GoTo debugs
Dim SaveINVMBMLmtSql As String
        SaveINVMBMLmtSql = ""
        tmpCASnumber = NewCaseNo
            tmpCLMGPFeeMedInj = 0
            tmpCLMGPFeeLTMed = 0
            tmpCLMSPFeeBasicMed = 0
            tmpCLMSPFeeLTMed = 0
            tmpCLMSPFeeMedInj = 0
            tmpCLMXRay = 0
            tmpCLMUltraSound = 0
            tmpCLMECG = 0
            tmpCLMDressing = 0
            tmpCLMNebulizer = 0
            tmpCLMID = 0
            tmpCLMTSSTO = 0
            tmpCLMMinorSurgery = 0
            tmpCLMPOthers = 0
            tmpCLMBloodDrawing = 0
            tmpCLMBloodGlucose = 0
            tmpCLMUrine = 0
            tmpCLMPapSmear = 0
            tmpCLMLOthers = 0
            tmpCLMDTPPolio = 0
            tmpCLMDTPPolioHIB = 0
            tmpCLMHepatitisB = 0
            tmpCLMMMR = 0
            tmpCLMIOthers = 0
            tmpCLMDental = 0
            tmpCLMMaternity = 0
            tmpCLMPRVCheckup = 0
            tmpCLMVision = 0
            tmpCLMTyphiad = 0
            tmpCLMOtherNoBalDed = 0
      

        SaveINVMBMLmtSql = "Exec SaveINVMBMLmt '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
                            tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & _
                            tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & _
                            tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & _
                            tmpCLMNebulizer & "," & tmpCLMID & "," & tmpCLMTSSTO & "," & _
                            tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & _
                            tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & _
                            tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & _
                            tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & _
                            tmpCLMDental & "," & tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & _
                            tmpCLMVision & ", " & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
        
        MediConnCISWS.Execute SaveINVMBMLmtSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Function GenerateCISCaseNo(PayorCode)
Const METHOD_NAME = "GenerateCISCaseNo"
On Error GoTo debugs
    Dim CISCAQ As New ADODB.Recordset
    Dim StrSql As String
    Dim CISSeqNo As String
    Dim CAQCharacter As String
    Dim Prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "WebSearchCASNo_New2013Test"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 900
    'Set prm1 = cmd.CreateParameter("PayCode", adVarChar, adParamInput, 2000, "DC")
    Set Prm1 = cmd.CreateParameter("PayCode", adVarChar, adParamInput, 2000, "EA")
    cmd.Parameters.Append Prm1
    
    cmd.Parameters.Append cmd.CreateParameter("CaseNo1", adVarChar, adParamOutput, 2000)
    cmd.Execute
    CISSeqNo = Format(cmd("CaseNo1"), "00000000")
    'GenerateCISCaseNo = PayorCode & "K" & CISSeqNo
    GenerateCISCaseNo = "EA" & CISSeqNo
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Sub txtmaternitynormal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub txtCaesarean_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub txtComplications_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
