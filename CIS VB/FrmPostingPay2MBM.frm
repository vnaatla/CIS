VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPostingPay2MBM 
   Caption         =   "Payment To Member Posting"
   ClientHeight    =   8184
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11916
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8184
   ScaleWidth      =   11916
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   60
      TabIndex        =   34
      Top             =   260
      Width           =   11775
      Begin VB.ComboBox CboGrpCom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   195
         Width           =   4455
      End
      Begin VB.CheckBox ChkClinic 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11355
         TabIndex        =   11
         Top             =   525
         Width           =   375
      End
      Begin VB.CheckBox ChkGrp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11355
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox CboClinic 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   4455
      End
      Begin MSMask.MaskEdBox txtPVDate1 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   180
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPVDate2 
         Height          =   315
         Left            =   3600
         TabIndex        =   1
         Top             =   180
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboPostStatus 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   770
         Width           =   4455
      End
      Begin MSMask.MaskEdBox txtPostDate1 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   450
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtInvNo1 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TxtBatch1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtPostDate2 
         Height          =   315
         Left            =   3600
         TabIndex        =   3
         Top             =   450
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtInvNo2 
         Height          =   285
         Left            =   3600
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TxtBatch2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   46
         Top             =   750
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "PV No. from"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   5760
         TabIndex        =   44
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "PV Date from"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Post Status"
         Height          =   255
         Left            =   5760
         TabIndex        =   42
         Top             =   795
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Post Date from"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   495
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   40
         Top             =   195
         Width           =   345
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   39
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Clinic"
         Height          =   255
         Left            =   5760
         TabIndex        =   38
         Top             =   525
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   11400
         TabIndex        =   37
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   36
         Top             =   1005
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "Batch No. from"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame FramePostInfo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   25
      Top             =   6720
      Width           =   11775
      Begin VB.TextBox LblCrAmt 
         Height          =   285
         Left            =   10680
         TabIndex        =   54
         Top             =   200
         Width           =   975
      End
      Begin VB.TextBox LblDbAmt 
         Height          =   285
         Left            =   8760
         TabIndex        =   53
         Top             =   200
         Width           =   975
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   6960
         TabIndex        =   52
         Top             =   200
         Width           =   975
      End
      Begin VB.TextBox LblPost 
         Height          =   285
         Left            =   4920
         TabIndex        =   51
         Top             =   200
         Width           =   975
      End
      Begin VB.TextBox lblBatchNo 
         Height          =   285
         Left            =   2880
         TabIndex        =   50
         Top             =   200
         Width           =   975
      End
      Begin VB.ComboBox cboAccPeriod 
         Height          =   315
         Left            =   960
         TabIndex        =   13
         Text            =   "cboAccPeriod"
         Top             =   200
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Batch No"
         Height          =   255
         Left            =   2160
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Post Status"
         Height          =   255
         Left            =   3960
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "A/C Period"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "CR Amt"
         Height          =   255
         Left            =   9960
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "DN Amt"
         Height          =   255
         Left            =   8040
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   6120
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CheckBox CheckAll 
      Caption         =   "Check All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1630
      Width           =   975
   End
   Begin MSComctlLib.ListView lstMBMPayList 
      Height          =   4815
      Left            =   60
      TabIndex        =   47
      Top             =   1920
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   8488
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   21
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Chk"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "J/T"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "J/S"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "DB"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Trans.Ref."
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "AccPeriod"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Trans.Date"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "AccCode"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "D/C"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Clinics"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "PLN Number"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Payor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "YYMM"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Grp Company"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Bordx Ref"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "MemNo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "BatchNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "ClinicCode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "PostIndex"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lstCDList 
      Height          =   4695
      Left            =   60
      TabIndex        =   48
      Top             =   2040
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   8276
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   20
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Chk"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "J/T"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "J/S"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "DB"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Trans.Ref."
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "AccPeriod"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Trans.Date"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "AccCode"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "D/C"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Clinics"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Payor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "InsuredType"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Plan"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Grp Company"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Bordx Ref"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Mem No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "ClinicCode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "BatchNo"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   32
      Top             =   7440
      Width           =   11775
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7050
         TabIndex        =   14
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8640
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9240
         TabIndex        =   21
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10920
         TabIndex        =   23
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10200
         TabIndex        =   22
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7800
         TabIndex        =   19
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdPost 
         Caption         =   "&Post"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8520
         TabIndex        =   20
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdSearchCredit 
         Caption         =   "&SearchCR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7920
         TabIndex        =   16
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7200
         TabIndex        =   15
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton cmdSearchDebit 
         Caption         =   "&SearchDN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8760
         TabIndex        =   17
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton CmdChkAll 
         Caption         =   "Chk All"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9720
         TabIndex        =   18
         Top             =   200
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Payment To Member Posting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "FrmPostingPay2MBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intExit As Boolean
Dim PostingType As String
Dim PostingStatus As Boolean
Private m_blnDirection(1 To 22) As Boolean

Private Sub ComboListing()
Dim tmpSelField As String
Dim tmpTable As String
Dim tmpSelCriteria As String
        CboGrpCom.Clear
        tmpSelField = "SELECT AccMonth as tmpCode, AccYear as tmpName "
        tmpTable = "AccPeriod"
        tmpSelCriteria = "0=0"
        Call LoadComboListing(tmpSelField, tmpSelCriteria, tmpTable, cboAccPeriod, medixmasterconnMaint, "SearchMaintTableRec")
    cboPostStatus.AddItem "0 - No"
    cboPostStatus.AddItem "1 - Yes"
    cboPostStatus.ListIndex = 0
End Sub

Private Sub CboGrpCom_Change()
    If CboGrpCom.Text = "" Then
        If InStr(CboGrpCom.Text, "null") Then
           CboGrpCom.Enabled = False
        End If
    End If
End Sub

Private Sub cboPostStatus_LostFocus()
    TxtBatch1.Enabled = False
    TxtBatch2.Enabled = False
    txtPostDate1.Enabled = False
    txtPostDate2.Enabled = False
    If Mid(cboPostStatus, 1, 1) = "1" Then
        TxtBatch1.Enabled = True
        TxtBatch2.Enabled = True
        txtPostDate1.Enabled = True
        txtPostDate2.Enabled = True
    End If
End Sub

Private Sub CheckAll_Click()
Dim lstCnt As Integer
    lstCnt = lstMBMPayList.ListItems.Count
    If LblPost = "Yes" Then
        CheckAll.Value = 0
        Exit Sub
    Else
        If lstCnt = 0 Then
            CheckAll.Value = 0
        Else
            For lstCnt = 1 To lstCnt
                If lstMBMPayList.ListItems(lstCnt).Checked = False Then
                    lstMBMPayList.ListItems(lstCnt).Checked = True
                    lstMBMPayList.ListItems(lstCnt).Text = "Post"
                ElseIf lstMBMPayList.ListItems(lstCnt).Checked = True Then
                    lstMBMPayList.ListItems(lstCnt).Checked = False
                    lstMBMPayList.ListItems(lstCnt).Text = ""
                End If
            Next lstCnt
        End If
    End If
End Sub

Private Sub cmdBack_Click()
    lstCDList.Visible = False
    lstMBMPayList.Visible = True
    CmdPost.Enabled = True
End Sub

Private Sub cmdchkAll_Click()
    Call SetCheckAllItems(True)
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    Call ClearInfo
    cboPostStatus.ListIndex = 0
    FramePostInfo.Enabled = False
    CmdChkAll.Enabled = False
    TxtBatch1.Enabled = False
    TxtBatch2.Enabled = False
    txtPostDate1.Enabled = False
    txtPostDate2.Enabled = False
    lstMBMPayList.ListItems.Clear
    lstCDList.ListItems.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPost_Click()
Const METHOD_NAME = "CmdPost_Click"
On Error GoTo debugs

Dim AA As Long
Dim StrInvSql As String
Dim tmpPvNo As String
Dim tmpInvNo As String
Dim UpdAsk
Dim BatchNo As String
Dim CompTrans As Boolean
Dim ItemChecked As Boolean
Dim UpdateSuc As Boolean
    
    ItemChecked = False
    CompTrans = True
    UpdateSuc = False
    For AA = 1 To lstMBMPayList.ListItems.Count
        If lstMBMPayList.ListItems(AA).Checked = True Then
            ItemChecked = True
            Exit For
        End If
    Next AA
    If ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
        Exit Sub
    End If
    If LblPost = "Yes" Then
        MsgBox "Record(s) has been posted ! ", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    
    If txtPVDate1 = "__/__/____" Or txtPVDate2 = "__/__/____" Then
        MsgBox "Please Enter PV Date! ", vbOKOnly + vbCritical, DialogBoxTitle
        txtPVDate1.SetFocus
        Exit Sub
    End If

    If Trim(cboAccPeriod.Text) = "" Then
        MsgBox "Please Keyin for Accounting Period ! ", vbOKOnly + vbCritical, DialogBoxTitle
        cboAccPeriod.SetFocus
        Exit Sub
    End If
    UpdAsk = MsgBox("Do you want to Post record(s) ?", vbYesNo + vbExclamation)
    If UpdAsk = vbYes Then
        Screen.MousePointer = vbHourglass
        On Err GoTo ErrorSave
           ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
                MediConnCISWS.beginTrans
                CompTrans = False
                BatchNo = GenSUNPostingBatchNo
                For AA = 1 To lstMBMPayList.ListItems.Count
                    tmpPvNo = Trim(lstMBMPayList.ListItems(AA).SubItems(4))
                    tmpInvNo = ""
                    If lstMBMPayList.ListItems(AA).Checked = True And Trim(lstMBMPayList.ListItems(AA).SubItems(20)) = "Post" Then
                        StrInvSql = ""
                        StrInvSql = "Insert Into SUNPosting(InvoiceNo, TransRef, Pay2MBMStatus," & _
                                    "PostStatus, PostType, PostDate, PostBy, " & _
                                    "AccMonth, AccYear, BatchNo) " & _
                                    "Values('" & Mid(tmpInvNo, 1, 20) & "','" & tmpPvNo & "','1'" & _
                                    ",'1','" & PostingType & "','" & Format(Date, "yyyy/mm/dd") & _
                                    "','" & Logon & "','" & Mid(cboAccPeriod, 1, 2) & _
                                    "','" & Mid(cboAccPeriod, 4, 4) & "','" & BatchNo & "')"
                        MediConnCISWS.Execute StrInvSql
                    End If
                Next AA
                MediConnCISWS.commitTrans
                CompTrans = True
                UpdateSuc = True
                lblBatchNo = BatchNo
                StrInvSql = ""
                StrInvSql = " and BatchNo = '" & Trim(lblBatchNo) & "'"
                Call MBMPayListing(StrInvSql, "G")
                PostingStatus = True
                Call CmdPrintFile_Click
                cboPostStatus.ListIndex = 1
            'MediConnCISWS.Close
            'Set MediConnCISWS = Nothing
            Screen.MousePointer = vbDefault
            MsgBox "Record(s) Posted Sucessfully... ! ", vbOKOnly
        End If
ErrorSave:
        Screen.MousePointer = vbDefault
        If CompTrans = False Then
            MediConnCISWS.RollbackTrans
            
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
        End If
        
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

        
End Sub

Private Sub CmdPrintFile_Click()
Dim SVRPath As String
Dim Filename As String
    
    Screen.MousePointer = vbHourglass
        
    If lstMBMPayList.ListItems.Count <> 0 Or lstCDList.ListItems.Count <> 0 Then
        If PostingStatus = False Then
            If lstMBMPayList.Visible = True Then
                Call ExportListViewtoExcel(lstMBMPayList)
            Else
                Call ExportListViewtoExcel(lstCDList)
            End If
        Else
            '\\Mars\accounts\Account Journal Import\JI Not Posted
            SVRPath = "\\Mars\accounts\Account Journal Import\JIS CLN Not Posted\"
            Filename = lblBatchNo & " CP-PayMBM " & Format(txtPVDate1, "yyyymmdd") & "-" & Mid(txtPVDate2, 1, 2)
            Call ExporttoExcelPosting(lstMBMPayList, SVRPath, Filename)
            PostingStatus = False
        End If
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSearch_Click()
    lstMBMPayList.ListItems.Clear
    lstMBMPayList.Visible = True
    lstCDList.Visible = False
    Call SearchRecord("G")
End Sub

Private Sub cmdSearchCredit_Click()
    lstMBMPayList.Visible = False
    lstCDList.Visible = True
    Call SearchRecord("C")
End Sub

Private Sub cmdSearchDebit_Click()
    lstMBMPayList.Visible = False
    lstCDList.Visible = True
    Call SearchRecord("D")
End Sub

Private Sub CmdStop_Click()
    intExit = True
    CmdExit.Enabled = True
End Sub
Private Sub Form_Load()
Call ActionLogUpdate(12, "FrmPostingPay2MBM")
    intExit = False
    PostingType = "Pay2MBM"
    Call ComboListing
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
Private Sub lstMBMPayList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Const METHOD_NAME = "lstMBMPayList_ItemCheck"
On Error GoTo debugs

Dim CurrChk As Integer
Dim ModResult As Integer
Dim NextChk As Integer
Dim NextChk2 As Integer
Dim lstRec As ListItem
Dim PrevText As String
Dim AA As Integer
Dim AB As Integer
    NextChk2 = 0
    AA = 0
    AB = 0
    If LblPost = "Yes" Then
        MsgBox "Record(s) has been posted ! ", vbOKOnly + vbCritical, DialogBoxTitle
        Item.Checked = False
    Else
        
        AA = Item.Index + 1
        
        If Item.Text = "" Then
            If Item.ListSubItems(10) = "C" Then
                For AA = AA To lstMBMPayList.ListItems.Count
                    If lstMBMPayList.ListItems(AA).SubItems(10) <> "C" Then
                        Item.Text = "Post"
                        lstMBMPayList.ListItems(AA).Text = "Post"
                        Item.Checked = True
                        lstMBMPayList.ListItems(AA).Checked = True
                    Else
                        Exit Sub
                    End If
                Next AA
            Else
                Item.Text = ""
                lstMBMPayList.ListItems(AA).Text = ""
                Item.Checked = False
                lstMBMPayList.ListItems(AA).Checked = False
                MsgBox "Please select Credit Item! ", vbOKOnly + vbCritical, DialogBoxTitle
            End If
        End If
        
        If Item.Text = "Post" Then
            If Item.ListSubItems(10) = "C" Then
                For AA = AA To lstMBMPayList.ListItems.Count
                    If lstMBMPayList.ListItems(AA).SubItems(10) <> "C" Then
                        Item.Text = ""
                        lstMBMPayList.ListItems(AA).Text = ""
                        Item.Checked = False
                        lstMBMPayList.ListItems(AA).Checked = False
                    Else
                        Exit Sub
                    End If
                Next AA
            Else
                Item.Checked = True
                lstMBMPayList.ListItems(AA - 1).Checked = True
                MsgBox "Please select Credit Item! ", vbOKOnly + vbCritical, DialogBoxTitle
            End If
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
        
End Sub

Private Sub txtPostDate1_LostFocus()
    If IsDate(txtPostDate1) Then
    Else
        If txtPostDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPostDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtPostDate2_LostFocus()
    If IsDate(txtPostDate2) Then
    Else
        If txtPostDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPostDate2.SetFocus
        End If
    End If
End Sub

Private Sub MBMPayListing(strAppend As String, tmpCDType As String)
Const METHOD_NAME = "MBMPayListing"
On Error GoTo debugs

Dim MBMPay As New ADODB.Recordset
Dim JournalType, AcCode, CreditType As String
Dim AA, CurrRec As Integer
Dim TotDbAmt, TotCrAmt As Double
Dim cmd As New ADODB.Command
Dim App1 As New ADODB.Parameter
Dim App2 As New ADODB.Parameter
Dim RecLoop As Integer
Dim tmpDesc As String
Dim lstRec As ListItem
Dim tmpMCOProcFee As Double
Dim CrCnt As Integer
Dim tmpDispAmt As Double
Dim tmpColPayAmt As Double
Dim strsql As String
Dim PayMBM As New ADODB.Recordset
    CurrRec = 0
    LblPost = "No"
    If tmpCDType = "G" Then
         lstMBMPayList.ListItems.Clear
    Else
         lstCDList.ListItems.Clear
    End If
    'VIEW_SUN_Posting_Pay2MBM
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "SUN_CIS_Posting"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 2400
    Set App1 = cmd.CreateParameter("RptType", adVarChar, adParamInput, 20, PostingType)
    cmd.Parameters.Append App1
    Set App2 = cmd.CreateParameter("StrCriterial", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append App2
    MBMPay.CursorLocation = adUseClient
    MBMPay.CursorType = adOpenForwardOnly
    MBMPay.CacheSize = 100
    Set MBMPay = cmd.Execute
    Do While Not MBMPay.EOF
       With MBMPay
            RecLoop = 1
            If tmpCDType = "G" Then RecLoop = 2
            CurrRec = CurrRec + 1
            JournalType = "CP"
            
            If MBMPay!Pay2MBMStatus = "1" Then
                LblPost = "Yes"
            End If
            AcCode = "7610"
            CreditType = "C"
            tmpDesc = Trim(!PVPayto)
            tmpDispAmt = IIf(IsNumeric(!MBMPaidAmt), !MBMPaidAmt, 0)
            If tmpCDType = "D" Then
                AcCode = "82" & !GRPCoCode
                CreditType = "D"
                tmpDesc = !PVPayto
                tmpDispAmt = IIf(IsNumeric(!MBMPaidAmt), !MBMPaidAmt, 0)
            End If
            CrCnt = 0
            For AA = 1 To RecLoop
                If AA = 2 Then
                    AcCode = "82" & !GRPCoCode
                    CreditType = "D"
                    tmpDesc = Trim(!PVPayto)
                    
                End If
                                
                If AA = 2 Then
                            Set lstRec = lstMBMPayList.ListItems.Add(, , IIf(MBMPay!Pay2MBMStatus = 1, "Post", ""))
                            
                            lstRec.SubItems(1) = JournalType
                            lstRec.SubItems(3) = "UHC"
                            lstRec.SubItems(4) = IIf(Len(Trim(!PVNo)), Mid(!PVNo, 2, 10), "")
                            If Trim(!AccMonth & !AccYear) <> "" Then
                                lstRec.SubItems(5) = !AccMonth & "/" & !AccYear
                            End If
                            lstRec.SubItems(6) = IIf(IsDate(!PVDate), Format(!PVDate, "DD/MM/YYYY"), "")
                            lstRec.SubItems(7) = Mid(Trim(tmpDesc), 1, 25)
                            lstRec.SubItems(8) = IIf(Len(AcCode), AcCode, "")
                            lstRec.SubItems(9) = Format(tmpDispAmt, "#,##0.00")
                            lstRec.SubItems(10) = Trim(CreditType)
                            lstRec.SubItems(11) = "1R"
                            
                            'lstRec.SubItems(12) = PLN
                            lstRec.SubItems(13) = IIf(Len(!PAYCode), Trim(!PAYCode), "")
                            'lstRec.SubItems(14) = YYMM
                            lstRec.SubItems(15) = IIf(Len(!GRPCoCode), Trim(!GRPCoCode), "")
                            'lstRec.SubItems(16) BordxRef
                            'lstRec.SubItems(17) MBMNo
                            lstRec.SubItems(18) = IIf(Len(Trim(!BatchNo)), Trim(!BatchNo), "")
                            'lstRec.SubItems(19) = IIf(Len(Trim(!ClinicCode)), Trim(!ClinicCode), "")
                            TotDbAmt = TotDbAmt + tmpDispAmt

                Else
                    Set lstRec = lstMBMPayList.ListItems.Add(, , IIf(MBMPay!Pay2MBMStatus = 1, "Post", ""))
                    lstRec.SubItems(1) = JournalType
                    'lstRec.SubItems(2) J/S
                    lstRec.SubItems(3) = "UHC"
                    lstRec.SubItems(4) = IIf(Len(Trim(!PVNo)), Mid(!PVNo, 2, 10), "")
                    If Trim(!AccMonth & !AccYear) <> "" Then
                        lstRec.SubItems(5) = !AccMonth & "/" & !AccYear
                    End If
                    lstRec.SubItems(6) = IIf(IsDate(!PVDate), Format(!PVDate, "DD/MM/YYYY"), "")
                    lstRec.SubItems(7) = Mid(Trim(tmpDesc), 1, 25)
                    lstRec.SubItems(8) = IIf(Len(AcCode), AcCode, "")
                    lstRec.SubItems(9) = Format(tmpDispAmt, "#,##0.00")
                    lstRec.SubItems(10) = Trim(CreditType)
                    lstRec.SubItems(11) = "1R"
                    
                    'lstRec.SubItems(12) = PLN
                    lstRec.SubItems(13) = IIf(Len(!PAYCode), Trim(!PAYCode), "")
                    'lstRec.SubItems(14) = YYMM
                    lstRec.SubItems(15) = IIf(Len(!GRPCoCode), Trim(!GRPCoCode), "")
                    'lstRec.SubItems(16) BordxRef
                    'lstRec.SubItems(17) MBMNo
                    lstRec.SubItems(18) = IIf(Len(Trim(!BatchNo)), Trim(!BatchNo), "")
                    'lstRec.SubItems(19) = IIf(Len(Trim(!ClinicCode)), Trim(!ClinicCode), "")
                    lstRec.SubItems(20) = "Post"
                    TotCrAmt = TotCrAmt + tmpDispAmt
                    'TotDbAmt = TotDbAmt + tmpDispAmt
            End If
                'If CreditType = "C" Then
                    
                'Else
                '    TotDbAmt = TotDbAmt + tmpDispAmt
                'End If
            Next AA
            LblDbAmt = Format(TotDbAmt, "#,##0.00")
            LblCrAmt = Format(TotCrAmt, "#,##0.00")
            lblCurrent = CurrRec
        End With

        MBMPay.MoveNext
        DoEvents
        If intExit = True Then
           Exit Do
        End If
    Loop
    intExit = 0
    
    MBMPay.Close
    Set MBMPay = Nothing
    If CurrRec = 0 Then
        MsgBox "No record selected ! ", vbOKOnly + vbCritical
    ElseIf LblPost = "No" Then
        FramePostInfo.Enabled = True
        CmdChkAll.Enabled = True
    End If
    cmdSearchCredit.Enabled = True
    cmdSearchDebit.Enabled = True
    CmdPost.Enabled = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub


Private Sub ClearInfo()
    lblCurrent = ""
    LblDbAmt = ""
    LblCrAmt = ""
    lblBatchNo = ""
    'cboAccPeriod.Text = ""
    FramePostInfo.Enabled = False
    CmdChkAll.Enabled = False
    CheckAll.Value = 0
    cmdSearchCredit.Enabled = False
    cmdSearchDebit.Enabled = False
    CmdPost.Enabled = False
End Sub

Private Sub SetCheckAllItems(bState As Boolean)
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim r As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstMBMPayList.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstMBMPayList.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub

Private Sub ChkClinic_Click()
Dim tmpSelField As String
Dim tmpTable As String
Dim tmpSelCriteria As String
    CboClinic.Clear
    If ChkClinic.Value = 1 Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            tmpSelField = "SELECT MBMCode as tmpCode, MBMName as tmpName "
            tmpTable = "MBM"
            tmpSelCriteria = "0=0"
            Call LoadComboListing(tmpSelField, tmpSelCriteria, tmpTable, CboClinic, medixmasterconn, "SearchHISTableRec")
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
    End If
End Sub

Private Sub ChkGrp_Click()
Dim tmpSelField As String
Dim tmpTable As String
Dim tmpSelCriteria As String

    CboGrpCom.Clear
    If ChkGrp.Value = 1 Then
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        
            tmpSelField = "SELECT GRPCoCode as tmpCode, GRPCoName as tmpName "
            tmpTable = "GRPCompany"
            tmpSelCriteria = ""
            Call LoadComboListing(tmpSelField, tmpSelCriteria, tmpTable, CboGrpCom, MediConnCISDB, "SearchCISDBTableRec")
        
      '  MediConnCISDB.Close
      '  Set MediConnCISDB = Nothing
    End If
End Sub

Private Sub LoadComboListing(tmpSelField As String, tmpSelCriteria As String, tmpTableName As String, tmpComboBox As ComboBox, tmpConn As ADODB.Connection, tmpStoredProc As String)
Const METHOD_NAME = "LoadComboListing"
On Error GoTo debugs


Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter

    tmpComboBox.Clear
    Set cmd.ActiveConnection = tmpConn
    cmd.CommandText = tmpStoredProc
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelField)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 255, tmpSelCriteria)
    cmd.Parameters.Append Prm3
    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    
    Do Until TmpRec.EOF
        With TmpRec
        If tmpTableName = "GRPCompany" Then
            tmpComboBox.AddItem !tmpName
        ElseIf tmpTableName = "MBM" Then
            tmpComboBox.AddItem !TmpCode & " - " & !tmpName
        ElseIf tmpTableName = "AccPeriod" Then
            tmpComboBox.AddItem !TmpCode & "/" & !tmpName
        End If
        TmpRec.MoveNext
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Function SearchRecord(tmpStatus As String)
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs


Dim strsql As String
Dim PostSql As String
Dim AA As Integer
Dim tmpMBM As String

    Call ClearInfo
    strsql = ""
    PostSql = ""
    If Trim(CboGrpCom) <> "" Then
        strsql = strsql & " And GRPCoName='" & Trim(CboGrpCom) & "'"
    End If
    If IsDate(txtPVDate1) Then
        strsql = strsql & " And PVDate >='" & Format(txtPVDate1, "yyyy/mm/dd") & "'"
    End If
    If IsDate(txtPVDate2) Then
        strsql = strsql & " And PVDate <='" & Format(txtPVDate2, "yyyy/mm/dd") & "'"
    End If
    If Trim(txtInvNo1) <> "" Then
        strsql = strsql & " And PVNo >='" & Trim(txtInvNo1) & "'"
    End If
    If Trim(txtInvNo2) <> "" Then
        strsql = strsql & " And PVNo <='" & Trim(txtInvNo2) & "'"
    End If
    If Mid(cboPostStatus, 1, 1) = "0" Then
        PostSql = " And (Pay2MBMStatus IS NULL OR Pay2MBMStatus = '0')"
        PostSql = PostSql & " And (PostStatus IS NULL OR PostStatus = '0')"
    ElseIf Mid(cboPostStatus, 1, 1) = "1" Then
        PostSql = " And Pay2MBMStatus ='1' And PostStatus ='1'"
    End If
    If IsDate(txtPostDate1) Then
        strsql = strsql & " And PostDate >='" & Format(txtPostDate1, "yyyy/mm/dd") & "'"
    End If
    If IsDate(txtPostDate2) Then
        strsql = strsql & " And PostDate <='" & Format(txtPostDate2, "yyyy/mm/dd") & "'"
    End If
    If Trim(TxtBatch1) <> "" Then
        strsql = strsql & " And BatchNo >='" & Trim(TxtBatch1) & "'"
    End If
    If Trim(TxtBatch2) <> "" Then
        strsql = strsql & " And BatchNo <='" & Trim(TxtBatch2) & "'"
    End If
    
    Screen.MousePointer = vbHourglass
        Call MBMPayListing(strsql & PostSql, tmpStatus)
    Screen.MousePointer = vbDefault
    

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
        
End Function




