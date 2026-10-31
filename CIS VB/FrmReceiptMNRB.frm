VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReceiptMNRB 
   Caption         =   "Receipt From MNRB"
   ClientHeight    =   8172
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11916
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8172
   ScaleWidth      =   11916
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   51
      Top             =   7080
      Width           =   11655
      Begin VB.TextBox lblSelectedAmt 
         Height          =   285
         Left            =   4440
         TabIndex        =   70
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox lbltotal 
         Height          =   285
         Left            =   1560
         TabIndex        =   69
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox lblcurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   68
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   330
         Left            =   9720
         TabIndex        =   61
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "S&earch"
         Height          =   330
         Left            =   6000
         TabIndex        =   60
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   6720
         TabIndex        =   59
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   58
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10320
         TabIndex        =   57
         Top             =   220
         Width           =   615
      End
      Begin VB.TextBox TxtPStatus 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         MaxLength       =   15
         TabIndex        =   56
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtPStatus2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   55
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   7920
         TabIndex        =   54
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Height          =   330
         Left            =   7320
         TabIndex        =   53
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton cmdEXPtoExcel 
         Caption         =   "&Print to File"
         Height          =   330
         Left            =   8640
         TabIndex        =   52
         Top             =   220
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   1200
         TabIndex        =   64
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   1680
         TabIndex        =   63
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Checked Amt"
         Height          =   255
         Left            =   3360
         TabIndex        =   62
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   885
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   11655
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7800
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   3705
      End
      Begin VB.TextBox txtReceiptAmt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cboAccPeriodMth 
         Height          =   315
         ItemData        =   "FrmReceiptMNRB.frx":0000
         Left            =   6000
         List            =   "FrmReceiptMNRB.frx":0028
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtReceiptNo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtReceiptDate 
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   480
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   445
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboAccPeriodYR 
         Height          =   315
         ItemData        =   "FrmReceiptMNRB.frx":0053
         Left            =   6000
         List            =   "FrmReceiptMNRB.frx":00DB
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtChequeNo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Cheque No"
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   550
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Rec Amt"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Remark"
         Height          =   255
         Left            =   7200
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Rec No"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   550
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "A/C Period Mth"
         Height          =   255
         Left            =   4800
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "A/C Period Yr"
         Height          =   255
         Left            =   4800
         TabIndex        =   10
         Top             =   550
         Width           =   1335
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Receipt from MNRB"
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
         TabIndex        =   1
         Top             =   50
         Width           =   3285
      End
   End
   Begin MSComctlLib.ListView lstBordx 
      Height          =   4125
      Left            =   120
      TabIndex        =   65
      Top             =   2160
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   7281
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Payor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Bordx Amt"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Receipt Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "OS Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Receipt Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstWorksheet 
      Height          =   4080
      Left            =   120
      TabIndex        =   66
      Top             =   2160
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   7197
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Payor"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Worksheet No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Receipt No"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Receipt Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Rec. Status"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1740
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   11655
      Begin VB.OptionButton OptWks 
         Caption         =   "By Worksheet"
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
         Left            =   3000
         TabIndex        =   26
         Top             =   200
         Width           =   1815
      End
      Begin VB.OptionButton OptBordx 
         Caption         =   "By Bordx"
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
         Left            =   1800
         TabIndex        =   28
         Top             =   200
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtRECNoFr 
         Height          =   285
         Left            =   8160
         MaxLength       =   15
         TabIndex        =   22
         Top             =   510
         Width           =   1215
      End
      Begin VB.TextBox txtRECNoTo 
         Height          =   285
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   21
         Top             =   510
         Width           =   1575
      End
      Begin VB.TextBox txtWSNoFr 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   20
         Top             =   580
         Width           =   1575
      End
      Begin VB.TextBox txtWSNoTo 
         Height          =   285
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   5400
         TabIndex        =   18
         Top             =   840
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtRECDateFr 
         Height          =   315
         Left            =   8160
         TabIndex        =   29
         Top             =   750
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRECDateTo 
         Height          =   315
         Left            =   9840
         TabIndex        =   30
         Top             =   720
         Width           =   1575
         _ExtentX        =   2773
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   23
         Top             =   820
         Width           =   1575
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   24
         Top             =   820
         Width           =   1575
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1080
         Width           =   3615
      End
      Begin VB.ComboBox CboWSStatus 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1320
         Width           =   3615
      End
      Begin VB.ComboBox CboBordxStatus 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1320
         Width           =   3615
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         Height          =   315
         ItemData        =   "FrmReceiptMNRB.frx":01E7
         Left            =   8160
         List            =   "FrmReceiptMNRB.frx":020F
         TabIndex        =   34
         Top             =   1035
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         Height          =   315
         ItemData        =   "FrmReceiptMNRB.frx":023A
         Left            =   9840
         List            =   "FrmReceiptMNRB.frx":0262
         TabIndex        =   32
         Top             =   1005
         Width           =   1575
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         Height          =   315
         ItemData        =   "FrmReceiptMNRB.frx":028D
         Left            =   8160
         List            =   "FrmReceiptMNRB.frx":0315
         TabIndex        =   35
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         Height          =   315
         ItemData        =   "FrmReceiptMNRB.frx":0421
         Left            =   9840
         List            =   "FrmReceiptMNRB.frx":04A9
         TabIndex        =   33
         Top             =   1275
         Width           =   1575
      End
      Begin VB.Label Label22 
         Caption         =   "A/C Period Mth Fr"
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
         Left            =   6600
         TabIndex        =   50
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label23 
         Caption         =   "To"
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
         Left            =   9480
         TabIndex        =   49
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Receipt Type"
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
         TabIndex        =   48
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
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
         TabIndex        =   47
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx No from"
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
         TabIndex        =   46
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "To"
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
         Left            =   3360
         TabIndex        =   45
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Payment Status"
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
         TabIndex        =   44
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "To"
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
         Left            =   9480
         TabIndex        =   43
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Recp No from"
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
         Left            =   6600
         TabIndex        =   42
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Recp Date from"
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
         Left            =   6600
         TabIndex        =   41
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "To"
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
         Left            =   9480
         TabIndex        =   40
         Top             =   765
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "WS No From"
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
         TabIndex        =   39
         Top             =   580
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "To"
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
         Left            =   3360
         TabIndex        =   38
         Top             =   585
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "To"
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
         Left            =   9480
         TabIndex        =   37
         Top             =   1365
         Width           =   375
      End
      Begin VB.Label Label25 
         Caption         =   "A/C Period Yr Fr"
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
         Left            =   6600
         TabIndex        =   36
         Top             =   1365
         Width           =   1815
      End
   End
   Begin VB.Label Label33 
      Caption         =   "Bold - Searchable"
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
      TabIndex        =   67
      Top             =   7770
      Width           =   1695
   End
End
Attribute VB_Name = "FrmReceiptMNRB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim BordxNoArray() As String
Dim one As Boolean
Dim two As Boolean
Dim Check1 As Boolean
Dim Check2 As Boolean
Dim ScanCode As String
Dim listloop1 As Integer
Private m_blnDirection(1 To 9) As Boolean

Private Sub ChkScan_Click()
    If OptBordx.Value = True And ChkScan.Value = 0 Then
        Label14.Visible = True
        TxtBordx2.Visible = True
        ChkScan.Left = 5400
        ChkScan.Top = 840
        TxtBordx1.SetFocus
    ElseIf OptBordx.Value = True And ChkScan.Value = 1 Then
        Label14.Visible = False
        TxtBordx2.Visible = False
        ChkScan.Left = 3360
        ChkScan.Top = 840
        TxtBordx1.SetFocus
    ElseIf OptWks.Value = True And ChkScan.Value = 0 Then
        Label21.Visible = True
        txtWSNoTo.Visible = True
        txtWSNoFr.SetFocus
        ChkScan.Left = 5400
        ChkScan.Top = 600
    ElseIf OptWks.Value = True And ChkScan.Value = 1 Then
        Label21.Visible = False
        txtWSNoTo.Visible = False
        ChkScan.Left = 3360
        ChkScan.Top = 600
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
Private Sub cmdEdit_Click()
    Screen.MousePointer = vbHourglass
    If OptBordx.Value = True Then
        one = True
    Else
        two = True
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdUpdate_Click()
    Screen.MousePointer = vbHourglass
    If OptBordx.Value = True Then
        Call UpdateBordx
    Else
        Call UpdateWorksheet
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Call ActionLogUpdate(12, "FrmReceiptMNRB")
    one = False
    two = False
    Call ComboListing
    Call PayorComboListing
    ReDim BordxNoArray(0)
    ReDim BordxNoArray(0)
    lstWorksheet.Visible = False
    txtWSNoFr.Enabled = False
    txtWSNoTo.Enabled = False
    '==============================
    Label14.Visible = False
    TxtBordx2.Visible = False
    ChkScan.Left = 3360
    ChkScan.Top = 840
    '==============================
    CboBordxStatus.ListIndex = 0
    CboWSStatus.Visible = False
    intExit = 0
    lblcurrent = 0
    lbltotal = 0
    lblSelectedAmt = 0
End Sub

'************** Start Command Button ****************************

Private Sub cmdSearch_Click()
Const METHOD_NAME = "cmdSearch_Click"
On Error GoTo debugs

If Mid(TxtBordx1, 1, 2) = "LO" Or Mid(txtWSNoFr, 1, 2) = "LO" Then
    If Mid(TxtBordx1, 1, 2) = "LO" Then
        ScanCode = TxtBordx1
        TxtBordx1 = ""
    ElseIf Mid(txtWSNoFr, 1, 2) = "LO" Then
        ScanCode = txtWSNoFr
        txtWSNoFr = ""
    End If
    'Call SearchForm(ScanCode)
Else
    one = False
    two = False
    
    CmdEdit.Enabled = False
    CmdUpdate.Enabled = False
    cmdEXPtoExcel.Enabled = False
    CmdSave.Enabled = False
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    ReDim BordxNoArray(0)
        
    Screen.MousePointer = vbHourglass
    
        If ChkScan.Value = 1 And OptWks.Value = True Then
            If txtWSNoFr = "" Then
                MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
                txtWSNoFr.SetFocus
                CmdExit.Enabled = True
                CmdEdit.Enabled = True
                CmdUpdate.Enabled = True
                cmdEXPtoExcel.Enabled = True
                CmdSave.Enabled = True
                CmdClear.Enabled = True
            Else
            
                Screen.MousePointer = vbHourglass
                
                listrecord = lstWorksheet.ListItems.Count
                If listrecord <> 0 Then
                    
                    For listloop1 = 1 To listrecord
                        Call CaseNoCheck
                    Next listloop1
                End If
                If Check1 = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    CmdExit.Enabled = True
                    CmdEdit.Enabled = True
                    CmdUpdate.Enabled = True
                    cmdEXPtoExcel.Enabled = True
                    CmdSave.Enabled = True
                    CmdClear.Enabled = True
                ElseIf Check1 = False Then
                    Call SearchRecord
                End If
                Check1 = False
                Screen.MousePointer = Default
                txtWSNoFr = ""
                txtWSNoFr.SetFocus
            End If
            txtWSNoFr.SetFocus
        ElseIf ChkScan.Value = 1 And OptBordx.Value = True Then
            If TxtBordx1 = "" Then
                MsgBox "Please Enter Bordx No", vbOKOnly + vbCritical, DialogBoxTitle
                TxtBordx1.SetFocus
                CmdExit.Enabled = True
                CmdEdit.Enabled = True
                CmdUpdate.Enabled = True
                cmdEXPtoExcel.Enabled = True
                CmdSave.Enabled = True
                CmdClear.Enabled = True
            Else
            
                Screen.MousePointer = vbHourglass
                
                listrecord = lstBordx.ListItems.Count
                If listrecord <> 0 Then
                    
                    For listloop1 = 1 To listrecord
                        Call BordxNoCheck
                    Next listloop1
                End If
                If Check2 = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    CmdExit.Enabled = True
                    CmdEdit.Enabled = True
                    CmdUpdate.Enabled = True
                    cmdEXPtoExcel.Enabled = True
                    CmdSave.Enabled = True
                    CmdClear.Enabled = True
                ElseIf Check2 = False Then
                    Call SearchRecord
                End If
                Check2 = False
                Screen.MousePointer = Default
                TxtBordx1 = ""
                TxtBordx1.SetFocus
            End If
            TxtBordx1.SetFocus
        Else
            Call SearchRecord
        End If
        Screen.MousePointer = Default
    End If
    Screen.MousePointer = Default
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    CmdUpdate.Enabled = True
    cmdEXPtoExcel.Enabled = True
    CmdSave.Enabled = True
    CmdClear.Enabled = True
End Sub

Private Sub CmdClear_Click()
    lstBordx.ListItems.Clear
    lstWorksheet.ListItems.Clear
    ReDim BordxNoArray(0)
    CboPayor.Clear
    Call PayorComboListing
    cboAccPeriodMth.Text = ""
    cboAccPeriodYR.Text = ""
    TxtBordx1.Text = ""
    TxtBordx2.Text = ""
    TxtReceiptNo.Text = ""
    txtReceiptDate.Text = "__/__/____"
    txtChequeNo.Text = ""
    txtRemark.Text = ""
    CboBordxStatus.ListIndex = 0
    CboWSStatus.ListIndex = 0
    lblcurrent = 0
    lbltotal = 0
    txtReceiptAmt = ""
    lblSelectedAmt = 0
    one = False
    two = False
    CmdSave.Enabled = True
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Const METHOD_NAME = "cmdSave_Click"
On Error GoTo debugs
    Screen.MousePointer = vbHourglass
    If OptBordx.Value = True Then
        Call SaveBordx
    Else
        Call SaveWorksheet
    End If
    Screen.MousePointer = vbDefault
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdEXPtoExcel_Click()
    Screen.MousePointer = vbHourglass
    If OptBordx.Value = True Then
        Call ExportToExcelReceiptMNRBBor(lstBordx)
    Else
        Call ExportToExcelReceiptMNRBWks(lstWorksheet)
    End If
    Screen.MousePointer = vbDefault
End Sub
'************** End Command Button ****************************

'************** Start SearchRecord ****************************
Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim strsql As String
Dim strappend As String
Dim sql As String
 
    strappend = ""
    
    If OptBordx.Value = True Then
       If Len(Trim(CboPayor)) Then
          strappend = strappend + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
       End If
    Else
       If Len(Trim(CboPayor)) Then
          strappend = strappend + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
       End If
    End If
    
    If Len(Trim(TxtBordx1)) Then
      strappend = strappend + " and BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
    End If
    
    If Len(Trim(TxtBordx2)) Then
      strappend = strappend + " and BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
    End If
    
    If Mid(CboWSStatus, 1, 1) = 2 Then
      strappend = strappend + " and (PVRecMNRBStatus = '' or PVRecMNRBStatus is null) "
    End If
    
    If Mid(CboWSStatus, 1, 1) = 1 Then
     strappend = strappend + " and PVRecMNRBStatus = 'F'"
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 1 Then
     strappend = strappend + " and BordxMNRBFrStatus = 'F'"
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 2 Then
     strappend = strappend + " and (BordxMNRBFrStatus = '' or BordxMNRBFrStatus is null) "
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 3 Then
     strappend = strappend + " and BordxMNRBFrStatus = 'P'"
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 4 Then
     strappend = strappend + " and (BordxMNRBFrStatus = 'P' or BordxMNRBFrStatus = 'F') "
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 5 Then
     strappend = strappend + " and (BordxMNRBFrStatus = '' or BordxMNRBFrStatus is null or BordxMNRBFrStatus = 'P') "
    End If
    
    If Len(Trim(txtRECNoFr)) Then
       strappend = strappend + " AND ReceiptNo >= '" & RemStr(Trim(txtRECNoFr)) & "'"
    End If
       
    If Len(Trim(txtRECNoTo)) Then
       strappend = strappend + " AND ReceiptNo <= '" & RemStr(Trim(txtRECNoTo)) & "'"
    End If
        
    If Len(Trim(txtRECDateFr)) > 0 And IsDate(txtRECDateFr) = True Then
        strappend = strappend + " And ReceiptDate >= '" & Format(txtRECDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtRECDateTo)) > 0 And IsDate(txtRECDateTo) = True Then
        strappend = strappend + " And ReceiptDate <= '" & Format(txtRECDateTo, "mm/dd/yyyy") & "'"
    End If
       
    If Len(cboACPeriodMthFr) Then
       strappend = strappend + " AND RECAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
       
    If Len(cboACPeriodMthTo) Then
       strappend = strappend + " AND RECAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If Len(cboACPeriodYRFr) Then
       strappend = strappend + " AND RECAccPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
    End If
       
    If Len(cboACPeriodYRTo) Then
       strappend = strappend + " AND RECAccPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
    
    If ChkScan.Value = 1 And OptWks.Value = True Then
        If Len(txtWSNoFr) Then
           strappend = strappend + " AND CASNumber = '" & Trim(txtWSNoFr) & "'"
        End If
    ElseIf ChkScan.Value = 0 And OptWks.Value = True Then
        If Len(txtWSNoFr) Then
           strappend = strappend + " AND CASNumber >= '" & Trim(txtWSNoFr) & "'"
        End If
        If Len(txtWSNoTo) Then
           strappend = strappend + " AND CASNumber <= '" & Trim(txtWSNoTo) & "'"
        End If
        
        If Len(Trim(TxtBordx1)) Then
          strappend = strappend + " and BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
        
        If Len(Trim(TxtBordx2)) Then
          strappend = strappend + " and BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
        End If
        
    End If
    
    If ChkScan.Value = 1 And OptBordx.Value = True Then
        If Len(Trim(TxtBordx1)) Then
          strappend = strappend + " and BordxRefNo = '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
    ElseIf ChkScan.Value = 0 And OptBordx.Value = True Then
        If Len(Trim(TxtBordx1)) Then
          strappend = strappend + " and BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
        
        If Len(Trim(TxtBordx2)) Then
          strappend = strappend + " and BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
        End If
        
        If Len(txtWSNoFr) Then
           strappend = strappend + " AND CASNumber >= '" & Trim(txtWSNoFr) & "'"
        End If
        If Len(txtWSNoTo) Then
           strappend = strappend + " AND CASNumber <= '" & Trim(txtWSNoTo) & "'"
        End If
        
    End If
    
    If OptBordx.Value = True And ChkScan.Value = 0 Then
       Call ReceiptBordxListing(strappend)
    ElseIf OptBordx.Value = True And ChkScan.Value = 1 Then
       Call ReceiptBordxListing2(strappend)
    ElseIf OptWks.Value = True And ChkScan.Value = 0 Then
       Call ReceiptWorksheetListing(strappend)
    ElseIf OptWks.Value = True And ChkScan.Value = 1 Then
       Call ReceiptWorksheetListing2(strappend)
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
'************** End SearchRecord ****************************

'************** Start SaveBordx ****************************

Private Function SaveBordx()
Const METHOD_NAME = "SaveBordx"
On Error GoTo debugs

Dim ii As Integer
Dim Update
Dim BordxNo As String
Dim caseno As String
Dim CASVer As String
Dim status As String
Dim delimiter1, delimiter2, delimiter3 As Integer
Dim strsqlBordx As String
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim Bordx As New ADODB.Recordset
    
    If lstBordx.ListItems.Count = 0 Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    'ElseIf UBound(BordxNoArray) = 0 Then
        'MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf TxtReceiptNo = "" Then
       MsgBox "Please Enter Receipt No!", vbCritical
       TxtReceiptNo.SetFocus
    ElseIf txtReceiptDate = "__/__/____" Then
       MsgBox "Please Enter Receipt Date !", vbCritical
       txtReceiptDate.SetFocus
    ElseIf txtChequeNo = "" Then
       MsgBox "Please Enter Cheque No!", vbCritical
       txtChequeNo.SetFocus
    ElseIf txtReceiptAmt = "" Then
        MsgBox "Please Enter Receipt Amount!", vbCritical
        txtReceiptAmt.SetFocus
    ElseIf txtReceiptAmt <> lblSelectedAmt And txtRemark = "" Then
        MsgBox "Please Enter Remark!", vbCritical
        txtRemark.SetFocus
    Else
        Update = MsgBox("Do you want to update the record ?", vbYesNo + vbExclamation)
        If Update = vbYes Then
        
            For ii = 0 To UBound(BordxNoArray)
                
                If Trim(BordxNoArray(ii)) <> "" Then
                   delimiter1 = InStr(1, BordxNoArray(ii), "~")
                   caseno = Trim(Mid(BordxNoArray(ii), 1, IIf(delimiter1 = 0, 10, delimiter1 - 1)))
                   'delimiter2 = InStr(delimiter1 + 1, BordxNoArray(ii), "~")
                   'CASVer = Trim(Mid(BordxNoArray(ii), delimiter1 + 1, delimiter2 - delimiter1 - 1))
                   BordxNo = Trim(Mid(BordxNoArray(ii), delimiter1 + 1, Len(BordxNoArray(ii)) - delimiter1))
                                  
                   Call UpdateRecord(caseno, BordxNo)
                End If
                            
            Next ii
                 
            Erase BordxNoArray
            Arraycnt = 0
            ReDim BordxNoArray(0)
            MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
            
            TxtReceiptNo.Text = ""
            txtReceiptDate.Text = "__/__/____"
            txtChequeNo.Text = ""
            txtRemark.Text = ""
            txtReceiptAmt = ""
            cboAccPeriodMth.Text = ""
            cboAccPeriodYR.Text = ""
            lblSelectedAmt = 0
            
            If OptBordx.Value = True And ChkScan.Value = 0 Then
                Call SearchRecord
            ElseIf OptBordx.Value = True And ChkScan.Value = 1 Then
                lstBordx.ListItems.Clear
                TxtBordx1.SetFocus
            End If
            
        End If
    End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Function
'************** End SaveBordx ****************************

Private Function UpdateBordx()
Const METHOD_NAME = "UpdateBordx"
On Error GoTo debugs
Dim RecordSaved
Dim RptPAY As New ADODB.Recordset
Dim UpdateRptPAY As New ADODB.Recordset
Dim strsql As String
Dim RptBOR As New ADODB.Recordset
Dim UpdateRptBOR As New ADODB.Recordset
Dim strsql2 As String
Dim Index As String
Dim Bordx As String
    
    If one = False Then
        ElseIf lstBordx.ListItems.Count = 0 Then
            MsgBox "Please select the record", vbOKOnly + vbCritical
        ElseIf TxtReceiptNo = "" Then
            MsgBox "Please Enter Receipt No!", vbCritical
            TxtReceiptNo.SetFocus
        ElseIf IsDate(txtReceiptDate) = False Then
            MsgBox "Invalid Receipt Date!", vbCritical
            txtReceiptDate.SetFocus
        ElseIf txtReceiptDate = "__/__/____" Then
            MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtReceiptDate.SetFocus
        ElseIf txtChequeNo = "" Then
            MsgBox "Please Enter Cheque No!", vbCritical
            txtChequeNo.SetFocus
        ElseIf txtReceiptAmt = "" Then
            MsgBox "Please Enter Receipt Amt!", vbCritical
            txtReceiptAmt.SetFocus
        Else
            
            Bordx = Trim(lstBordx.SelectedItem.SubItems(2))
            Index = Trim(lstBordx.SelectedItem.SubItems(4))
            
            strsql = "Select * From ReceiptFrMNRB Where BordxNo = '" & Trim(Bordx) & "' AND ReceiptNo = '" & Trim(Index) & "'"
            strsql2 = "Select * From Bordx Where BordxRefNo = '" & Trim(Bordx) & "'"
                                 
            RptPAY.Open strsql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
            RptBOR.Open strsql2, MediConnCISWS, adOpenKeyset, adLockOptimistic
            
            RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                UpdateRptPAY.Open "ReceiptFrMNRB", MediConnCISPayment, adOpenKeyset, adLockOptimistic
                If RptPAY.RecordCount = 0 Then
                    RptPAY.AddNew
                End If
                Do Until RptPAY.EOF
                    With RptPAY
                        !RECEIPTNO = ChkStr(TxtReceiptNo)
                        !ReceiptDate = txtReceiptDate
                        !ReceiptAmt = txtReceiptAmt
                        !ChequeNo = ChkStr(txtChequeNo)
                        !Remark = ChkStr(txtRemark)
                        !MNRBLastUpdateDate = Date
                        !MNRBLastUpdateUser = Logon
                        If cboAccPeriodMth <> "" Then
                            !RECAccPeriodMth = cboAccPeriodMth
                        End If
                        If cboAccPeriodYR <> "" Then
                            !RECAccPeriodYR = cboAccPeriodYR
                        End If
                        .Update
                    End With
                    RptPAY.MoveNext
                Loop
                
                With RptBOR
                    !BordxMNRBRecDate = txtReceiptDate
                End With
                
                MsgBox "Record Updated...", vbOKOnly + vbInformation, "Information"
          
                RptPAY.Close
                Set UpdateRptPAY = Nothing
                Set UpdateRptBOR = Nothing
                
                TxtReceiptNo.Text = ""
                txtReceiptDate.Text = "__/__/____"
                txtChequeNo.Text = ""
                txtRemark.Text = ""
                txtReceiptAmt = ""
                cboAccPeriodMth.Text = ""
                cboAccPeriodYR.Text = ""
                lblSelectedAmt = 0
                Call SearchRecord
                one = False
                CmdSave.Enabled = True
            End If
            'RptPAY.Close
            Set RptPAY = Nothing
        
        CmdSave.Enabled = True
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

'************** Start SaveWorksheet ****************************

Private Function SaveWorksheet()
Const METHOD_NAME = "SaveWorksheet"
On Error GoTo debugs

Dim CASNo As String
Dim CASVer As String
Dim status As String
Dim BordxNo As String
Dim PAYCode As String
Dim ii As Integer
Dim cnt As Integer
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim delimiter1, delimiter2, delimiter3 As Integer
Dim ClmCheck1 As New ADODB.Recordset
Dim ClmCheck2 As New ADODB.Recordset
Dim Update
Dim intloop1 As Integer
Dim intloop2 As Integer
cnt = 0
    If lstWorksheet.ListItems.Count = 0 Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf TxtReceiptNo = "" Then
       MsgBox "Please Enter Receipt No !", vbCritical
       TxtReceiptNo.SetFocus
    ElseIf txtReceiptDate = "__/__/____" Then
        MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDate.SetFocus
    ElseIf txtChequeNo = "" Then
       MsgBox "Please Enter Cheque No. !", vbCritical
       txtChequeNo.SetFocus
    ElseIf txtReceiptAmt <> lblSelectedAmt And txtRemark = "" Then
        MsgBox "Please Enter Remark!", vbCritical
        txtRemark.SetFocus
    Else
        Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
        If Update = vbYes Then
           For ii = 0 To UBound(BordxNoArray)
             If Trim(BordxNoArray(ii)) <> "" Then
                
                delimiter1 = InStr(1, BordxNoArray(ii), "~")
                CASNo = Trim(Mid(BordxNoArray(ii), 1, IIf(delimiter1 = 0, 10, delimiter1 - 1)))
                delimiter2 = InStr(delimiter1 + 1, BordxNoArray(ii), "~")
                'CASVer = Trim(Mid(BordxNoArray(ii), delimiter1 + 1, delimiter2 - delimiter1 - 1))
                'delimiter3 = InStr(delimiter2 + 1, BordxNoArray(ii), "~")
                BordxNo = Trim(Mid(BordxNoArray(ii), delimiter1 + 1, delimiter2 - delimiter1 - 1))
                delimiter3 = InStr(delimiter2 + 1, BordxNoArray(ii), "~")
                status = Trim(Mid(BordxNoArray(ii), delimiter2 + 1, delimiter3 - delimiter2 - 1))
                PaymentAmt = Trim(Mid(BordxNoArray(ii), delimiter3 + 1, delimiter2))
                                
                                
                Call UpdateRecord2(CASNo, BordxNo, status, PaymentAmt)
                
                STRCHECK1 = " Select BordxRefNo " & _
                           " From Claim Where BordxRefno = '" & Trim(BordxNo) & "'"
                
                STRCHECK2 = " Select BordxRefNo, PVRecMNRBStatus From Claim " & _
                           " Where PVRecMNRBStatus = 'F' and BordxRefno = '" & Trim(BordxNo) & "'"
                
                Set ClmCheck1 = MediConnCISWS.Execute(STRCHECK1)
                Set ClmCheck2 = MediConnCISWS.Execute(STRCHECK2)
                        
                intloop1 = 0
                intloop2 = 0
                
                Do Until ClmCheck1.EOF
                    With ClmCheck1
                        intloop1 = intloop1 + 1
                    End With
                    ClmCheck1.MoveNext
                Loop
                
                Do Until ClmCheck2.EOF
                    With ClmCheck2
                        intloop2 = intloop2 + 1
                    End With
                    ClmCheck2.MoveNext
                Loop
                
                If intloop1 = intloop2 Then
                    Call UpdateBordxStatus(BordxNo, "F")
                Else
                    Call UpdateBordxStatus(BordxNo, "P")
                End If
                
             End If
           Next ii
            
           
           Erase BordxNoArray
           Arraycnt = 0
           ReDim BordxNoArray(0)
           MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
           
            TxtReceiptNo.Text = ""
            txtReceiptDate.Text = "__/__/____"
            txtChequeNo.Text = ""
            txtRemark.Text = ""
            txtReceiptAmt = ""
            cboAccPeriodMth.Text = ""
            cboAccPeriodYR.Text = ""
            lblSelectedAmt = 0
           
           If OptWks.Value = True And ChkScan.Value = 0 Then
                Call SearchRecord
           ElseIf OptWks.Value = True And ChkScan.Value = 1 Then
                lstWorksheet.ListItems.Clear
                txtWSNoFr.SetFocus
           End If
           
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
'************** End SaveWorksheet ****************************

Private Function UpdateWorksheet()
Const METHOD_NAME = "UpdateWorksheet"
On Error GoTo debugs

Dim RecordSaved
Dim WKs As New ADODB.Recordset
Dim UpdateWKs As New ADODB.Recordset
Dim strsql As String
Dim BOR As New ADODB.Recordset
Dim UpdateBOR As New ADODB.Recordset
Dim strsql2 As String
Dim Index As String
Dim CASNo As String
Dim WSNo As String
    
    If two = False Then
        ElseIf lstWorksheet.ListItems.Count = 0 Then
            MsgBox "Please select the record", vbOKOnly + vbCritical
        ElseIf TxtReceiptNo = "" Then
            MsgBox "Please Enter Receipt No!", vbCritical
            TxtReceiptNo.SetFocus
        ElseIf IsDate(txtReceiptDate) = False Then
            MsgBox "Invalid Receipt Date!", vbCritical
            txtReceiptDate.SetFocus
        ElseIf txtReceiptDate = "__/__/____" Then
            MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtReceiptDate.SetFocus
        ElseIf txtChequeNo = "" Then
            MsgBox "Please Enter Cheque No!", vbCritical
            txtChequeNo.SetFocus
        ElseIf txtReceiptAmt = "" Then
            MsgBox "Please Enter Receipt Amt!", vbCritical
            txtReceiptAmt.SetFocus
        Else
            
            Bordx = Trim(lstWorksheet.SelectedItem.SubItems(2))
            CASNo = Trim(lstWorksheet.SelectedItem.SubItems(3))
            'WSNo = Trim(Mid(lstWorksheet.SelectedItem.SubItems(3), InStr(1, lstWorksheet.SelectedItem.SubItems(3), "-") + 1, Len(lstWorksheet.SelectedItem.SubItems(3)) - InStr(1, lstWorksheet.SelectedItem.SubItems(3), "-")))
            
            strsql = " Select * From ReceiptFrMNRB Where BordxNo = '" & Trim(Bordx) & "' AND CASNumber = '" & Trim(CASNo) & "'"
                     
            strsql2 = " Select * From Bordx Where BordxRefNo = '" & Trim(Bordx) & "'"
                                 
            WKs.Open strsql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
            BOR.Open strsql2, MediConnCISWS, adOpenKeyset, adLockOptimistic
            
            RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                UpdateWKs.Open "ReceiptFrMNRB", MediConnCISPayment, adOpenKeyset, adLockOptimistic
                If WKs.RecordCount = 0 Then
                    WKs.AddNew
                End If
                Do Until WKs.EOF
                    With WKs
                        !RECEIPTNO = ChkStr(TxtReceiptNo)
                        !ReceiptDate = txtReceiptDate
                        !ReceiptAmt = txtReceiptAmt
                        !ChequeNo = ChkStr(txtChequeNo)
                        !Remark = ChkStr(txtRemark)
                        !MNRBLastUpdateDate = Date
                        !MNRBLastUpdateUser = Logon
                        .Update
                    End With
                    WKs.MoveNext
                Loop
                
                With BOR
                    !BordxMNRBRecDate = txtReceiptDate
                End With
                
                MsgBox "Record Updated...", vbOKOnly + vbInformation, "Information"
          
                WKs.Close
                Set UpdateWKs = Nothing
                Set UpdateBOR = Nothing
                
                TxtReceiptNo.Text = ""
                txtReceiptDate.Text = "__/__/____"
                txtChequeNo.Text = ""
                txtRemark.Text = ""
                txtReceiptAmt = ""
                cboAccPeriodMth.Text = ""
                cboAccPeriodYR.Text = ""
                lblSelectedAmt = 0
                Call SearchRecord
                two = False
                CmdSave.Enabled = True
            End If
            Set WKs = Nothing
            
        CmdSave.Enabled = True
    End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function


Private Sub lstBordx_Click()
Const METHOD_NAME = "lstBordx_Click"
On Error GoTo debugs

Dim rst3 As New ADODB.Recordset
Dim strsql As String
Dim BordxNo As String
Dim RECEIPTNO As String
If one = True Then
CmdSave.Enabled = False
    If lstBordx.ListItems.Count Then
        BordxNo = Trim(lstBordx.SelectedItem.SubItems(2))
        strsql = " SELECT * From ReceiptFrMNRB Where BordxNo = '" & Trim(BordxNo) & "'"
        rst3.Open strsql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
            If rst3.RecordCount = 0 Then
                MsgBox "Payment hasn't been made!", vbOKOnly + vbCritical, DialogBoxTitle
            Else
                TxtReceiptNo.SetFocus
                TxtReceiptNo = rst3!RECEIPTNO
                txtReceiptDate = rst3!ReceiptDate
                txtReceiptAmt = rst3!ReceiptAmt
                txtChequeNo = rst3!ChequeNo
                txtRemark = rst3!Remark
                cboAccPeriodMth = rst3!RECAccPeriodMth
                cboAccPeriodYR = rst3!RECAccPeriodYR
        
                CmdSave.Enabled = False
                rst3.Close
                Set rst3 = Nothing
            End If
    End If
End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

'************** start ItemCheck ****************************

Private Sub lstBordx_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Const METHOD_NAME = "lstBordx_ItemCheck"
On Error GoTo debugs

Dim i, newbound As Integer
Dim tempArray() As String
Dim Found As String
Dim listloop As Integer
Dim ListCount As Integer
Dim TotalAmt As Variant
Dim RECStatus As String
Dim RECAmt As Integer
Dim strsqlBOR As String
Dim strsqlREC As String
Dim BOR As New ADODB.Recordset
Dim REC As New ADODB.Recordset
Dim BordxNo As String
Dim caseno As String
Dim CASVer As String
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Static Arraycnt As Integer

    
    If Item.Checked = True Then
        BordxNo = Trim(Item.SubItems(2))
        
        strsqlBOR = "select BordxRefNo, BordxMNRBFrStatus from Bordx where BordxRefNo = '" & Trim(BordxNo) & "'"
        BOR.Open strsqlBOR, MediConnCISWS, adOpenKeyset, adLockOptimistic
        
        If BOR!BordxMNRBFRStatus = "F" Then
            MsgBox "Record has been paid !", vbCritical
            Item.Checked = False
        Else
        
            If BOR!BordxMNRBFRStatus = "P" Then
                MsgBox "Record has been Partial Paid !", vbInformation
                
                strsqlREC = "select * from ReceiptfrMNRB where BordxNo = '" & Trim(BordxNo) & "'"
                REC.Open strsqlREC, MediConnCISPayment, adOpenKeyset, adLockOptimistic
                
                strsqlCLM = " Select CASNumber from Claim where BordxRefNo = '" & Trim(BordxNo) & "'" & _
                            " AND PVRecMNRBStatus IS NULL "
                CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
            
                If CLM.RecordCount Then
                    Do Until CLM.EOF
                        caseno = CLM!CASNumber
                                            
                        ReDim Preserve BordxNoArray(Arraycnt)
                        BordxNoArray(Arraycnt) = CLM!CASNumber & "~" & Item.SubItems(2)
                        Arraycnt = Arraycnt + 1
                    
                    CLM.MoveNext
                    Loop
                
                    Item.SubItems(6) = "Full"
                    lblSelectedAmt = lblSelectedAmt + CDbl(Format(Item.SubItems(5), "#,##0.00"))
                
                    REC.Close
                    Set REC = Nothing
                    CLM.Close
                    Set CLM = Nothing
                End If
            
            Else
            
            
            strsqlREC = "select * from ReceiptfrMNRB where BordxNo = '" & Trim(BordxNo) & "'"
            REC.Open strsqlREC, MediConnCISPayment, adOpenKeyset, adLockOptimistic
            
            strsqlCLM = "Select CASNumber from Claim where BordxRefNo = '" & Trim(BordxNo) & "'"
            CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
        
            If CLM.RecordCount Then
                Do Until CLM.EOF
                    caseno = CLM!CASNumber
                                    
                    ReDim Preserve BordxNoArray(Arraycnt)
                    BordxNoArray(Arraycnt) = CLM!CASNumber & "~" & Item.SubItems(2)
                    Arraycnt = Arraycnt + 1
                
                CLM.MoveNext
                Loop
            
                Item.SubItems(6) = "Full"
                lblSelectedAmt = lblSelectedAmt + CDbl(Format(Item.SubItems(3), "#,##0.00"))
            
                REC.Close
                Set REC = Nothing
                CLM.Close
                Set CLM = Nothing
            End If
            End If
        End If
        BOR.Close
        Set BOR = Nothing
         
    Else
       
        newbound = 0
       
        If UBound(BordxNoArray) <> 0 Then
            For i = 0 To UBound(BordxNoArray)
                If (Item.SubItems(2)) <> BordxNoArray(i) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = BordxNoArray(i)
                    newbound = newbound + 1
                End If
            Next
          
            If Item.SubItems(6) = "Full" Then
                lblSelectedAmt = lblSelectedAmt - CDbl(Format(Item.SubItems(3), "#,##0.00"))
                Item.SubItems(6) = ""
            End If
          
            Erase BordxNoArray
            ReDim BordxNoArray(0)
            ReDim tempArray(0)
          
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
          
            ReDim BordxNoArray(UBound(tempArray))
            For i = 0 To UBound(tempArray)
                BordxNoArray(i) = tempArray(i)
            Next
            Erase tempArray
            ReDim tempArray(0)
        
        Else
            If Item.SubItems(6) = "Full" Then
                lblSelectedAmt = lblSelectedAmt - CDbl(Format(Item.SubItems(3), "#,##0.00"))
                Item.SubItems(6) = ""
            Else
                Item.SubItems(6) = ""
            End If
          
            Erase BordxNoArray
            ReDim BordxNoArray(0)
            Arraycnt = 0

        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstWorksheet_Click()
Const METHOD_NAME = "lstWorksheet_Click"
On Error GoTo debugs

Dim rst3 As New ADODB.Recordset
Dim strsql As String
Dim Bordx As String
Dim CASNo As String
Dim WSNo As String
If two = True Then
CmdSave.Enabled = False
    If lstWorksheet.ListItems.Count Then
        Bordx = Trim(lstWorksheet.SelectedItem.SubItems(2))
        CASNo = Trim(lstWorksheet.SelectedItem.SubItems(3))
        'WSNo = Trim(Mid(lstWorksheet.SelectedItem.SubItems(3), InStr(1, lstWorksheet.SelectedItem.SubItems(3), "-") + 1, Len(lstWorksheet.SelectedItem.SubItems(3)) - InStr(1, lstWorksheet.SelectedItem.SubItems(3), "-")))
        strsql = " SELECT * From ReceiptFrMNRB Where BordxNo = '" & Trim(Bordx) & "'" & _
                 " AND CASNumber = '" & Trim(CASNo) & "'"
        rst3.Open strsql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
            If rst3.RecordCount = 0 Then
                MsgBox "Payment hasn't been made!", vbOKOnly + vbCritical, DialogBoxTitle
            Else
                TxtReceiptNo.SetFocus
                TxtReceiptNo = rst3!RECEIPTNO
                txtReceiptDate = rst3!ReceiptDate
                txtReceiptAmt = rst3!ReceiptAmt
                txtChequeNo = rst3!ChequeNo
                txtRemark = rst3!Remark
                If Len(rst3!RECAccPeriodMth) Then
                    cboAccPeriodMth = rst3!RECAccPeriodMth
                End If
                If Len(rst3!RECAccPeriodYR) Then
                    cboAccPeriodYR = rst3!RECAccPeriodYR
                End If
                CmdSave.Enabled = False
                rst3.Close
                Set rst3 = Nothing
            End If
    End If
End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstWorksheet_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ProcessCheckedItem(Item)
End Sub

Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
Const METHOD_NAME = "ProcessCheckedItem"
On Error GoTo debugs

Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim temp
Dim listloop As Integer
Dim ListCount As Integer
Dim TotalAmt As Variant
Dim RECStatus As String
Dim RECAmt As Integer
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim strsqlREC As String
Dim REC As New ADODB.Recordset
Dim caseno As String
Dim caseversion As String
Dim status As String
Static Arraycnt As Integer
Dim BordxNo As String
Dim strsqlBOR As String
Dim BOR As New ADODB.Recordset
    
                        
    If Item.Checked = True Then
        BordxNo = Trim(Item.SubItems(2))
        caseno = Trim(Item.SubItems(3))
        'caseversion = Trim(Mid(Item.SubItems(3), InStr(1, Item.SubItems(3), "-") + 1, Len(Item.SubItems(3)) - InStr(1, Item.SubItems(3), "-")))
        
        strsqlBOR = " Select BordxRefNo, PVRecMNRBStatus from Claim where CASNumber = '" & Trim(caseno) & "'" & _
                    " AND BordxRefNo = '" & Trim(BordxNo) & "'"
        BOR.Open strsqlBOR, MediConnCISWS, adOpenKeyset, adLockOptimistic
        
        If BOR!PVRecMNRBStatus = "F" Then
            MsgBox "Record has been paid !", vbCritical
            Item.Checked = False
        Else
            strsqlREC = "select * from ReceiptfrMNRB where BordxNo = '" & Trim(BordxNo) & "'"
            REC.Open strsqlREC, MediConnCISPayment, adOpenKeyset, adLockOptimistic
            
            strsqlCLM = "Select CASNumber from Claim where BordxRefNo = '" & Trim(BordxNo) & "'"
            CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
        
                                
                    If Item.SubItems(7) = "Partial" Then
                        lblSelectedAmt = CDbl(lblSelectedAmt) + CDbl(Format(Item.SubItems(7), "#,##0.00"))
                    Else
                    Item.SubItems(7) = "Full"
                    
                    lblSelectedAmt = CDbl(lblSelectedAmt) + CDbl(Format(Item.SubItems(4), "#,##0.00"))
                    End If
                
                    ReDim Preserve BordxNoArray(Arraycnt)
                    BordxNoArray(Arraycnt) = caseno & "~" & BordxNo & "~" & Mid(Item.SubItems(7), 1, 1) & "~" & Item.SubItems(4) & Item.SubItems(5) & Item.SubItems(6)
                    Arraycnt = Arraycnt + 1
                
                REC.Close
                Set REC = Nothing
                CLM.Close
                Set CLM = Nothing
            
        End If
        BOR.Close
        Set BOR = Nothing
         
    Else
       
        newbound = 0
       
        If UBound(BordxNoArray) <> 0 Then
            For i = 0 To UBound(BordxNoArray)
                If (Item.SubItems(2)) <> BordxNoArray(i) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = BordxNoArray(i)
                    newbound = newbound + 1
                End If
            Next
          
            If Item.SubItems(7) = "Full" Then
                lblSelectedAmt = CDbl(lblSelectedAmt) - CDbl(Format(Item.SubItems(4), "#,##0.00"))
                Item.SubItems(7) = ""
            End If
          
            Erase BordxNoArray
            ReDim BordxNoArray(0)
            ReDim tempArray(0)
          
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
          
            ReDim BordxNoArray(UBound(tempArray))
            For i = 0 To UBound(tempArray)
                BordxNoArray(i) = tempArray(i)
            Next
            Erase tempArray
            ReDim tempArray(0)
        
        Else
            If Item.SubItems(7) = "Full" Then
                lblSelectedAmt = CDbl(Format(lblSelectedAmt, "#,##0.00")) - CDbl(Format(Item.SubItems(4), "#,##0.00"))
                lblSelectedAmt = Format(lblSelectedAmt, "#,##0.00")
                Item.SubItems(7) = ""
            Else
                Item.SubItems(4) = ""
            End If
          
            Erase BordxNoArray
            ReDim BordxNoArray(0)
            Arraycnt = 0

        End If
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
'************** End ItemCheck ****************************

'************** Start ReceiptBordxListing ****************************
Private Sub ReceiptBordxListing(Optional strappend As String)
Const METHOD_NAME = "ReceiptBordxListing"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim CLM As New ADODB.Recordset
Dim strsqlCLM As String
Dim sql As String
Dim sql2 As String
Dim Total As String
Dim Current As String
Dim check As Boolean
Dim RECFull As Boolean
Dim RECPartial As String
Dim Bordx As String
Dim status As String

Total = 0
Current = 0
lblSelectedAmt = 0

lstBordx.ListItems.Clear


sql = " Select BordxRefNo, BordxClaimAmt, BordxPAYCode, " & _
      " BordxMNRBFRStatus, SUM(CheckAmt) AS CheckAmt, " & _
      " RECAccPeriodMth, RECAccPeriodYR " & _
      " From VIEW_CISBORReceiptMNRB where 0 = 0 "

sql2 = " Group by BordxRefNo, BordxClaimAmt, BordxPAYCode, " & _
       " BordxMNRBFRStatus, " & _
       " RECAccPeriodMth, RECAccPeriodYR "
    
    If Len(Trim(strappend)) Then
        sql = sql + strappend + sql2
    Else
        sql = sql + sql2
    End If

rst2.Open sql, MediConnCISWS, adOpenKeyset, adLockOptimistic

Do
    Do Until rst2.EOF
    
    TxtPStatus = ""
        
    If Len(rst2!BordxMNRBFRStatus) Then
        TxtPStatus = rst2!BordxMNRBFRStatus
    End If
    
    With rst2
    
        Set Record = lstBordx.ListItems.Add(, , "")
               
        If Len(!BordxPAYCode) Then
           Record.SubItems(1) = !BordxPAYCode
        End If
        If Len(!Bordxrefno) Then
           Record.SubItems(2) = Trim(!Bordxrefno)
        End If
        If Len(!BordxClaimAmt) Then
           Record.SubItems(3) = Format(!BordxClaimAmt, "#,##0.00")
        End If
        
        If Len(!CheckAmt) Then
           Record.SubItems(4) = Format(!CheckAmt, "#,##0.00")
        End If
        
        If Len(!CheckAmt) Then
            Record.SubItems(5) = Format(!BordxClaimAmt, "#,##0.00") - Format(!CheckAmt, "#,##0.00")
            Record.SubItems(5) = Format(Record.SubItems(5), "#,##0.00")
        Else
            Record.SubItems(5) = Format(!BordxClaimAmt, "#,##0.00")
        End If
        
        If TxtPStatus = "F" Then
            Record.SubItems(6) = "Full"
        ElseIf TxtPStatus = "P" Then
            Record.SubItems(6) = "Partial"
        ElseIf TxtPStatus = "" Then
            Record.SubItems(6) = ""
        End If
        
        Current = Current + 1
        lblcurrent = Current
    End With
    rst2.MoveNext
        
    DoEvents
        If intExit = 1 Then
            check = False
            Exit Do
        End If
     
    Loop
   
Loop Until check = False
intExit = 0

If lstBordx.ListItems.Count = 0 Then
    MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    lblcurrent = 0
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    CmdUpdate.Enabled = True
    cmdEXPtoExcel.Enabled = True
    CmdSave.Enabled = True
    CmdClear.Enabled = True
End If

rst2.Close
Set rst2 = Nothing
CmdExit.Enabled = True
CmdEdit.Enabled = True
CmdUpdate.Enabled = True
cmdEXPtoExcel.Enabled = True
CmdSave.Enabled = True
CmdClear.Enabled = True

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ReceiptBordxListing2(Optional strappend As String)
Const METHOD_NAME = "ReceiptBordxListing2"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim CLM As New ADODB.Recordset
Dim strsqlCLM As String
Dim sql As String
Dim sql2 As String
Dim Total As String
Dim Current As String
Dim check As Boolean
Dim RECFull As Boolean
Dim RECPartial As String
Dim Bordx As String
Dim status As String

Total = 0
Current = 0
lblSelectedAmt = 0

sql = " Select BordxRefNo, BordxClaimAmt, BordxPAYCode, " & _
      " BordxMNRBFRStatus, SUM(CheckAmt) AS CheckAmt, " & _
      " RECAccPeriodMth, RECAccPeriodYR " & _
      " From VIEW_CISBORReceiptMNRB where 0 = 0 "

sql2 = " Group by BordxRefNo, BordxClaimAmt, BordxPAYCode, " & _
       " BordxMNRBFRStatus, " & _
       " RECAccPeriodMth, RECAccPeriodYR "
    
    If Len(Trim(strappend)) Then
        sql = sql + strappend + sql2
    Else
        sql = sql + sql2
    End If

rst2.Open sql, MediConnCISWS, adOpenKeyset, adLockOptimistic

Do
    Do Until rst2.EOF
    
    TxtPStatus = ""
        
    If Len(rst2!BordxMNRBFRStatus) Then
        TxtPStatus = rst2!BordxMNRBFRStatus
    End If
    
    With rst2
    
        Set Record = lstBordx.ListItems.Add(, , "")
               
        If Len(!BordxPAYCode) Then
           Record.SubItems(1) = !BordxPAYCode
        End If
        If Len(!Bordxrefno) Then
           Record.SubItems(2) = Trim(!Bordxrefno)
        End If
        If Len(!BordxClaimAmt) Then
           Record.SubItems(3) = Format(!BordxClaimAmt, "#,##0.00")
        End If
        
        If Len(!CheckAmt) Then
           Record.SubItems(4) = Format(!CheckAmt, "#,##0.00")
        End If
        
        If Len(!CheckAmt) Then
            Record.SubItems(5) = Format(!BordxClaimAmt, "#,##0.00") - Format(!CheckAmt, "#,##0.00")
            Record.SubItems(5) = Format(Record.SubItems(5), "#,##0.00")
        Else
            Record.SubItems(5) = Format(!BordxClaimAmt, "#,##0.00")
        End If
        
        If TxtPStatus = "F" Then
            Record.SubItems(6) = "Full"
        ElseIf TxtPStatus = "P" Then
            Record.SubItems(6) = "Partial"
        ElseIf TxtPStatus = "" Then
            Record.SubItems(6) = ""
        End If
        
        Current = Current + 1
        lblcurrent = Current
    End With
    rst2.MoveNext
        
    DoEvents
        If intExit = 1 Then
            check = False
            Exit Do
        End If
     
    Loop
   
Loop Until check = False
intExit = 0

If lstBordx.ListItems.Count = 0 Then
    MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    lblcurrent = 0
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    CmdUpdate.Enabled = True
    cmdEXPtoExcel.Enabled = True
    CmdSave.Enabled = True
    CmdClear.Enabled = True
End If

rst2.Close
Set rst2 = Nothing
CmdExit.Enabled = True
CmdEdit.Enabled = True
CmdUpdate.Enabled = True
cmdEXPtoExcel.Enabled = True
CmdSave.Enabled = True
CmdClear.Enabled = True

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub
'************** End ReceiptBordxListing ****************************

'************** Start ReceiptWorksheetListing ****************************

Private Sub ReceiptWorksheetListing(Optional strappend As String)
Const METHOD_NAME = "ReceiptWorksheetListing"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim sql As String
Dim Total As Integer
Dim Current As Integer
Dim check As Boolean
Dim RECFull As Boolean
Dim RECPartial As String
Dim CASNumber As String
Dim WSVer As String
Dim INVAmt As Variant

lstWorksheet.ListItems.Clear

sql = " Select CASNumber, BordxRefNo, CLMTotalApproved, " & _
      " ReceiptDate, ReceiptNo, RECAccPeriodMth, RECAccPeriodYR, " & _
      " BordxPAYCode, PVRecMNRBStatus, CheckAmt From VIEW_CISWKSReceiptMNRB where 0 = 0"

If Len(Trim(strappend)) Then
   sql = sql + strappend
End If

rst2.Open sql, MediConnCISWS, adOpenKeyset, adLockOptimistic

If rst2.RecordCount = 0 Then
    MsgBox "No record found !", vbOKOnly + vbCritical, DialogBoxTitle
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    CmdUpdate.Enabled = True
    cmdEXPtoExcel.Enabled = True
    CmdSave.Enabled = True
    CmdClear.Enabled = True
End If
Do
    Do Until rst2.EOF
    TxtPStatus = ""
    If Len(rst2!PVRecMNRBStatus) Then
        TxtPStatus = rst2!PVRecMNRBStatus
    End If
    
    With rst2
        Set Record = lstWorksheet.ListItems.Add(, , "")
               
        If Len(!BordxPAYCode) Then
           Record.SubItems(1) = !BordxPAYCode
        End If
        If Len(!Bordxrefno) Then
           Record.SubItems(2) = !Bordxrefno
        End If
        If Len(!CASNumber) Then
           Record.SubItems(3) = Trim(!CASNumber)
        End If
        
        CASNumber = rst2!CASNumber
                
        If Len(!CLMTotalApproved) Then
           Record.SubItems(4) = Format(!CLMTotalApproved, "#,##0.00")
        End If
        
        If Len(!RECEIPTNO) Then
           Record.SubItems(5) = Trim(!RECEIPTNO)
        End If
        If Len(!ReceiptDate) Then
           Record.SubItems(6) = !ReceiptDate
        End If
                
        If TxtPStatus = "F" Then
            Record.SubItems(7) = "Full"
        ElseIf TxtPStatus = "P" Then
            Record.SubItems(7) = "Partial"
        Else
            Record.SubItems(7) = ""
        End If
        
        Current = Current + 1
        lblcurrent = Current
    End With
    rst2.MoveNext
    
    DoEvents
        If intExit = 1 Then
            check = False
            Exit Do
        End If
    Loop
Loop Until check = False
intExit = 0
rst2.Close
Set rst2 = Nothing
CmdExit.Enabled = True
CmdEdit.Enabled = True
CmdUpdate.Enabled = True
cmdEXPtoExcel.Enabled = True
CmdSave.Enabled = True
CmdClear.Enabled = True

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ReceiptWorksheetListing2(Optional strappend As String)
Const METHOD_NAME = "ReceiptWorksheetListing2"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim sql As String
Dim Total As Integer
Dim Current As Integer
Dim check As Boolean
Dim RECFull As Boolean
Dim RECPartial As String
Dim CASNumber As String
Dim WSVer As String
Dim INVAmt As Variant

sql = " Select CASNumber, BordxRefNo, CLMTotalApproved, " & _
      " ReceiptDate, ReceiptNo, RECAccPeriodMth, RECAccPeriodYR, " & _
      " BordxPAYCode, PVRecMNRBStatus, CheckAmt From VIEW_CISWKSReceiptMNRB where 0 = 0"

If Len(Trim(strappend)) Then
   sql = sql + strappend
End If
rst2.Open sql, MediConnCISWS, adOpenKeyset, adLockOptimistic
If rst2.RecordCount = 0 Then
    MsgBox "No record found !", vbOKOnly + vbCritical, DialogBoxTitle
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    CmdUpdate.Enabled = True
    cmdEXPtoExcel.Enabled = True
    CmdSave.Enabled = True
    CmdClear.Enabled = True
End If
Do
    Do Until rst2.EOF
    TxtPStatus = ""
    If Len(rst2!PVRecMNRBStatus) Then
        TxtPStatus = rst2!PVRecMNRBStatus
    End If
    
    With rst2
        Set Record = lstWorksheet.ListItems.Add(, , "")
        If Len(!BordxPAYCode) Then
           Record.SubItems(1) = !BordxPAYCode
        End If
        If Len(!Bordxrefno) Then
           Record.SubItems(2) = !Bordxrefno
        End If
        If Len(!CASNumber) Then
           Record.SubItems(3) = Trim(!CASNumber)
        End If
        CASNumber = rst2!CASNumber
        If Len(!CLMTotalApproved) Then
           Record.SubItems(4) = Format(!CLMTotalApproved, "#,##0.00")
        End If
        If Len(!RECEIPTNO) Then
           Record.SubItems(5) = Trim(!RECEIPTNO)
        End If
        If Len(!ReceiptDate) Then
           Record.SubItems(6) = !ReceiptDate
        End If
        If TxtPStatus = "F" Then
            Record.SubItems(7) = "Full"
        ElseIf TxtPStatus = "P" Then
            Record.SubItems(7) = "Partial"
        Else
            Record.SubItems(7) = ""
        End If
        Current = Current + 1
        lblcurrent = Current
    End With
    rst2.MoveNext
    
    DoEvents
        If intExit = 1 Then
            check = False
            Exit Do
        End If
    Loop
Loop Until check = False
intExit = 0
rst2.Close
Set rst2 = Nothing
CmdExit.Enabled = True
CmdEdit.Enabled = True
CmdUpdate.Enabled = True
cmdEXPtoExcel.Enabled = True
CmdSave.Enabled = True
CmdClear.Enabled = True

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'************** End ReceiptWorksheetListing ****************************


Private Sub OptBordx_GotFocus()
    txtWSNoFr.Enabled = False
    txtWSNoTo.Enabled = False
    txtWSNoFr = ""
    txtWSNoTo = ""
End Sub


Private Sub TxtBordx1_LostFocus()
    TxtBordx1 = ChkStr(TxtBordx1)
    CmdSearch.Default = False
End Sub

Private Sub TxtBordx2_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtBordx2_LostFocus()
    TxtBordx2 = ChkStr(TxtBordx2)
End Sub

'*************************** Start Diable Non Integer Value
Private Sub TxtReceiptAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

'*************************** Start Combobox Change/ KeyPress

Private Sub CboPayor_Change()
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Dim strsql As String

  If KeyAscii >= 32 And KeyAscii <> 127 Then
    'NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayor.ListCount - 1
 
        If NewText = LCase(Left(CboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayor.List(i)
        End If
    Next i
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboPayor.Text = ValidValue
            CboPayor.SelStart = 0
            CboPayor.SelLength = Len(ValidValue)
        End If
  End If
End Sub
'*************************** End Combobox Change/ KeyPress

'*************************** Start UpdateRecord

Private Sub UpdateRecord(caseno As String, BordxNo As String)
Const METHOD_NAME = "UpdateRecord"
On Error GoTo debugs
Dim Update
Dim strsqlCLM As String
Dim strsqlREC As String
Dim strsqlBordx As String
Dim CLM As New ADODB.Recordset
Dim REC As New ADODB.Recordset
Dim Bordx As New ADODB.Recordset
Dim PAYAmt As String
Dim SearchDB As New ADODB.Recordset
Dim sql As String

sql = " Select CASNumber, BordxRefNo, CLMTotalApproved, PVRecMNRBStatus " & _
      " From Claim where CASNumber = '" & Trim(caseno) & "' And BordxRefNo = '" & Trim(BordxNo) & "'"

SearchDB.Open sql, MediConnCISWS, adOpenKeyset, adLockOptimistic
REC.Open "ReceiptfrMNRB", MediConnCISPayment, adOpenKeyset, adLockOptimistic
If SearchDB.RecordCount Then
    PAYAmt = SearchDB!CLMTotalApproved
End If
    REC.AddNew
    With REC
        !RECEIPTNO = TxtReceiptNo
        !ReceiptDate = txtReceiptDate
        !ReceiptAmt = txtReceiptAmt
        !ChequeNo = txtChequeNo
        !BordxNo = BordxNo
        !CASNumber = caseno
        !Remark = Trim(txtRemark)
        !MNRBLastUpdateDate = Date
        !MNRBLastUpdateUser = Logon
        If cboAccPeriodMth <> "" Then
            !RECAccPeriodMth = cboAccPeriodMth
        End If
        If cboAccPeriodYR <> "" Then
            !RECAccPeriodYR = cboAccPeriodYR
        End If
        !RECMNRBAmt = PAYAmt
        !CheckAmt = PAYAmt
    End With
    REC.Update
    REC.Close
    Set REC = Nothing
    
    With SearchDB
        !PVRecMNRBStatus = "F"
    End With
    SearchDB.Update
    
    strsqlBordx = " select * from Bordx where BordxrefNo = '" & Trim(BordxNo) & "'"
    Bordx.Open strsqlBordx, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    With Bordx
            !BordxMNRBRecDate = txtReceiptDate
            !BordxMNRBFRStatus = "F"
    End With
    Bordx.Update
    Bordx.Close
    Set Bordx = Nothing
    Set SearchDB = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub UpdateRecord2(CASNo As String, BordxNo As String, status As String, PaymentAmt As Variant)
Const METHOD_NAME = "UpdateRecord2"
On Error GoTo debugs
Dim Update
Dim strsql As String
Dim REC1 As New ADODB.Recordset
Dim REC2 As New ADODB.Recordset
Dim InvCheck As New ADODB.Recordset
Dim strsqlREC1 As String
Dim strsqlREC2 As String
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
    strsqlREC1 = " Select * from ReceiptfrMNRB " & _
                 " Where CASNumber = '" & Trim(CASNo) & "' AND BordxNo = '" & Trim(BordxNo) & "'"
    REC1.Open strsqlREC1, MediConnCISPayment, adOpenKeyset, adLockOptimistic
    strsqlCLM = " Select CASNumber, BordxRefNo, PVRecMNRBStatus From Claim " & _
                " Where BordxRefNo = '" & Trim(BordxNo) & "' And CASNumber = '" & Trim(CASNo) & "'"
    CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    If REC1.RecordCount = 0 Then
        REC1.AddNew
    End If
        
        With REC1
            !CASNumber = ChkStr(CASNo)
            !RECEIPTNO = ChkStr(TxtReceiptNo)
            !ReceiptDate = Format(txtReceiptDate, "dd/mm/yyyy")
            !ReceiptAmt = CDbl(lblSelectedAmt)
            !ChequeNo = ChkStr(txtChequeNo)
            !Remark = ChkStr(txtRemark)
            !BordxNo = ChkStr(BordxNo)
            !MNRBLastUpdateDate = Date
            !MNRBLastUpdateUser = Logon
            !RECMNRBAmt = txtReceiptAmt
            !CheckAmt = PaymentAmt
            If cboAccPeriodMth <> "" Then
                !RECAccPeriodMth = cboAccPeriodMth
            End If
            If cboAccPeriodYR <> "" Then
                !RECAccPeriodYR = cboAccPeriodYR
            End If
            REC1.Update
        End With
        With CLM
            !PVRecMNRBStatus = status
        End With
        CLM.Update
        REC1.Close
        Set REC1 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub UpdateBordxStatus(BordxNo As String, status As String)
Const METHOD_NAME = "UpdateBordxStatus"
On Error GoTo debugs

    Dim strsqlBordx  As String
    Dim Bordx As New ADODB.Recordset
    
    strsqlBordx = "select * from Bordx where BordxrefNo = '" & Trim(BordxNo) & "'"
    Bordx.Open strsqlBordx, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    With Bordx
        !BordxMNRBFRStatus = status
        !BordxMNRBRecDate = txtReceiptDate
    End With
    Bordx.Update
    Bordx.Close
    Set Bordx = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    

End Sub
'*************************** End UpdateRecord

'*************************** Start Option Click

Private Sub OptBordx_Click()
     If OptBordx.Value = True Then
        ChkScan.Value = 1
        lstBordx.Visible = True
        CboBordxStatus.Visible = True
        CboWSStatus.Visible = False
        CboBordxStatus.ListIndex = 0
        CboWSStatus.ListIndex = 0
        TxtReceiptNo = ""
        txtReceiptDate = "__/__/____"
        txtChequeNo = ""
        txtReceiptAmt = ""
        txtRemark = ""
        lstWorksheet.Visible = False
        txtWSNoFr.Enabled = False
        txtWSNoTo.Enabled = False
        txtWSNoTo.Visible = True
        Label21.Visible = True
        txtWSNoFr = ""
        txtWSNoTo = ""
        ChkScan.Enabled = True
        Label14.Visible = False
        TxtBordx2.Visible = False
        TxtBordx1.SetFocus
        ChkScan.Left = 3360
        ChkScan.Top = 840
     End If
End Sub

Private Sub OptWks_Click()
    If OptWks.Value = True Then
        ChkScan.Value = 1
       lstWorksheet.Visible = True
       TxtReceiptNo = ""
       txtReceiptDate = "__/__/____"
       txtReceiptAmt = ""
       txtChequeNo = ""
       txtRemark = ""
       CboBordxStatus.Visible = False
       CboWSStatus.Visible = True
       CboWSStatus.ListIndex = 0
       CboBordxStatus.ListIndex = 0
       lstBordx.Visible = False
       TxtBordx2.Visible = True
       txtWSNoFr.Enabled = True
       Label21.Visible = False
       Label14.Visible = True
       txtWSNoTo.Visible = False
       txtWSNoFr.SetFocus
       ChkScan.Enabled = True
       ChkScan.Value = 1
       ChkScan.Left = 3360
       ChkScan.Top = 600
    End If
End Sub
'*************************** End Option Click

'************** Start ComboListing ****************************
Private Sub ComboListing()
    
    With CboBordxStatus
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
        .AddItem "3 - Partial Paid"
        .AddItem "4 - Paid & Partial Paid"
        .AddItem "5 - Unpaid & Partial Paid"
    End With
    
    With CboWSStatus
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
    End With
    
End Sub

Private Sub PayorComboListing()
Const METHOD_NAME = "PayorComboListing"
On Error GoTo debugs

Dim PAY As New ADODB.Recordset
PAY.Open "Select PAYCode, PAYCompanyName, HLTCode from PAY Where HLTCode = 'S'", medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until PAY.EOF
        CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'************** End ComboListing ****************************

Private Sub txtReceiptAmt_LostFocus()
    If txtReceiptAmt <> lblSelectedAmt Then
        MsgBox "Receipt Amt & Checked Amt Are Different! " & vbNewLine & _
               "Please Fill in Remark.", vbOKOnly + vbCritical
    End If
    
End Sub

Private Sub txtReceiptDate_LostFocus()
    If IsDate(txtReceiptDate) Then
    Else
    If txtReceiptDate <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtReceiptDate.SetFocus
    End If
    End If
End Sub


Private Sub lstBordx_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    'Case 5
        'SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(5)
    'Case 6
        'SortListView lstBordx, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 5
        SortListView lstBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(7)
    End Select
End Sub

Private Sub lstWorksheet_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstWorksheet, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstWorksheet, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
    Case 8
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(8)
    End Select
End Sub

Private Function ClearForm(frmCall)
On Error Resume Next
Dim ctl As Control

For Each ctl In frmCall.Controls
    If TypeOf ctl Is TextBox Then
        ctl.Text = ""
    ElseIf TypeOf ctl Is ListBox Then
            ctl.Clear
    ElseIf TypeOf ctl Is ComboBox Then
            ctl.Text = ""
    ElseIf TypeOf ctl Is MaskEdBox Then
        ctl.Text = "__/__/____"
    End If
Next ctl
    txtACPeriod.mask = ""
    txtACPeriod.Text = ""
    txtACPeriod.mask = "##/####"
End Function

Private Function CaseNoCheck()
    If Trim(ChkStr(txtWSNoFr)) = Trim(lstWorksheet.ListItems(listloop1).SubItems(3)) Then
        Check1 = True
    End If
End Function

Private Function BordxNoCheck()
    If Trim(ChkStr(TxtBordx1)) = Trim(lstBordx.ListItems(listloop1).SubItems(2)) Then
        Check2 = True
    End If
End Function

Private Sub txtWSNoFr_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub txtWSNoFr_LostFocus()
    CmdSearch.Default = False
    txtWSNoFr = ChkStr(txtWSNoFr)
End Sub
