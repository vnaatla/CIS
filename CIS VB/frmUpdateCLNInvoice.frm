VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdateCLNInvoice 
   Caption         =   "Update Clinic Invoice"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11856
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   11856
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdCheck 
      Caption         =   "&Check All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   43
      Top             =   2040
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   19
      Top             =   360
      Width           =   11775
      Begin VB.TextBox TxtWSNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   0
         Top             =   255
         Width           =   1560
      End
      Begin VB.TextBox TxtWSNoTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   1
         Top             =   255
         Width           =   1580
      End
      Begin VB.CheckBox ChkInvNo 
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
         Left            =   5280
         TabIndex        =   7
         Top             =   1120
         Width           =   255
      End
      Begin VB.CheckBox ChkClinic 
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
         Left            =   5280
         TabIndex        =   3
         Top             =   550
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpComp 
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
         Left            =   5280
         TabIndex        =   5
         Top             =   840
         Width           =   375
      End
      Begin MSMask.MaskEdBox txtVisitDateFr 
         Height          =   315
         Left            =   7680
         TabIndex        =   8
         Top             =   240
         Width           =   1335
         _ExtentX        =   2350
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
      Begin MSMask.MaskEdBox txtVisitDateTo 
         Height          =   315
         Left            =   9480
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         _ExtentX        =   2350
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
      Begin MSMask.MaskEdBox txtInvDateFr 
         Height          =   315
         Left            =   7680
         TabIndex        =   10
         Top             =   480
         Width           =   1335
         _ExtentX        =   2350
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
      Begin MSMask.MaskEdBox txtInvDateTo 
         Height          =   315
         Left            =   9480
         TabIndex        =   11
         Top             =   480
         Width           =   1335
         _ExtentX        =   2350
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
      Begin VB.ComboBox CboCLN 
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   3615
      End
      Begin VB.ComboBox CboPayor 
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   825
         Width           =   3615
      End
      Begin VB.TextBox TxtInvNo 
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
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1115
         Width           =   3615
      End
      Begin VB.TextBox TxtInvAmtTo 
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
         Left            =   9480
         MaxLength       =   8
         TabIndex        =   13
         Top             =   760
         Width           =   1335
      End
      Begin VB.TextBox txtInvAmtfr 
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
         Left            =   7680
         MaxLength       =   8
         TabIndex        =   12
         Top             =   760
         Width           =   1320
      End
      Begin VB.Label Label21 
         Caption         =   "To"
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
         Index           =   1
         Left            =   9120
         TabIndex        =   40
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Visit Date"
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
         Index           =   1
         Left            =   6480
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Claim Number"
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
         Left            =   120
         TabIndex        =   30
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   29
         Top             =   285
         Width           =   375
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   5580
         TabIndex        =   28
         Top             =   1170
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "To"
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
         Index           =   0
         Left            =   9120
         TabIndex        =   27
         Top             =   510
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Invoice Date"
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
         Index           =   0
         Left            =   6480
         TabIndex        =   26
         Top             =   510
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "To"
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
         Left            =   9120
         TabIndex        =   25
         Top             =   795
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Invoice Amount"
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
         Left            =   6480
         TabIndex        =   24
         Top             =   795
         Width           =   1575
      End
      Begin VB.Label Label26 
         Caption         =   "Invoice No"
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
         Left            =   120
         TabIndex        =   23
         Top             =   1170
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Group Company"
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
         Left            =   120
         TabIndex        =   22
         Top             =   885
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Clinic"
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
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label36 
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
         Left            =   5320
         TabIndex        =   20
         Top             =   440
         Width           =   255
      End
   End
   Begin VB.Frame FramePV 
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   6480
      Width           =   11775
      Begin VB.TextBox TxtInvoiceNo 
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
         Left            =   960
         MaxLength       =   30
         TabIndex        =   14
         Top             =   200
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtInvoiceDate 
         Height          =   315
         Left            =   4080
         TabIndex        =   15
         Top             =   200
         Width           =   1215
         _ExtentX        =   2138
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
      Begin VB.Label Label39 
         Caption         =   "* - Search Empty Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   8280
         TabIndex        =   42
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label37 
         Caption         =   "* Listing"
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
         Left            =   10680
         TabIndex        =   41
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Invoice Date"
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
         Left            =   3120
         TabIndex        =   18
         Top             =   195
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Invoice No"
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
         Left            =   120
         TabIndex        =   17
         Top             =   200
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstInvoiceListing 
      Height          =   4095
      Left            =   0
      TabIndex        =   37
      Top             =   2400
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   7218
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Case No"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "MBM Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cover ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Visit Date"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Clinic Name"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Inv. No"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Inv Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Inv Reg Date"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Inv Entered By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Inv Amt"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "FFS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Approved Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Grp Company"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   0
      TabIndex        =   31
      Top             =   7080
      Width           =   11775
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Top             =   210
         Width           =   885
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print to &File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9480
         TabIndex        =   36
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8040
         TabIndex        =   35
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11040
         TabIndex        =   34
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10440
         TabIndex        =   33
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
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
         Height          =   330
         Left            =   8760
         TabIndex        =   32
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
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
         Left            =   1080
         TabIndex        =   44
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   " Update Clinic Invoice"
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
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmUpdateCLNInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpCASnumber As String
Dim strAppend As String

Private Sub CboCLN_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboCLN_KeyPress"
On Error GoTo debugs

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCLN.Text, CboCLN.SelStart) + Chr(KeyAscii) + Mid(CboCLN.Text, CboCLN.SelStart + CboCLN.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To CboCLN.ListCount - 1
        If NewText = LCase(Left(CboCLN.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCLN.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCLN.Text = ValidValue
                CboCLN.SelStart = 0
                CboCLN.SelLength = Len(ValidValue)
            End If
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub Form_Load()
  Call ActionLogUpdate(12, "frmUpdateCLNInvoice")
End Sub
Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboPayor_KeyPress"
On Error GoTo debugs

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To CboPayor.ListCount - 1
        If NewText = LCase(Left(CboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCLN.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPayor.Text = ValidValue
                CboPayor.SelStart = 0
                CboPayor.SelLength = Len(ValidValue)
            End If
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ChkClinic_Click()
Const METHOD_NAME = "ChkClinic_Click"
On Error GoTo debugs
Dim ClnRec As New ADODB.Recordset
    CboCLN.Clear
    If ChkClinic.Value = 1 Then
            ClnRec.Open "Select CLNCode, CLNName from CLN WHERE CLNPanel = 'Y' OR CLNPanel = 'C'", medixmasterconn, adOpenForwardOnly, adLockReadOnly
            Do Until ClnRec.EOF
                If Len(ClnRec!CLNName) Then
                    CboCLN.AddItem ClnRec!CLNCode & " - " & Trim(ClnRec!CLNName)
                End If
                ClnRec.MoveNext
            Loop
            ClnRec.Close
            Set ClnRec = Nothing
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ChkGrpComp_Click()
Const METHOD_NAME = "ChkGrpComp_Click"
On Error GoTo debugs
    Dim Company As New ADODB.Recordset
    Company.Open "Select GRPCoName from GRPCompany", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do Until Company.EOF
        CboPayor.AddItem Trim(Company!GrpCoName)
        Company.MoveNext
    Loop
    Company.Close
    Set Company = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdCheck_Click()
Const METHOD_NAME = "CmdCheck_Click"
On Error GoTo debugs
    Dim LstCnt As Integer
    If lstInvoiceListing.ListItems.Count <> 0 Then
        LstCnt = lstInvoiceListing.ListItems.Count
    
        For LstCnt = 1 To LstCnt
            If lstInvoiceListing.ListItems(LstCnt).Checked = False Then
                lstInvoiceListing.ListItems(LstCnt).Checked = True
                lstInvoiceListing.ListItems(LstCnt).Text = "Checked"
            ElseIf lstInvoiceListing.ListItems(LstCnt).Checked = True Then
                lstInvoiceListing.ListItems(LstCnt).Checked = False
                lstInvoiceListing.ListItems(LstCnt).Text = ""
            End If
        Next LstCnt
    Else
        MsgBox "Please select a record.", vbInformation
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
    
    If lstInvoiceListing.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstInvoiceListing)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSearch_Click()
Const METHOD_NAME = "CmdSearch_Click"
On Error GoTo debugs
    Call SearchRecord
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdUpdate_Click()

Const METHOD_NAME = "CmdUpdate_Click"
On Error GoTo debugs
Dim tmpBordxNo As String
Dim i As Integer
Dim Update
Dim ItemChecked As Boolean
Dim PymtIndex As String
Dim ZZ As String
Dim TempBank As String
Dim PVSql As String
Dim BankRec As New ADODB.Recordset
Dim BankSql As String
Dim TmpPVCode As String
Dim UPdRec As New ADODB.Recordset
Dim StrInvSql As String
Dim tmpNonGpAmt As String
Dim tmpFFSAmt As Double
Dim PVNumber As String
Dim CLNCnt As Integer
Dim CompTrans As Boolean
Dim tmpstr As String
    ItemChecked = False
    
    For i = 1 To lstInvoiceListing.ListItems.Count
        If lstInvoiceListing.ListItems(i).Checked = True Then
            ItemChecked = True
            Exit For
        End If
    Next i
    CLNCnt = 0
    
    If ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    ElseIf IsDate(txtInvoiceDate) = False Then
        MsgBox "Invalid Invoice Date!", vbCritical
    ElseIf TxtInvoiceNo = "" Then
        MsgBox "Please Enter Cheque Amt!", vbCritical
        TxtInvoiceNo.SetFocus
    ElseIf txtInvoiceDate = "__/__/____" Then
        MsgBox "Please Enter Invoice Date! ", vbCritical
        txtInvoiceDate.SetFocus
    Else
        Update = MsgBox("Do you want to update the record ? ", vbYesNo + vbExclamation)
        If Update = vbYes Then
            CompTrans = False
            MediConnCISDB.beginTrans
            For i = 1 To lstInvoiceListing.ListItems.Count
                If lstInvoiceListing.ListItems(i).Checked = True Then
                    tmpCASnumber = Trim(lstInvoiceListing.ListItems(i).SubItems(1))
                    Call UpdateRecord(tmpCASnumber)
                End If
            Next i
            MediConnCISDB.commitTrans
            CompTrans = True
            lstInvoiceListing.ListItems.Clear
            tmpstr = " and BillNo = '" & Trim(TxtInvoiceNo) & "' And BillDate = '" & Format(txtInvoiceDate, "mm/dd/yyyy") & "'"
            Call InvoiceListing(tmpstr)
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            
        End If
ErrorSave:
        Screen.MousePointer = vbDefault
        If CompTrans = False Then
            MediConnCISDB.RollbackTrans
            MsgBox Err.Description
        End If
    End If
    
debugs:
If CompTrans = False Then
    MediConnCISDB.RollbackTrans
    MsgBox Err.Description
End If
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub lstInvoiceListing_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = True Then
        Item.Text = "Checked"
    Else
        Item.Text = ""
    End If
End Sub

Private Sub txtInvDateFr_LostFocus()
    If IsDate(txtInvDateFr) Then
    Else
        If txtInvDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtInvDateFr.SetFocus
        End If
    End If
End Sub

Private Sub txtInvDateTo_LostFocus()
    If IsDate(txtInvDateTo) Then
    Else
        If txtInvDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtInvDateTo.SetFocus
        End If
    End If
End Sub

Private Sub txtInvoiceDate_LostFocus()
    If IsDate(txtInvoiceDate) Then
    Else
        If txtInvoiceDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtInvoiceDate.SetFocus
        End If
    End If

End Sub

Private Sub txtVisitDateFr_LostFocus()
    If IsDate(txtVisitDateFr) Then
    Else
        If txtVisitDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtVisitDateFr.SetFocus
        End If
    End If
End Sub

Private Sub txtVisitDateTo_LostFocus()
    If IsDate(txtVisitDateTo) Then
    Else
        If txtVisitDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtVisitDateTo.SetFocus
        End If
    End If
End Sub

Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim sql As String
Dim GrpCompanyName As String
    strAppend = ""
    GrpCompanyName = ""
    If Len(TxtWSNo) Then
        strAppend = strAppend & " AND CASNumber >= '" & Trim(TxtWSNo) & "'"
    End If

    If Len(TxtWSNoTo) Then
        strAppend = strAppend & " AND CASNumber <= '" & Trim(TxtWSNoTo) & "'"
    End If
    
    If Len(Trim(CboCLN)) Then
        strAppend = strAppend & " AND ClinicCode = '" & Trim(Left(CboCLN, IIf(InStr(1, CboCLN, "-") - 1 < 0, Len(CboCLN), InStr(1, CboCLN, "-") - 1))) & "'"
    End If
    
    If Len(Trim(CboPayor)) Then
         GrpCompanyName = Replace(CboPayor, "'", "''")
         strAppend = strAppend & " AND MBMGroupCompany = '" & Trim(GrpCompanyName) & "'"
    End If
    If Len(TxtInvNo) Then
        strAppend = strAppend & " and BillNo = '" & Trim(TxtInvNo) & "'"
    End If
    If (txtVisitDateFr) <> "__/__/____" And IsDate(txtVisitDateFr) = True _
        And (txtVisitDateTo) <> "__/__/____" And IsDate(txtVisitDateTo) = True Then
        sql = ""
        sql = " AND DateVisit >= '" & Format(Trim(txtVisitDateFr), "mm/dd/yyyy") & _
              "' And  DateVisit <= '" & Format(Trim(txtVisitDateTo), "mm/dd/yyyy") & "'"
        strAppend = strAppend & sql
    End If
    If Len(Trim(txtVisitDateFr)) > 0 And IsDate(txtVisitDateFr) = True Then
        strAppend = strAppend & " And DateVisit >= '" & Format(txtVisitDateFr, "mm/dd/yyyy") & "'"
    End If
    If Len(Trim(txtVisitDateTo)) > 0 And IsDate(txtVisitDateTo) = True Then
        strAppend = strAppend & " And DateVisit <= '" & Format(txtVisitDateTo, "mm/dd/yyyy") & "'"
    End If
    If ChkInvNo.Value = 1 Then
        sql = ""
        sql = "AND (BillNo Is null OR BillNo = '') " & ""
        strAppend = strAppend & sql
    End If
    If (txtInvDateFr) <> "__/__/____" And IsDate(txtInvDateFr) = True _
        And (txtInvDateTo) <> "__/__/____" And IsDate(txtInvDateTo) = True Then
        sql = ""
        sql = " AND BillDate >= '" & Format(Trim(txtInvDateFr), "mm/dd/yyyy") & _
              "' And  BillDate <= '" & Format(Trim(txtInvDateTo), "mm/dd/yyyy") & "'"
        strAppend = strAppend & sql
    End If
    If Len(Trim(txtInvDateFr)) > 0 And IsDate(txtInvDateFr) = True Then
        strAppend = strAppend & " And BillDate >= '" & Format(txtInvDateFr, "mm/dd/yyyy") & "'"
    End If
    If Len(Trim(txtInvDateTo)) > 0 And IsDate(txtInvDateTo) = True Then
        strAppend = strAppend & " And BillDate <= '" & Format(txtInvDateTo, "mm/dd/yyyy") & "'"
    End If
    If Len(txtInvAmtfr) Then
        strAppend = strAppend & " and INVAmount >= '" & Trim(txtInvAmtfr) & "'"
    End If
    If Len(TxtInvAmtTo) Then
        strAppend = strAppend & " and INVAmount <= '" & Trim(TxtInvAmtTo) & "'"
    End If
    Call InvoiceListing(strAppend)
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Sub InvoiceListing(Optional strAppend As String)
Const METHOD_NAME = "InvoiceListing"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim Total As String
Dim Current As String
Dim TotAmt As Variant
Dim TotalInvNo As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
    Screen.MousePointer = vbHourglass

    Current = 0
    TotAmt = 0
    lstInvoiceListing.ListItems.Clear
    strAppendCase = "1"
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimBordxModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
            With rst2
                Set Record = lstInvoiceListing.ListItems.Add(, , "")
                If Len(rst2!CASNumber) Then
                    Record.SubItems(1) = Trim(rst2!CASNumber)
                End If
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(2) = Trim(rst2!MBMNumber)
                End If
                If Len(rst2!coverid) Then
                    Record.SubItems(3) = Trim(rst2!coverid)
                End If
                If Len(rst2!PatientName) Then
                    Record.SubItems(4) = Trim(rst2!PatientName)
                End If
                If Len(rst2!DateVisit) Then
                    Record.SubItems(5) = Format(rst2!DateVisit, "dd/mm/yyyy")
                End If
                                                
                If Len(rst2!CLNName) Then
                    Record.SubItems(6) = Trim(rst2!CLNName)
                End If
                
                If Len(rst2!BillNo) Then
                    Record.SubItems(7) = Trim(rst2!BillNo)
                End If
                If Len(rst2!BillDate) Then
                   Record.SubItems(8) = Format(rst2!BillDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!INVEnteredDate) Then
                    Record.SubItems(9) = Format(rst2!INVEnteredDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!INVEnteredBy) Then
                    Record.SubItems(10) = Trim(rst2!INVEnteredBy)
                End If
                
                If Len(rst2!INVAmount) Then
                    Record.SubItems(11) = Format(rst2!INVAmount, "#,##0.00")
                Else
                    Record.SubItems(11) = "0.00"
                End If
                                
                If Len(rst2!CLMMCOProcFee) Then
                    Record.SubItems(12) = Format(rst2!CLMMCOProcFee, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(12) = "0.00"
                End If

                If Len(rst2!CLMFinalApprove) Then
                    Record.SubItems(13) = Format(rst2!CLMFinalApprove, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(13) = "0.00"
                End If
                                
                If Len(rst2!MBMGroupCompany) Then
                    Record.SubItems(14) = Trim(rst2!MBMGroupCompany)
                End If
                                                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!INVAmount) Then
                    TotAmt = TotAmt + rst2!INVAmount
                End If
            End With
            rst2.MoveNext
        Loop
    
    If lstInvoiceListing.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdSearch.Enabled = True
    End If
    
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub UpdateRecord(tmpCASnumber As String)
Const METHOD_NAME = "UpdateRecord"
On Error GoTo debugs
Dim Update
Dim strsql As String
Dim INV As New ADODB.Recordset

    strsql = " Select CASNumber, * from CISCas " & _
                " Where CASNumber = '" & Trim(tmpCASnumber) & "'"
    INV.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    With INV
        !InvoiceNo = TxtInvoiceNo
        !InvoiceDate = Format(txtInvoiceDate, "dd/mm/yyyy")
        !RegisterDate = Date
        !RegisterBy = Logon
    End With
    INV.Update
    
    INV.Close
    Set INV = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

