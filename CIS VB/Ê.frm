VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdateInvoiceBatch 
   Caption         =   "Update Invoice Batch"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame FramePV 
      Height          =   615
      Left            =   60
      TabIndex        =   34
      Top             =   6960
      Width           =   11775
      Begin VB.TextBox txtinvclncode 
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
         Left            =   6960
         MaxLength       =   4
         TabIndex        =   80
         Top             =   200
         Width           =   495
      End
      Begin VB.TextBox TxtInvoiceNo 
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
         Left            =   4080
         MaxLength       =   46
         TabIndex        =   13
         Top             =   200
         Width           =   2535
      End
      Begin VB.TextBox TxtOPBatchNo 
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
         MaxLength       =   30
         TabIndex        =   35
         Top             =   200
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtInvoiceDate 
         Height          =   315
         Left            =   8640
         TabIndex        =   14
         Top             =   195
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label lblbw 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   81
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "Invoice No"
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
         Left            =   3120
         TabIndex        =   50
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label5 
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
         Left            =   7680
         TabIndex        =   49
         Top             =   195
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "OP Batch No"
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
         Left            =   120
         TabIndex        =   36
         Top             =   200
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   60
      TabIndex        =   17
      Top             =   360
      Width           =   11775
      Begin VB.ComboBox cboEntityType 
         Height          =   315
         ItemData        =   "Ê.frx":0000
         Left            =   1560
         List            =   "Ê.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   2270
         Width           =   2055
      End
      Begin VB.ComboBox cboreferal 
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
         Left            =   7560
         Sorted          =   -1  'True
         TabIndex        =   77
         Top             =   2145
         Width           =   1335
      End
      Begin VB.ComboBox CboPayTo 
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
         Left            =   1560
         TabIndex        =   74
         Top             =   1960
         Width           =   2055
      End
      Begin VB.ComboBox cmbInvoice 
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
         ItemData        =   "Ê.frx":002B
         Left            =   7560
         List            =   "Ê.frx":0035
         TabIndex        =   69
         Top             =   1870
         Width           =   1335
      End
      Begin VB.CheckBox ChkBatchNo 
         Height          =   255
         Left            =   10920
         TabIndex        =   64
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox ChkBordxNo 
         Height          =   255
         Left            =   10920
         TabIndex        =   59
         Top             =   1590
         Width           =   255
      End
      Begin VB.TextBox TxtWSNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
            Size            =   8.25
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   20
         Top             =   1120
         Width           =   255
      End
      Begin VB.CheckBox ChkClinic 
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
         Left            =   5280
         TabIndex        =   19
         Top             =   550
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpComp 
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
         Left            =   5280
         TabIndex        =   18
         Top             =   840
         Width           =   375
      End
      Begin VB.ComboBox CboCLN 
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   3615
      End
      Begin VB.ComboBox CboPayor 
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   825
         Width           =   3615
      End
      Begin VB.TextBox TxtInvNo 
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
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1115
         Width           =   3615
      End
      Begin MSMask.MaskEdBox txtVisitDateFr 
         Height          =   315
         Left            =   7560
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox txtVisitDateTo 
         Height          =   315
         Left            =   9480
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox txtInvDateFr 
         Height          =   315
         Left            =   7560
         TabIndex        =   7
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox txtInvDateTo 
         Height          =   315
         Left            =   9480
         TabIndex        =   8
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
      Begin VB.TextBox txtInvAmtfr 
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
         Left            =   7560
         MaxLength       =   8
         TabIndex        =   9
         Top             =   760
         Width           =   1320
      End
      Begin VB.TextBox TxtInvAmtTo 
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
         Left            =   9480
         MaxLength       =   8
         TabIndex        =   10
         Top             =   760
         Width           =   1335
      End
      Begin VB.TextBox txtOPBatchNoFr 
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
         Left            =   7560
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1020
         Width           =   1320
      End
      Begin VB.TextBox txtOPBatchNoTo 
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
         Left            =   9480
         MaxLength       =   15
         TabIndex        =   12
         Top             =   1020
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtRegDateFr 
         Height          =   300
         Left            =   7560
         TabIndex        =   53
         Top             =   1275
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRegDateTo 
         Height          =   300
         Left            =   9480
         TabIndex        =   54
         Top             =   1280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboOMNIStatus 
         Height          =   315
         ItemData        =   "Ê.frx":004A
         Left            =   1560
         List            =   "Ê.frx":005A
         TabIndex        =   55
         Top             =   1370
         Width           =   3615
      End
      Begin VB.TextBox TxtBordx1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   7560
         MaxLength       =   20
         TabIndex        =   60
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TxtBordx2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   9480
         MaxLength       =   20
         TabIndex        =   61
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox CboClinicType 
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
         Left            =   1560
         TabIndex        =   65
         Top             =   1650
         Width           =   1455
      End
      Begin VB.ComboBox cboInfoFFS 
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
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   67
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label lblEntityType 
         Caption         =   "Entity Type"
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
         Left            =   120
         TabIndex        =   78
         ToolTipText     =   "Date Pass On"
         Top             =   2270
         Width           =   975
      End
      Begin VB.Label lblreferletter 
         Caption         =   "Referal Letter"
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
         Left            =   6240
         TabIndex        =   76
         Top             =   2145
         Width           =   1095
      End
      Begin VB.Label lblcashtype 
         Caption         =   "Clinic Type"
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
         TabIndex        =   75
         Top             =   1960
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Invoice Updated"
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
         Left            =   6240
         TabIndex        =   70
         Top             =   1870
         Width           =   1335
      End
      Begin VB.Label txtFFSStatus 
         Caption         =   "FFS Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   68
         Top             =   1725
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Clinic Type"
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
         Left            =   120
         TabIndex        =   66
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label47 
         Caption         =   "To"
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
         Left            =   9120
         TabIndex        =   63
         Top             =   1620
         Width           =   255
      End
      Begin VB.Label Label27 
         Caption         =   "Bordx No"
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
         Left            =   6240
         TabIndex        =   62
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "DOR"
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
         Left            =   6240
         TabIndex        =   58
         ToolTipText     =   "Date of Admission"
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Index           =   2
         Left            =   9120
         TabIndex        =   57
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "OMNI"
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
         TabIndex        =   56
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "To"
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
         Left            =   9120
         TabIndex        =   46
         Top             =   1050
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "OP Batch No"
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
         Left            =   6240
         TabIndex        =   45
         Top             =   1050
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "To"
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
         Left            =   9120
         TabIndex        =   33
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Visit Date"
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
         Left            =   6240
         TabIndex        =   32
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Claim Number"
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
         Left            =   120
         TabIndex        =   31
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   30
         Top             =   285
         Width           =   375
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   5580
         TabIndex        =   29
         Top             =   1170
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "To"
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
         Left            =   9120
         TabIndex        =   28
         Top             =   510
         Width           =   495
      End
      Begin VB.Label Label20 
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
         Index           =   0
         Left            =   6240
         TabIndex        =   27
         Top             =   510
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "To"
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
         Left            =   9120
         TabIndex        =   26
         Top             =   775
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Invoice Amount"
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
         Left            =   6240
         TabIndex        =   25
         Top             =   780
         Width           =   1575
      End
      Begin VB.Label Label26 
         Caption         =   "Invoice No"
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
         TabIndex        =   24
         Top             =   1170
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Group Company"
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
         TabIndex        =   23
         Top             =   885
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Clinic"
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
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label36 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5320
         TabIndex        =   21
         Top             =   440
         Width           =   255
      End
   End
   Begin VB.CommandButton CmdCheck 
      Caption         =   "&Check All"
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
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   855
   End
   Begin MSComctlLib.ListView lstInvoiceListing 
      Height          =   3135
      Left            =   0
      TabIndex        =   37
      Top             =   3840
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5530
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   34
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Case No"
         Object.Width           =   1940
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
         Object.Width           =   1940
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
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "OP Batch No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "PV Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "PV Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Cheque Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Reason 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Reason 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Reason 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Reason 4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Reason 5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Other Reason"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "MC Days"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "CoPay"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "Diff Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "Department"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "ICNumber"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "Update Invoice"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "CashType"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "OP Batch Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "Last Updated Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   60
      TabIndex        =   38
      Top             =   7560
      Width           =   11775
      Begin VB.TextBox lblTotalFFS 
         Height          =   285
         Left            =   4320
         TabIndex        =   73
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox lblTotalInvAmt 
         Height          =   285
         Left            =   1800
         TabIndex        =   72
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   71
         Top             =   210
         Width           =   855
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
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
         Left            =   8880
         TabIndex        =   52
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print to &File"
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
         Left            =   9480
         TabIndex        =   42
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
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
         Left            =   7440
         TabIndex        =   15
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
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
         Left            =   11040
         TabIndex        =   41
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
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
         Left            =   10440
         TabIndex        =   40
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
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
         Height          =   330
         Left            =   8160
         TabIndex        =   39
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Total FFS Amt"
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
         Left            =   5640
         TabIndex        =   48
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Total Inv Amt"
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
         Left            =   3120
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
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
         Left            =   1080
         TabIndex        =   43
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label16 
      Caption         =   "*Double-Click to see claim details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   51
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   " Update Invoice Batch"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmUpdateInvoiceBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpCASnumber As String
Dim strAppend As String
Dim OPBatchNo As String
Private m_blnDirection(1 To 34) As Boolean
Dim intExit As Integer
Dim CompTrans As Boolean
Dim tmpupdateOP As Boolean
Dim tmpMsno As Integer
Dim tmpVdate As String
Dim tmpno As String

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
Dim CLNRec As New ADODB.Recordset
    CboCLN.Clear
    If ChkClinic.Value = 1 Then
        CLNRec.Open "Select CLNCode, CLNName from CLN WHERE CLNPanel = 'Y' OR CLNPanel = 'C'", medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do Until CLNRec.EOF
            If Len(CLNRec!CLNName) Then
                CboCLN.AddItem CLNRec!CLNCode & " - " & Trim(CLNRec!CLNName)
            End If
            CLNRec.MoveNext
        Loop
        CLNRec.Close
        Set CLNRec = Nothing
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
    CboPayor.Clear
    If ChkGrpComp.Value = 1 Then
        Company.Open "Select GRPCoName from GRPCompany order by GRPCoName", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do Until Company.EOF
            CboPayor.AddItem Trim(Company!GrpCoName)
            Company.MoveNext
        Loop
        Company.Close
        Set Company = Nothing
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdCheck_Click()
Const METHOD_NAME = "CmdCheck_Click"
On Error GoTo debugs
Dim LstCnt As Integer
Dim cnt As Integer
cnt = 0
    
    If lstInvoiceListing.ListItems.Count <> 0 Then
        LstCnt = lstInvoiceListing.ListItems.Count
        For LstCnt = 1 To LstCnt
            If Trim(lstInvoiceListing.SelectedItem.SubItems(15)) <> "" And Trim(lstInvoiceListing.SelectedItem.SubItems(7)) <> "" Then
                lstInvoiceListing.ListItems(LstCnt).Checked = False
                lstInvoiceListing.ListItems(LstCnt).Text = ""
                cnt = cnt + 1
            Else
                lstInvoiceListing.ListItems(LstCnt).Checked = True
                lstInvoiceListing.ListItems(LstCnt).Text = "Check"
            End If
        Next LstCnt
    Else
        MsgBox "Please select a record.", vbInformation
    End If
    
    If cnt >= 1 Then
        MsgBox "Selected Few Cases Already OP Batch And Invoice No generated..! ", vbCritical
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
End Sub

Private Sub cmdExit_Click()
    Unload Me
    tmpMsno = 0
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
    
    If lstInvoiceListing.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstInvoiceListing)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
        Call SearchRecord
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
End Sub

Private Sub cmdUpdate_Click()

Const METHOD_NAME = "CmdUpdate_Click"
On Error GoTo debugs

Dim tmpBordxNo As String
Dim i As Integer
Dim update
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
Dim tmpstr As String
Dim INvnoset As New ADODB.Recordset
Dim strsql As String
Dim KeyAscii As Integer
ItemChecked = False
    
    If Len(Trim(TxtInvoiceNo)) Then
        If lstInvoiceListing.ListItems.Count = 0 Then
            MsgBox "Please Search Data..!", vbInformation
            Exit Sub
        ElseIf (txtVisitDateFr = "__/__/____" Or txtVisitDateTo = "__/__/____") Then
            MsgBox "Please select Visit Date.", vbInformation
            txtVisitDateFr.SetFocus
            Exit Sub
        ElseIf Len(CboCLN) = False Then
            MsgBox "Above claims must belong to same clinic code.", vbInformation
            Exit Sub
        ElseIf tmpMsno > 1 Then
            MsgBox "Above claims must belong to same clinic code.", vbInformation
            Exit Sub
        ElseIf Len(TxtInvoiceNo) Then
            strsql = ""
            KeyAscii = 13
            txtInvoiceDate = "__/__/____"
            txtinvclncode = ""
            txtInvoiceDate = DateSerial(Year(tmpVdate), Month(tmpVdate) + 1, 0)
            txtinvclncode = Mid(CboCLN, 1, 4)
            tmpno = TxtInvoiceNo + "-" + txtinvclncode
            KeyAscii = 0
            If Len(tmpno) Then
                strsql = "Select InvoiceNo from CISDB..CISCas where InvoiceNo= '" & Trim(tmpno) & "'"
                INvnoset.Open strsql, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
                    If INvnoset.EOF = False Then
                        MsgBox "Given Invoice No is already used...!", vbInformation
                        txtInvoiceDate = "__/__/____"
                        txtinvclncode = ""
                        Exit Sub
                    End If
                INvnoset.Close
                Set INvnoset = Nothing
            End If
        End If
    End If
    
    
    
    
    For i = 1 To lstInvoiceListing.ListItems.Count
        If lstInvoiceListing.ListItems(i).Checked = True Then
            ItemChecked = True
            Exit For
        End If
    Next i
    CLNCnt = 0
    TxtOPBatchNo = ""
    If ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    Else
        update = MsgBox("Do you want to update the record ? ", vbYesNo + vbExclamation)
        If update = vbYes Then
            For i = 1 To lstInvoiceListing.ListItems.Count
                tmpupdateOP = False
                If lstInvoiceListing.ListItems(i).Checked = True Then
                    If TxtOPBatchNo = "" Then
                        If lstInvoiceListing.ListItems(i).SubItems(15) = "" Then
                            OPBatchNo = GenerateOPSeq("OC")
                            TxtOPBatchNo = OPBatchNo
                        End If
                    End If
                    If lstInvoiceListing.ListItems(i).SubItems(15) = "" Then
                        tmpupdateOP = True
                    Else
                        tmpupdateOP = False
                    End If
                    tmpCASnumber = Trim(lstInvoiceListing.ListItems(i).SubItems(1))
                    Call UpdateRecord(tmpCASnumber)
                End If
            Next i
            If CompTrans = True Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                lstInvoiceListing.ListItems.Clear
                txtInvoiceDate = "__/__/____"
                TxtInvoiceNo = ""
                txtinvclncode = ""
            Else
                CompTrans = True
            End If
        Else
            Exit Sub
        End If
        
    'ErrorSave:
        'Screen.MousePointer = vbDefault
        'If CompTrans = False Then
        '    MediConnCISDB.RollbackTrans
        '    MsgBox Err.Description
        'End If
    End If
    
debugs:
Screen.MousePointer = vbDefault
   If CompTrans = False Then
       MediConnCISDB.RollbackTrans
       MsgBox Err.Description
   End If
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub Form_Load()
    intExit = 0
    cboOMNIStatus.Clear
    cboOMNIStatus.AddItem "Y" & " - " & "Yes"

    CboClinicType.AddItem "0 - Specialised"
    CboClinicType.AddItem "1 - General Practisional"
    CboClinicType.AddItem "2 - Others"
    
    CboPayTo.AddItem "C-Cash Less"
    CboPayTo.AddItem "M-Reimbursement"
    
    cboreferal.AddItem "YES"
    cboreferal.AddItem "NO"
    cboEntityType.ListIndex = 0
End Sub

Private Sub lstInvoiceListing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)

    Select Case ColumnHeader.Index
    Case 1
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
    Case 11
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    Case 13
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
    Case 14
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
    Case 15
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(18)
    Case 19
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(19)
    Case 20
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(21)
    Case 22
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(22)
    Case 23
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(23)
    Case 24
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(24)
    Case 25
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(25)
        
    Case 26
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
    Case 27
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
    Case 28
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
    Case 29
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(29)
    Case 30
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(30)
    Case 31
        SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(31)
    Case 32
       SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(32)
    Case 33
       SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(33)
    Case 34
       SortListView lstInvoiceListing, ColumnHeader.Index, ldtString, m_blnDirection(34)
    End Select
End Sub
Private Sub lstInvoiceListing_DblClick()
Const METHOD_NAME = "lstInvoiceListing_DblClick"
On Error GoTo debugs
If lstInvoiceListing.ListItems.Count > 0 Then
    frmInvoiceAdj.txtCaseNo = Trim(lstInvoiceListing.SelectedItem.SubItems(1))
    Call frmInvoiceAdj.CmdGo_Click
    Load frmInvoiceAdj
    frmInvoiceAdj.Show
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstInvoiceListing_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Trim(lstInvoiceListing.SelectedItem.SubItems(15)) <> "" And Trim(lstInvoiceListing.SelectedItem.SubItems(7)) <> "" Then
        Item.Text = ""
        Item.Checked = False
        MsgBox "Selected Case Already OP Batch And Invoice No generated..! ", vbCritical
    ElseIf Item.Checked = True Then
        Item.Text = "Check"
        Item.Checked = True
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

Private Sub TxtInvoiceNo_LostFocus()
Dim strsql As String
Dim INvnoset As New ADODB.Recordset
Dim KeyAscii As Integer

    If Len(Trim(TxtInvoiceNo)) Then
        If lstInvoiceListing.ListItems.Count = 0 Then
            MsgBox "Please Search Data..!", vbInformation
            Exit Sub
        ElseIf (txtVisitDateFr = "__/__/____" Or txtVisitDateTo = "__/__/____") Then
            MsgBox "Please select Visit Date.", vbInformation
            txtVisitDateFr.SetFocus
            Exit Sub
        ElseIf Len(CboCLN) = False Then
            MsgBox "Above claims must belong to same clinic code.", vbInformation
            Exit Sub
        ElseIf tmpMsno > 1 Then
            MsgBox "Above claims must belong to same clinic code.", vbInformation
            Exit Sub
        ElseIf Len(TxtInvoiceNo) Then
            KeyAscii = 13
            txtInvoiceDate = "__/__/____"
            txtinvclncode = ""
            txtInvoiceDate = DateSerial(Year(tmpVdate), Month(tmpVdate) + 1, 0)
            txtinvclncode = Mid(CboCLN, 1, 4)
            tmpno = Trim(TxtInvoiceNo) + "-" + Trim(txtinvclncode)
            KeyAscii = 0
            If Len(tmpno) Then
                strsql = "Select InvoiceNo from CISDB..CISCas where InvoiceNo= '" & Trim(tmpno) & "'"
                INvnoset.Open strsql, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
                    If INvnoset.EOF = False Then
                        MsgBox "Given Invoice No is already used...!", vbInformation
                        TxtInvoiceNo.SetFocus
                        txtInvoiceDate = "__/__/____"
                        txtinvclncode = ""
                        Exit Sub
                    Else
                        Call cmdUpdate_Click
                    End If
                INvnoset.Close
                Set INvnoset = Nothing
            End If
        End If
    Else
        KeyAscii = 13
        txtInvoiceDate = "__/__/____"
        txtinvclncode = ""
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
    txtinvclncode = ""
    TxtInvoiceNo = ""
    txtInvoiceDate = "__/__/____"
    strAppend = ""
    GrpCompanyName = ""
    TxtOPBatchNo = ""
    strAppend = strAppend & " And CASStatus <> 'D' And (Substring(MBMNumber,1,2) <> 'EM') AND (Substring(MBMNumber,1,2) <> 'ES')"
    If ChkBatchNo.Value = 1 Then
        strAppend = strAppend & " AND (OPBatchNo Is null OR OPBatchNo = '') " & ""
    End If
        
    If ChkBordxNo.Value = 1 Then
        strAppend = strAppend & " AND (BordxRefNo Is null OR BordxRefNo = '')" & ""
    End If
        
    If Len(Trim(TxtBordx1)) Then
        strAppend = strAppend + " AND BordxRefNo >= '" & Trim(TxtBordx1) & "'"
    End If
    
    If Len(Trim(TxtBordx2)) Then
        strAppend = strAppend + " AND BordxRefNo <= '" & Trim(TxtBordx2) & "'"
    End If
        
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
    
    If Len(txtOPBatchNoFr) Then
        strAppend = strAppend & " and OPBatchNo >= '" & Trim(txtOPBatchNoFr) & "'"
    End If
        
    If Len(txtOPBatchNoTo) Then
        strAppend = strAppend & " and OPBatchNo <= '" & Trim(txtOPBatchNoTo) & "'"
    End If
    
    If Len(Trim(cboOMNIStatus)) Then
        strAppend = strAppend + " AND GRPOMNICLM = 'Y'"
    End If
    
    If Len(Trim(txtRegDateFr)) > 0 And IsDate(txtRegDateFr) = True Then
        strAppend = strAppend + " And RegisterDate >= '" & Format(txtRegDateFr, "yyyy/mm/dd") & "'"
    End If
    
    If Len(Trim(txtRegDateTo)) > 0 And IsDate(txtRegDateTo) = True Then
        strAppend = strAppend + " And RegisterDate <= '" & Format(txtRegDateTo, "yyyy/mm/dd") & "'"
    End If
    
    If Mid(CboClinicType, 1, 1) = 0 Then
        strAppend = strAppend & " AND  CLNType = 'S'"
    ElseIf Mid(CboClinicType, 1, 1) = 1 Then
        strAppend = strAppend & " AND  CLNType = 'G'"
    ElseIf Mid(CboClinicType, 1, 1) = 2 Then
        strAppend = strAppend & " AND ( CLNType = '' OR  CLNType IS NULL)"
    End If
    
    If Mid(cboInfoFFS, 1, 1) = "Y" Then
        strAppend = strAppend & " AND CLNFFSStatus = 'Y'"
    ElseIf Mid(cboInfoFFS, 1, 1) = "N" Then
        strAppend = strAppend & " AND CLNFFSStatus = 'N' OR (CLNFFSStatus = '' OR CLNFFSStatus IS NULL)"
    End If
    
    If Mid(cmbInvoice, 1, 1) = "Y" Then
        strAppend = strAppend & " AND (OPBatchNo Is Not Null) " & ""
    ElseIf Mid(cmbInvoice, 1, 1) = "N" Then
        strAppend = strAppend & " AND (OPBatchNo Is null OR OPBatchNo = '') " & ""
    End If
    
    If Len(Trim(CboPayTo)) Then
        strAppend = strAppend + " AND INVPayTo = '" & Trim(Mid(CboPayTo, 1, IIf(InStr(1, CboPayTo, "-") = 0, 4, InStr(1, CboPayTo, "-") - 1))) & "'"
    End If
    
    If Trim(cboreferal) = "YES" Then
    strAppend = strAppend + " AND CLMRefLetter='Y'"
    ElseIf Trim(cboreferal) = "NO" Then
    strAppend = strAppend + " AND (CLMRefLetter='N' OR CLMRefLetter='' or CLMRefLetter IS NULL) "
    End If
    
    If cboEntityType.ListIndex = 0 Then
        strAppend = strAppend + " AND Substring(CasNumber, 1, 2) <>'SJ' "
    ElseIf cboEntityType.ListIndex = 1 Then
        strAppend = strAppend + " AND Substring(CasNumber, 1, 2) ='SJ' "
    End If
    
    If Len(strAppend) Then
        strAppend = strAppend + " Order by MSNO Asc  "
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
Dim TotFFS As String
Dim TmpPending As String
Dim sqlpymt As String
Dim rst3 As New ADODB.Recordset
Dim Check As Boolean
Dim strsqlDIA As String
Dim DIA As New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    
    Current = 0
    TotAmt = 0
    TotFFS = 0
    tmpMsno = 0
    
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
    Do
    
        Do Until rst2.EOF
            If rst2!Pending <> "Y" Then
            With rst2
                Set Record = lstInvoiceListing.ListItems.Add(, , "")
                If Len(rst2!CasNumber) Then
                    Record.SubItems(1) = Trim(rst2!CasNumber)
                End If
                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(2) = Trim(rst2!MBMNumber)
                End If
                
                If Len(rst2!coverid) Then
                    Record.SubItems(3) = Trim(rst2!coverid)
                End If
                
                Record.SubItems(4) = Trim(rst2!PatientName)
                
                If Len(rst2!DateVisit) Then
                    Record.SubItems(5) = Format(rst2!DateVisit, "dd/mm/yyyy")
                    tmpVdate = Format(rst2!DateVisit, "dd/mm/yyyy")
                End If
                
                If !MSNO = 1 Then
                    tmpMsno = tmpMsno + 1
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
                
                If Len(rst2!OPBatchNo) Then
                    Record.SubItems(15) = Trim(rst2!OPBatchNo)
                End If
                
                If rst2!INVPayTo = "C" Then
                        If Len(rst2!PVCLNNo) Then
                            Record.SubItems(16) = Trim(rst2!PVCLNNo)
                        End If
                        
                        If Len(rst2!PVCLNDate) Then
                            Record.SubItems(17) = Trim(rst2!PVCLNDate)
                        End If
                                             
                        If Len(rst2!PVCLNChequeNo) Then
                            Record.SubItems(18) = Trim(rst2!PVCLNChequeNo)
                        End If
                ElseIf rst2!INVPayTo = "M" Then
                
                    sqlpymt = " Select Top 1 * From Payment2MBM where CASNumber = '" & Trim(rst2!CasNumber) & "'"
                    rst3.Open sqlpymt, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
                    Do While Not rst3.EOF
                        If Len(rst3!PVNo) Then
                            Record.SubItems(16) = Trim(rst3!PVNo)
                        End If
                        
                        If Len(rst3!PVDate) Then
                            Record.SubItems(17) = Trim(rst3!PVDate)
                        End If
                                             
                        If Len(rst3!PVChequeNo) Then
                            Record.SubItems(18) = Trim(rst3!PVChequeNo)
                        End If
                    rst3.MoveNext
                    Loop
                    rst3.Close
                    Set rst3 = Nothing
                End If
                
                Record.SubItems(19) = IIf(Len(rst2!ReasonDesc1), rst2!ReasonDesc1, "")
                Record.SubItems(20) = IIf(Len(rst2!ReasonDesc2), rst2!ReasonDesc2, "")
                Record.SubItems(21) = IIf(Len(rst2!ReasonDesc3), rst2!ReasonDesc3, "")
                Record.SubItems(22) = IIf(Len(rst2!ReasonDesc4), rst2!ReasonDesc4, "")
                Record.SubItems(23) = IIf(Len(rst2!ReasonDesc5), rst2!ReasonDesc5, "")
                If Len(rst2!CASDiagnosis5) Then
                    Record.SubItems(24) = Trim(rst2!CASDiagnosis5)
                End If
                
                If Len(rst2!MCDays) Then
                    Record.SubItems(25) = Trim(rst2!MCDays)
                End If
                
                If Len(rst2!CLMCoPayCLNColl) Then
                    Record.SubItems(26) = Trim(rst2!CLMCoPayCLNColl)
                End If
                If IsNumeric(rst2!CLMFinalApprove) Then
                    Record.SubItems(27) = CDbl(rst2!CLMFinalApprove) - CDbl(rst2!INVAmount)
                Else
                    If IsNumeric(rst2!INVAmount) Then
                        Record.SubItems(27) = CDbl(rst2!INVAmount)
                    Else
                        Record.SubItems(27) = 0
                    End If
                End If
                
                Record.SubItems(28) = IIf(Len(rst2!MBMDepartment), rst2!MBMDepartment, "")
                
                 If Len(rst2!coverid) Then
                   If (rst2!coverid) = "00" Then
                        Record.SubItems(29) = IIf(Len(rst2!MBMIcBcPp), rst2!MBMIcBcPp, "")
                   Else
                        Record.SubItems(29) = IIf(Len(rst2!MBMCICBCPPNo), rst2!MBMCICBCPPNo, "")
                   End If
                 End If
                
                    If (Len(rst2!OPBatchNo)) Then
                         Record.SubItems(30) = "YES"
                    Else
                         Record.SubItems(30) = "NO"
                    End If
                    
                    If (Len(rst2!INVPayTo)) Then
                     If (rst2!INVPayTo) = "C" Then
                        Record.SubItems(31) = "Cash Less"
                   Else
                        Record.SubItems(31) = "Reimbursement"
                   End If
                    End If
                    
                    If Len(rst2!OPBatchDate) Then
                    Record.SubItems(32) = Format(rst2!OPBatchDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!LastUpdateDate) Then
                    Record.SubItems(33) = !LastUpdateDate
                    End If
                    
                    'Record.SubItems(33) = IIf(Len(rst2!OPBatchtime), rst2!OPBatchtime, "")
                                        
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!INVAmount) Then
                    TotAmt = TotAmt + rst2!INVAmount
                End If
                lblTotalInvAmt = Format(TotAmt, "#,##0.00")
                If Len(rst2!CLMMCOProcFee) Then
                    TotFFS = TotFFS + rst2!CLMMCOProcFee
                End If
                lblTotalFFS = Format(TotFFS, "#,##0.00")
            End With
            Else
                TmpPending = rst2!Pending
            End If
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
    'Ravi
    If Mid(TmpPending, 1, 1) = "Y" Then
         MsgBox "Some of the Cases You are trying to Listout are in pending state. Please Check!", vbInformation
    End If
    
    If lstInvoiceListing.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdSearch.Enabled = True
    End If
    
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
Dim update
Dim INV As New ADODB.Recordset
Dim strsql As String
    
    CompTrans = False
    MediConnCISDB.beginTrans
    strsql = ""
    If txtInvoiceDate <> "__/__/____" Then
        If tmpupdateOP = True Then
            strsql = "update CISDB..CISCas set OPBatchNo = '" & Trim(TxtOPBatchNo) & "', OPNoEnterUser =  '" & Logon & "',InvoiceNo = '" & tmpno & "', InvoiceDate = '" & Format(txtInvoiceDate, "YYYY/MM/DD") & "' , OPBatchDate = '" & Format(DateTime.Now, "YYYY/MM/DD") & "', CASInvMonth = '" & Format(Month(txtInvoiceDate), "00") & "', CASInvYear = '" & Year(txtInvoiceDate) & "'" & _
                    " where CASNumber ='" & Trim(tmpCASnumber) & "'"
        Else
            strsql = "update CISDB..CISCas set InvoiceNo = '" & tmpno & "', InvoiceDate = '" & Format(txtInvoiceDate, "YYYY/MM/DD") & "', CASInvMonth = '" & Format(Month(txtInvoiceDate), "00") & "', CASInvYear = '" & Year(txtInvoiceDate) & "'" & _
                    " where CASNumber ='" & Trim(tmpCASnumber) & "'"
        End If
            MediConnCISDB.Execute strsql
            MediConnCISDB.commitTrans
            CompTrans = True
    ElseIf txtInvoiceDate = "__/__/____" Then
        If tmpupdateOP = True Then
            strsql = "update CISDB..CISCas set OPBatchNo = '" & Trim(TxtOPBatchNo) & "', OPNoEnterUser =  '" & Logon & "',  OPBatchDate = '" & Format(DateTime.Now, "YYYY/MM/DD") & "'" & _
                " where CASNumber ='" & Trim(tmpCASnumber) & "'"
            MediConnCISDB.Execute strsql
            MediConnCISDB.commitTrans
            CompTrans = True
        End If
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Function GenerateOPSeq(OPVar)
Const METHOD_NAME = "GenerateOPSeq"
On Error GoTo debugs

Dim OPRec As New ADODB.Recordset
Dim strsql As String
Dim SeqNo As String
Dim SEQ As New ADODB.Recordset
Dim SEQInsert As New ADODB.Recordset
Dim currentNo, i As Integer
Dim SEQNumber, SEQCode As String
Dim CharacterArray
Dim firstChar, currentChar As String
    SEQNumber = "0000001"
    CharacterArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    strsql = "Select OPCode,OPBatchSeq from OPBatch Where OPCode = '" & OPVar & "'"
    OPRec.Open strsql, MediConnCISDB, adOpenStatic, adLockOptimistic
    currentNo = 1
    currentChar = "A"
    
    If OPRec.RecordCount = 0 Then
        With OPRec
            .AddNew
                !OPCode = OPVar
                !OPBatchSeq = currentChar & "0000001"
                .update
                SEQNumber = currentChar & "0000001"
            .update
        End With
    Else
        currentNo = Mid(OPRec!OPBatchSeq, 2, 8)
        firstChar = Mid(OPRec!OPBatchSeq, 1, 1)
        Do Until i = 25
            If firstChar = CharacterArray(i) Then
                If currentNo <> "9999999" Then
                    currentChar = CharacterArray(i)
                    currentNo = currentNo + 1
                Else
                    currentChar = CharacterArray(i + 1)
                    currentNo = "0000001"
                End If
                Exit Do
            End If
            i = i + 1
        Loop
        SEQNumber = currentChar & Format(currentNo, "0000000")
        OPRec!OPBatchSeq = SEQNumber
        OPRec.update
    End If

    GenerateOPSeq = "OC" & SEQNumber
    OPRec.Close
    Set OPRec = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function




