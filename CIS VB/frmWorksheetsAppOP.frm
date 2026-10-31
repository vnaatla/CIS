VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmWorksheetsAppOP 
   Caption         =   "Worksheets for Approval - Outpatient"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   60
      TabIndex        =   57
      Top             =   7060
      Width           =   11775
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
         TabIndex        =   63
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
         TabIndex        =   62
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   58
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lblTotalFFS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Left            =   4440
         TabIndex        =   69
         Top             =   210
         Width           =   1125
      End
      Begin VB.Label lblTotalInvAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Left            =   1920
         TabIndex        =   68
         Top             =   210
         Width           =   1125
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         TabIndex        =   64
         Top             =   210
         Width           =   885
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
      TabIndex        =   55
      Top             =   2640
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   60
      TabIndex        =   7
      Top             =   360
      Width           =   11775
      Begin VB.CheckBox ChkBatchNo 
         Height          =   255
         Left            =   10920
         TabIndex        =   26
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox ChkBordxNo 
         Height          =   255
         Left            =   10920
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   1115
         Width           =   3615
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         MaxLength       =   8
         TabIndex        =   14
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
         MaxLength       =   8
         TabIndex        =   13
         Top             =   1020
         Width           =   1335
      End
      Begin VB.ComboBox cboOMNIStatus 
         Height          =   315
         ItemData        =   "frmWorksheetsAppOP.frx":0000
         Left            =   1560
         List            =   "frmWorksheetsAppOP.frx":0010
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   1650
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtVisitDateFr 
         Height          =   315
         Left            =   7560
         TabIndex        =   27
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
         TabIndex        =   28
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
         TabIndex        =   29
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
         TabIndex        =   30
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
      Begin MSMask.MaskEdBox txtRegDateFr 
         Height          =   300
         Left            =   7560
         TabIndex        =   31
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
         TabIndex        =   32
         Top             =   1280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
         ToolTipText     =   "Date of Admission"
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Index           =   2
         Left            =   9120
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   440
         Width           =   255
      End
   End
   Begin VB.Frame FramePV 
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   6480
      Width           =   11775
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
         Left            =   3960
         MaxLength       =   30
         TabIndex        =   2
         Top             =   200
         Width           =   1935
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
         TabIndex        =   1
         Top             =   200
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtInvoiceDate 
         Height          =   315
         Left            =   7080
         TabIndex        =   3
         Top             =   200
         Width           =   1215
         _ExtentX        =   2143
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
         TabIndex        =   6
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
         Left            =   6120
         TabIndex        =   5
         Top             =   200
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
         TabIndex        =   4
         Top             =   200
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstInvoiceListing 
      Height          =   3495
      Left            =   0
      TabIndex        =   56
      Top             =   3000
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6165
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
      NumItems        =   28
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
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Worksheets for Approval - Outpatient"
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
      TabIndex        =   71
      Top             =   0
      Width           =   11805
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
      TabIndex        =   70
      Top             =   2640
      Width           =   6015
   End
End
Attribute VB_Name = "frmWorksheetsAppOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpCASnumber As String
Dim strAppend As String
Dim OPBatchNo As String
Private m_blnDirection(1 To 28) As Boolean
Dim intExit As Integer
Dim CompTrans As Boolean
Dim tmpupdateOP As Boolean

Private Sub CboCLN_KeyPress(KeyAscii As Integer)
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
End Sub
Private Sub CboPayor_KeyPress(KeyAscii As Integer)
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

End Sub

Private Sub ChkClinic_Click()
Dim ClnRec As New ADODB.Recordset
    CboCLN.Clear
    If ChkClinic.Value = 1 Then
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        ClnRec.Open "Select CLNCode, CLNName from CLN WHERE CLNPanel = 'Y' OR CLNPanel = 'C'", medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do Until ClnRec.EOF
            If Len(ClnRec!CLNname) Then
                CboCLN.AddItem ClnRec!CLNCode & " - " & Trim(ClnRec!CLNname)
            End If
            ClnRec.MoveNext
        Loop
        ClnRec.Close
        Set ClnRec = Nothing
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        'MediConnCISWS.Close
        'Set MediConnCISWS = Nothing
    End If
End Sub

Private Sub ChkGrpComp_Click()
Dim Company As New ADODB.Recordset

    CboPayor.Clear
    If ChkGrpComp.Value = 1 Then
    
       ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
        Company.Open "Select GRPCoName from GRPCompany order by GRPCoName", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do Until Company.EOF
            CboPayor.AddItem Trim(Company!GrpCoName)
            Company.MoveNext
        Loop
        Company.Close
        Set Company = Nothing
       ' MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
       ' MediConnCISWS.Close
       ' Set MediConnCISWS = Nothing
        
    End If
End Sub

Private Sub CmdCheck_Click()
Dim LstCnt As Integer
    If lstInvoiceListing.ListItems.Count <> 0 Then
        LstCnt = lstInvoiceListing.ListItems.Count
    
        For LstCnt = 1 To LstCnt

            'If Len(lstInvoiceListing.ListItems(LstCnt).SubItems(15)) And Len(lstInvoiceListing.ListItems(LstCnt).SubItems(7)) Then
            '    MsgBox "This claim has been updated.", vbInformation
            '    lstInvoiceListing.ListItems(LstCnt).Checked = False
            'Else
                lstInvoiceListing.ListItems(LstCnt).Checked = True
                lstInvoiceListing.ListItems(LstCnt).Text = "Check"
            'End If
        Next LstCnt
    Else
        MsgBox "Please select a record.", vbInformation
    End If
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
End Sub

Private Sub cmdExit_Click()
    Unload Me
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
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
   ' MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"

        Call SearchRecord
        
   ' MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' MediConnCISWS.Close
   ' Set MediConnCISWS = Nothing
   ' MediConnCISPayment.Close
   ' Set MediConnCISPayment = Nothing
        
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
End Sub

Private Sub CmdUpdate_Click()
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
    ItemChecked = False
    
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
                            OPBatchNo = GenerateOPSeq("OP")
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
            
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            lstInvoiceListing.ListItems.Clear
            'tmpstr = " OPBatchNo = '" & Trim(TxtOPBatchNo) & "'"
            Call InvoiceListing(strAppend)
           '
          '  MediConnCISDB.Close
          '  Set MediConnCISDB = Nothing
          '  medixmasterconn.Close
          '  Set medixmasterconn = Nothing
          '  MediConnCISWS.Close
          '  Set MediConnCISWS = Nothing
          '  MediConnCISPayment.Close
          '  Set MediConnCISPayment = Nothing
        End If
ErrorSave:
        Screen.MousePointer = vbDefault
        If CompTrans = False Then
            MediConnCISDB.RollbackTrans
           ' MediConnCISDB.Close
           ' Set MediConnCISDB = Nothing
           ' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
           ' MediConnCISWS.Close
           ' Set MediConnCISWS = Nothing
           ' MediConnCISPayment.Close
           ' Set MediConnCISPayment = Nothing
            MsgBox Err.Description
        End If
    End If
 
End Sub

Private Sub Form_Load()
    intExit = 0
    cboOMNIStatus.Clear
    cboOMNIStatus.AddItem "Y" & " - " & "Yes"

    CboClinicType.AddItem "0 - Specialised"
    CboClinicType.AddItem "1 - General Practisional"
    CboClinicType.AddItem "2 - Others"
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
        
    End Select
    
        
End Sub


Private Sub lstInvoiceListing_DblClick()
    If lstInvoiceListing.ListItems.Count > 0 Then
        frmInvoiceAdj.txtCaseNo = Trim(lstInvoiceListing.SelectedItem.SubItems(1))
        Call frmInvoiceAdj.CmdGo_Click
        Load frmInvoiceAdj
        frmInvoiceAdj.Show
    End If

End Sub

Private Sub lstInvoiceListing_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = True Then
       ' If Len(Item.SubItems(15)) And Len(Item.SubItems(7)) Then
       '     MsgBox "This claim has been updated", vbCritical
       '     Item.Checked = False
       ' Else
            Item.Text = "Check"
            Item.Checked = True
        'End If
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
Dim sql As String
Dim GrpCompanyName As String
    
    
    strAppend = ""
    GrpCompanyName = ""
    TxtOPBatchNo = ""
    
    'HISDB.dbo.MBM.PAYCode <> 'EM') AND (HISDB.dbo.MBM.PAYCode <> 'ES
    strAppend = strAppend & " CISDB.dbo.CISCas.CASStatus <> 'D' And (Substring(HISDB.dbo.MBM.MBMNumber,1,2) <> 'EM') AND (Substring(HISDB.dbo.MBM.MBMNumber,1,2) <> 'ES')"
        
    If ChkBatchNo.Value = 1 Then
        strAppend = strAppend & " AND (OPBatchNo Is null OR OPBatchNo = '') " & ""
    End If
        
    If ChkBordxNo.Value = 1 Then
        strAppend = strAppend & " AND (dbo.Claim.BordxRefNo Is null OR dbo.Claim.BordxRefNo = '')" & ""
    End If
        
    If Len(Trim(TxtBordx1)) Then
        strAppend = strAppend + " AND dbo.Claim.BordxRefNo >= '" & Trim(TxtBordx1) & "'"
    End If
    
    If Len(Trim(TxtBordx2)) Then
        strAppend = strAppend + " AND dbo.Claim.BordxRefNo <= '" & Trim(TxtBordx2) & "'"
    End If
        
    If Len(TxtWSNo) Then
        strAppend = strAppend & " AND CISDB.dbo.CISCas.CASNumber >= '" & Trim(TxtWSNo) & "'"
    End If

    If Len(TxtWSNoTo) Then
        strAppend = strAppend & " AND CISDB.dbo.CISCas.CASNumber <= '" & Trim(TxtWSNoTo) & "'"
    End If
    
    If Len(Trim(CboCLN)) Then
        strAppend = strAppend & " AND ClinicCode = '" & Trim(Left(CboCLN, IIf(InStr(1, CboCLN, "-") - 1 < 0, Len(CboCLN), InStr(1, CboCLN, "-") - 1))) & "'"
    End If
    
    If Len(Trim(CboPayor)) Then
         GrpCompanyName = Replace(CboPayor, "'", "''")
         strAppend = strAppend & " AND MBMGroupCompany = '" & Trim(GrpCompanyName) & "'"
    End If
    
    
    If Len(TxtInvNo) Then
        strAppend = strAppend & " and InvoiceNo = '" & Trim(TxtInvNo) & "'"
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
        sql = "AND (InvoiceNo Is null OR InvoiceNo = '') " & ""
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
    

   ' strAppend = strAppend & " and INVPayTo <> 'M'"
    
    
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
        strAppend = strAppend & " AND  HISDB.dbo.CLN.CLNType = 'S'"
    ElseIf Mid(CboClinicType, 1, 1) = 1 Then
        strAppend = strAppend & " AND  HISDB.dbo.CLN.CLNType = 'G'"
    ElseIf Mid(CboClinicType, 1, 1) = 2 Then
        strAppend = strAppend & " AND ( HISDB.dbo.CLN.CLNType = '' OR  HISDB.dbo.CLN.CLNType IS NULL)"
    End If
    
    If Mid(cboInfoFFS, 1, 1) = "Y" Then
        strAppend = strAppend & " AND CLNFFSStatus = 'Y'"
    ElseIf Mid(cboInfoFFS, 1, 1) = "N" Then
        strAppend = strAppend & " AND CLNFFSStatus = 'N' OR (CLNFFSStatus = '' OR CLNFFSStatus IS NULL)"
    End If
    
    
    Call InvoiceListing(strAppend)

End Function

Private Sub InvoiceListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotAmt As Variant
Dim TotalInvNo As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String
Dim TotFFS As String
Dim sqlpymt As String
Dim rst3 As New ADODB.Recordset
Dim Check As Boolean
Dim strsqlDIA As String
Dim DIA As New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
    
    Current = 0
    TotAmt = 0
    TotFFS = 0
    'lblTotalAmt = ""
    lstInvoiceListing.ListItems.Clear
    
    
    strAppendCase = "1"
            
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimInvoiceTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    
    Do
        Do Until rst2.EOF
            
            'lblTotal = rst2.RecordCount
            
            With rst2
                
                Set Record = lstInvoiceListing.ListItems.add(, , "")
                
                
                If Len(rst2!CasNumber) Then
                    Record.SubItems(1) = Trim(rst2!CasNumber)
                End If
                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(2) = Trim(rst2!MBMNumber)
                End If
                
                If Len(rst2!CoverID) Then
                    Record.SubItems(3) = Trim(rst2!CoverID)
                End If
                                
                If Trim(rst2!CoverID) = "00" Then
                    Record.SubItems(4) = Trim(rst2!MBMName)
                Else
                    Record.SubItems(4) = IIf(Len(Trim(rst2!MBMCName)), Trim(rst2!MBMCName), "")
                End If
                
                If Len(rst2!DateVisit) Then
                    Record.SubItems(5) = Format(rst2!DateVisit, "dd/mm/yyyy")
                End If
                                                
                If Len(rst2!CLNname) Then
                    Record.SubItems(6) = Trim(rst2!CLNname)
                End If
                
                If Len(rst2!InvoiceNo) Then
                    Record.SubItems(7) = Trim(rst2!InvoiceNo)
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
                
                If rst2!InvPayTo = "C" Then
                    sqlpymt = " Select Top 1 * From Payment2CLN where CASNumber = '" & Trim(rst2!CasNumber) & "'"
                    rst3.Open sqlpymt, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
                    
                    Do While Not rst3.EOF
                        If Len(rst3!PVCLNNo) Then
                            Record.SubItems(16) = Trim(rst3!PVCLNNo)
                        End If
                        
                        If Len(rst3!PVCLNDate) Then
                            Record.SubItems(17) = Trim(rst3!PVCLNDate)
                        End If
                                             
                        If Len(rst3!PVCLNChequeNo) Then
                            Record.SubItems(18) = Trim(rst3!PVCLNChequeNo)
                        End If
                    rst3.MoveNext
                    Loop
                    rst3.Close
                    Set rst3 = Nothing
                    
                ElseIf rst2!InvPayTo = "M" Then
                
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
                
                If Len(rst2!CASTreatReason1) Then
                    strsqlDIA = " Select ReasonCode, ReasonDesc from TreatmentReason Where ReasonCode = '" & Trim(rst2!CASTreatReason1) & "'"
                    DIA.Open strsqlDIA, MediConnCISDB, adOpenKeyset, adLockOptimistic
                    If DIA.recordCount Then
                        If Len(DIA!ReasonDesc) Then
                            Record.SubItems(19) = Trim(DIA!ReasonDesc)
                        End If
                    End If
                    DIA.Close
                    Set DIA = Nothing
                End If
                If Len(rst2!CASTreatReason2) Then
                    strsqlDIA = " Select ReasonCode, ReasonDesc from TreatmentReason Where ReasonCode = '" & Trim(rst2!CASTreatReason2) & "'"
                    DIA.Open strsqlDIA, MediConnCISDB, adOpenKeyset, adLockOptimistic
                    If DIA.recordCount Then
                        If Len(DIA!ReasonDesc) Then
                            Record.SubItems(20) = Trim(DIA!ReasonDesc)
                        End If
                    End If
                    DIA.Close
                    Set DIA = Nothing
                End If
                If Len(rst2!CASTreatReason3) Then
                    strsqlDIA = " Select ReasonCode, ReasonDesc from TreatmentReason Where ReasonCode = '" & Trim(rst2!CASTreatReason3) & "'"
                    DIA.Open strsqlDIA, MediConnCISDB, adOpenKeyset, adLockOptimistic
                    If DIA.recordCount Then
                        If Len(DIA!ReasonDesc) Then
                            Record.SubItems(21) = Trim(DIA!ReasonDesc)
                        End If
                    End If
                    DIA.Close
                    Set DIA = Nothing
                End If
                If Len(rst2!CASTreatReason4) Then
                    strsqlDIA = " Select ReasonCode, ReasonDesc from TreatmentReason Where ReasonCode = '" & Trim(rst2!CASTreatReason4) & "'"
                    DIA.Open strsqlDIA, MediConnCISDB, adOpenKeyset, adLockOptimistic
                    If DIA.recordCount Then
                        If Len(DIA!ReasonDesc) Then
                            Record.SubItems(22) = Trim(DIA!ReasonDesc)
                        End If
                    End If
                    DIA.Close
                    Set DIA = Nothing
                End If
                If Len(rst2!CASTreatReason5) Then
                    strsqlDIA = " Select ReasonCode, ReasonDesc from TreatmentReason Where ReasonCode = '" & Trim(rst2!CASTreatReason5) & "'"
                    DIA.Open strsqlDIA, MediConnCISDB, adOpenKeyset, adLockOptimistic
                    If DIA.recordCount Then
                        If Len(DIA!ReasonDesc) Then
                            Record.SubItems(23) = Trim(DIA!ReasonDesc)
                        End If
                    End If
                    DIA.Close
                    Set DIA = Nothing
                End If
                
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
                
                'If IsNumeric(rst2!CLMMCOProcFee) Then
                '    Record.SubItems(28) = CDbl(rst2!CLMMCOProcFee)
                'Else
                '    Record.SubItems(28) = 0
                'End If
                
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
            rst2.MoveNext
            DoEvents
           
                If intExit = 1 Then
                    Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If lstInvoiceListing.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        'CmdPrintFile.Enabled = True
        CmdExit.Enabled = True
        CmdSearch.Enabled = True
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    'CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
    'MediConnCISPayment.Close
    'Set MediConnCISPayment = Nothing
    Screen.MousePointer = Default

End Sub

Private Sub UpdateRecord(tmpCASnumber As String)
Dim update
Dim StrSql As String
Dim INV As New ADODB.Recordset
    
   ' strsql = " Select CASNumber, OPBatchNo,InvoiceNo,InvoiceDate,OPNoEnterUser from CISCas " & _
                " Where CASNumber = '" & Trim(tmpCASnumber) & "'"
   ' INV.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
   ' With INV
   '     !OPBatchNo = TxtOPBatchNo
   '     !OPNoEnterUser = Logon
   '     If Len(TxtInvoiceNo) Then
   '         !InvoiceNo = TxtInvoiceNo
       ' End If
      '  If txtInvoiceDate <> "__/__/____" Then
     '       !InvoiceDate = Format(txtInvoiceDate, "dd/mm/yyyy")
    '    End If
    'End With
    
    'INV.Update
    
    'INV.Close
    'Set INV = Nothing
    
    CompTrans = False
    MediConnCISDB.beginTrans

    StrSql = ""
    If txtInvoiceDate <> "__/__/____" Then
        If tmpupdateOP = True Then
            StrSql = "update CISCas set OPBatchNo = '" & Trim(TxtOPBatchNo) & "', OPNoEnterUser =  '" & Logon & "',InvoiceNo = '" & TxtInvoiceNo & "', InvoiceDate = '" & Format(txtInvoiceDate, "YYYY/MM/DD") & "', CASInvMonth = '" & Format(Month(txtInvoiceDate), "00") & "', CASInvYear = '" & Year(txtInvoiceDate) & "'" & _
                    " where CASNumber ='" & Trim(tmpCASnumber) & "'"
        Else
            StrSql = "update CISCas set InvoiceNo = '" & TxtInvoiceNo & "', InvoiceDate = '" & Format(txtInvoiceDate, "YYYY/MM/DD") & "', CASInvMonth = '" & Format(Month(txtInvoiceDate), "00") & "', CASInvYear = '" & Year(txtInvoiceDate) & "'" & _
                    " where CASNumber ='" & Trim(tmpCASnumber) & "'"
        
        End If
            MediConnCISDB.Execute StrSql
    
    ElseIf txtInvoiceDate = "__/__/____" Then
        If tmpupdateOP = True Then
    
            StrSql = "update CISCas set OPBatchNo = '" & Trim(TxtOPBatchNo) & "', OPNoEnterUser =  '" & Logon & "'" & _
                " where CASNumber ='" & Trim(tmpCASnumber) & "'"
                
            MediConnCISDB.Execute StrSql
        'Else
       '
       '     strsql = "update CISCas set OPBatchNo = '" & Trim(TxtOPBatchNo) & "', OPNoEnterUser =  '" & Logon & "'" & _
       '         " where CASNumber ='" & Trim(tmpCASnumber) & "'"
        
        End If
    End If
        


    MediConnCISDB.commitTrans
    CompTrans = True


End Sub

Private Function GenerateOPSeq(OPVar)
Dim OPRec As New ADODB.Recordset
Dim StrSql As String
Dim SeqNo As String
Dim SEQ As New ADODB.Recordset
Dim SEQInsert As New ADODB.Recordset
Dim currentNo, i As Integer
'Dim strsql As String
Dim SEQNumber, SEQCode As String
Dim CharacterArray
Dim firstChar, currentChar As String
    
 

    
    SEQNumber = "0001"
    CharacterArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    'SEQCode = Trim(PayorCode & Trim(PlanCode) & Trim(InsuredType) & Trim(BusinessYear))
    'SEQCode = Trim(PayorCode & Trim(PlanCode) & Trim(InsuredType))
    'SEQCode = Trim(PayorCode & Trim(InsuredType))
    
    StrSql = "Select OPCode,OPBatchSeq from OPBatch Where OPCode = '" & OPVar & "'"
    OPRec.Open StrSql, MediConnCISDB, adOpenStatic, adLockOptimistic
    
    currentNo = 1
    currentChar = "A"
    
    If OPRec.recordCount = 0 Then
        'SEQInsert.Open "Select * from OPBatch ", medixmasterconnMaint, adOpenStatic, adLockOptimistic
        With OPRec
            .AddNew
            'With OPRec
                '.AddNew
                !OPCode = OPVar
                !OPBatchSeq = currentChar & "0001"
                .update
                SEQNumber = currentChar & "0001"
           ' End With
            .update
        End With
        'SEQInsert.Close
        'Set SEQInsert = Nothing
    Else
        currentNo = Mid(OPRec!OPBatchSeq, 2, 5)
        firstChar = Mid(OPRec!OPBatchSeq, 1, 1)
        Do Until i = 25
            If firstChar = CharacterArray(i) Then
                If currentNo <> "9999" Then
                    currentChar = CharacterArray(i)
                    currentNo = currentNo + 1
                Else
                    currentChar = CharacterArray(i + 1)
                    currentNo = "00001"
                End If
                Exit Do
            End If
            i = i + 1
        Loop
        SEQNumber = currentChar & Format(currentNo, "0000")
        OPRec!OPBatchSeq = SEQNumber
        OPRec.update
    End If

    GenerateOPSeq = "OP" & SEQNumber
    OPRec.Close
    Set OPRec = Nothing
    
End Function






