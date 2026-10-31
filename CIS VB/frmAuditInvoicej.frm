VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAuditInvoice 
   Caption         =   "Audit Invoice"
   ClientHeight    =   10320
   ClientLeft      =   0
   ClientTop       =   240
   ClientWidth     =   12840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10320
   ScaleWidth      =   12840
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   550
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   5985
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4980
         TabIndex        =   32
         Top             =   200
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   31
         Top             =   200
         Width           =   540
      End
      Begin VB.CommandButton CmdPrev 
         Caption         =   "Pre&v"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3900
         TabIndex        =   30
         Top             =   200
         Width           =   540
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   1
         Top             =   200
         Width           =   540
      End
      Begin VB.TextBox txtCaseNo 
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
         Left            =   840
         TabIndex        =   0
         Top             =   200
         Width           =   2415
      End
      Begin VB.Label Label42 
         Caption         =   "Case No."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   255
         Width           =   735
      End
   End
   Begin VB.Frame FrameCAS 
      Height          =   550
      Left            =   6120
      TabIndex        =   21
      Top             =   240
      Width           =   6615
      Begin VB.ComboBox cboPending 
         Enabled         =   0   'False
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
         Left            =   5040
         TabIndex        =   29
         Text            =   "--Select--"
         Top             =   160
         Width           =   975
      End
      Begin VB.ComboBox cboClaimability 
         Enabled         =   0   'False
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
         Left            =   2760
         TabIndex        =   27
         Text            =   "--Select--"
         Top             =   160
         Width           =   1455
      End
      Begin VB.ComboBox cboStatus 
         Enabled         =   0   'False
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
         Left            =   600
         TabIndex        =   28
         Text            =   "--Select--"
         Top             =   160
         Width           =   1095
      End
      Begin VB.Label Label49 
         Caption         =   "Pending"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label44 
         Caption         =   "Claimability"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label107 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   5160
      TabIndex        =   37
      Top             =   840
      Width           =   7515
      Begin VB.TextBox txtMBMNo 
         Enabled         =   0   'False
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
         Left            =   840
         TabIndex        =   38
         Top             =   200
         Width           =   3585
      End
      Begin VB.TextBox txtMBMName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   40
         Top             =   480
         Width           =   3585
      End
      Begin VB.TextBox txtPatientIC 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   39
         Top             =   770
         Width           =   3585
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
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
         Left            =   840
         TabIndex        =   41
         Top             =   1050
         Width           =   3585
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
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
         Left            =   120
         TabIndex        =   45
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label54 
         Caption         =   "MBM No"
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
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPatientIC 
         Caption         =   "IC No."
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
         TabIndex        =   43
         Top             =   800
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Dept"
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
         Left            =   120
         TabIndex        =   42
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      Height          =   555
      Left            =   120
      TabIndex        =   18
      Top             =   9600
      Width           =   12600
      Begin VB.CommandButton CmdClearAll 
         Caption         =   "&Clear All"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9840
         TabIndex        =   12
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10680
         TabIndex        =   13
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9120
         TabIndex        =   11
         Top             =   160
         Width           =   735
      End
   End
   Begin VB.Frame FrameMBM 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   4890
      Begin VB.TextBox txtClinicCode 
         Enabled         =   0   'False
         Height          =   330
         Left            =   795
         TabIndex        =   2
         Top             =   180
         Width           =   580
      End
      Begin VB.TextBox txtClinicName 
         Enabled         =   0   'False
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
         Left            =   1360
         TabIndex        =   3
         Top             =   180
         Width           =   3210
      End
      Begin VB.TextBox txtNPClinic 
         Height          =   330
         Left            =   1360
         TabIndex        =   4
         Top             =   180
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.TextBox txtInvNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   795
         MaxLength       =   30
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtInvDate 
         Height          =   315
         Left            =   795
         TabIndex        =   7
         Top             =   765
         Width           =   1575
         _ExtentX        =   2794
         _ExtentY        =   550
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
      Begin MSMask.MaskEdBox txtInvRecDate 
         Height          =   315
         Left            =   3240
         TabIndex        =   8
         Top             =   480
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   550
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
      Begin VB.ComboBox cboProcMth 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         TabIndex        =   10
         Top             =   765
         Width           =   1335
      End
      Begin VB.ComboBox CboPayto 
         Enabled         =   0   'False
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
         Left            =   795
         TabIndex        =   5
         Text            =   "C - Clinic"
         Top             =   1050
         Width           =   1575
      End
      Begin VB.ComboBox CboInvRecMode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3240
         TabIndex        =   9
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label Label101 
         Caption         =   "Mth"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   36
         ToolTipText     =   "Processing Month"
         Top             =   795
         Width           =   735
      End
      Begin VB.Label lblClnType 
         Caption         =   "lblClnType"
         Height          =   255
         Left            =   3240
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label110 
         Caption         =   "RecMode"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   34
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label78 
         Caption         =   "RecDate"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   33
         Top             =   525
         Width           =   735
      End
      Begin VB.Label Label109 
         Caption         =   "Pay To"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1095
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Clinic"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         Top             =   260
         Width           =   735
      End
      Begin VB.Label Label67 
         Caption         =   "InvDate"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   795
         Width           =   735
      End
      Begin VB.Label Label35 
         Caption         =   "InvNo"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   555
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   46
      Top             =   2520
      Width           =   12495
      _ExtentX        =   22035
      _ExtentY        =   12510
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "General Information"
      TabPicture(0)   =   "frmAuditInvoicej.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameInv"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameMedINjDiag"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmauditreview"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Documents"
      TabPicture(1)   =   "frmAuditInvoicej.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame20(19)"
      Tab(1).Control(1)=   "List1"
      Tab(1).Control(2)=   "Label2(6)"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame20 
         Height          =   1035
         Index           =   19
         Left            =   -70920
         TabIndex        =   122
         Top             =   840
         Width           =   6765
         Begin VB.CommandButton CmdBrowse 
            Caption         =   "Browse"
            Height          =   255
            Left            =   5760
            TabIndex        =   126
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TxtSrc 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   125
            Top             =   240
            Width           =   4575
         End
         Begin VB.TextBox TxtDest 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   124
            Top             =   600
            Width           =   4575
         End
         Begin VB.CommandButton CmdCopy 
            Caption         =   "Copy"
            Height          =   255
            Left            =   5760
            TabIndex        =   123
            Top             =   630
            Width           =   735
         End
         Begin MSComDlg.CommonDialog CmnDlg 
            Left            =   6840
            Top             =   360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label65 
            Caption         =   "Source :"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   128
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Destination :"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   127
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.ListBox List1 
         Height          =   1008
         ItemData        =   "frmAuditInvoicej.frx":0038
         Left            =   -74880
         List            =   "frmAuditInvoicej.frx":003A
         TabIndex        =   120
         ToolTipText     =   "right click to randomize contents"
         Top             =   840
         Width           =   3615
      End
      Begin VB.Frame frmauditreview 
         Caption         =   "Audit Review"
         Height          =   2535
         Left            =   120
         TabIndex        =   114
         Top             =   3960
         Width           =   12255
         Begin VB.TextBox txtremarks 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   7560
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   135
            Top             =   320
            Width           =   4575
         End
         Begin VB.TextBox txtanswer 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1680
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   134
            Top             =   1040
            Width           =   4575
         End
         Begin VB.TextBox txtquery 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1680
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   133
            Top             =   320
            Width           =   4575
         End
         Begin VB.TextBox txtauditby 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7560
            MaxLength       =   30
            TabIndex        =   131
            Top             =   1030
            Width           =   1575
         End
         Begin VB.TextBox txtdiscountreceived 
            Height          =   285
            Left            =   1680
            TabIndex        =   119
            Top             =   1755
            Width           =   2775
         End
         Begin MSMask.MaskEdBox txtauditdate 
            Height          =   315
            Left            =   7560
            TabIndex        =   129
            Top             =   1330
            Width           =   1575
            _ExtentX        =   2773
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
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
         Begin VB.Label lblauditby 
            Caption         =   "Audit By"
            Height          =   255
            Index           =   1
            Left            =   6480
            TabIndex        =   132
            Top             =   1030
            Width           =   735
         End
         Begin VB.Label lblauditdate 
            Caption         =   "Audit Date"
            Height          =   255
            Index           =   0
            Left            =   6480
            TabIndex        =   130
            Top             =   1330
            Width           =   975
         End
         Begin VB.Label lbldiscountreceived 
            Caption         =   "Cost Avoidance"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   118
            Top             =   1745
            Width           =   1215
         End
         Begin VB.Label lblremarks 
            Caption         =   "Remarks"
            Height          =   255
            Index           =   2
            Left            =   6480
            TabIndex        =   117
            Top             =   320
            Width           =   735
         End
         Begin VB.Label lblanswer 
            Caption         =   "Answer"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   116
            Top             =   1040
            Width           =   975
         End
         Begin VB.Label lblquery 
            Caption         =   "Query"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   115
            Top             =   320
            Width           =   975
         End
      End
      Begin VB.Frame FrameMedINjDiag 
         Height          =   2835
         Left            =   120
         TabIndex        =   56
         Top             =   1080
         Width           =   12255
         Begin VB.Frame FrameDiag 
            Caption         =   "Diagnosis"
            Height          =   1815
            Left            =   120
            TabIndex        =   96
            Top             =   120
            Width           =   4455
            Begin VB.TextBox txtDiagnosis1 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               MaxLength       =   10
               TabIndex        =   102
               Top             =   285
               Width           =   615
            End
            Begin VB.TextBox txtDiagnosis2 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               MaxLength       =   10
               TabIndex        =   103
               Top             =   555
               Width           =   615
            End
            Begin VB.TextBox txtDiagnosis3 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               MaxLength       =   10
               TabIndex        =   104
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox txtDiagnosis4 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               MaxLength       =   10
               TabIndex        =   105
               Top             =   1125
               Width           =   615
            End
            Begin VB.TextBox TxtDiagnosis5 
               Enabled         =   0   'False
               Height          =   285
               Left            =   240
               MaxLength       =   100
               TabIndex        =   106
               Top             =   1395
               Width           =   615
            End
            Begin VB.TextBox TxtDiagDesc1 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   885
               MaxLength       =   10
               TabIndex        =   97
               Top             =   285
               Width           =   3195
            End
            Begin VB.TextBox TxtDiagDesc2 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   885
               MaxLength       =   10
               TabIndex        =   98
               Top             =   555
               Width           =   3195
            End
            Begin VB.TextBox TxtDiagDesc3 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   885
               MaxLength       =   10
               TabIndex        =   99
               Top             =   840
               Width           =   3195
            End
            Begin VB.TextBox TxtDiagDesc4 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   885
               MaxLength       =   10
               TabIndex        =   100
               Top             =   1125
               Width           =   3195
            End
            Begin VB.TextBox TxtDiagDesc5 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   885
               MaxLength       =   10
               TabIndex        =   101
               Top             =   1395
               Width           =   3195
            End
            Begin VB.Label Label28 
               Caption         =   "5"
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
               TabIndex        =   111
               Top             =   1455
               Width           =   255
            End
            Begin VB.Label Label6 
               Caption         =   "4"
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
               TabIndex        =   110
               Top             =   1185
               Width           =   255
            End
            Begin VB.Label Label52 
               Caption         =   "3"
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
               TabIndex        =   109
               Top             =   900
               Width           =   255
            End
            Begin VB.Label Label51 
               Caption         =   "2"
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
               TabIndex        =   108
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Label12 
               Caption         =   "1"
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
               TabIndex        =   107
               Top             =   360
               Width           =   135
            End
         End
         Begin VB.Frame FrameMed 
            Caption         =   "Medication && Injection"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   4680
            TabIndex        =   57
            Top             =   120
            Width           =   7335
            Begin VB.TextBox txtMedType1 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   240
               MaxLength       =   50
               TabIndex        =   58
               Top             =   405
               Width           =   4335
            End
            Begin VB.TextBox txtMedType1 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   240
               TabIndex        =   69
               Top             =   690
               Width           =   4335
            End
            Begin VB.TextBox txtMedType1 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   240
               TabIndex        =   70
               Top             =   960
               Width           =   4335
            End
            Begin VB.TextBox txtMedType1 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   240
               TabIndex        =   71
               Top             =   1200
               Width           =   4335
            End
            Begin VB.TextBox txtMedType1 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   240
               TabIndex        =   72
               Top             =   1490
               Width           =   4335
            End
            Begin VB.TextBox txtMedType1 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   5
               Left            =   240
               TabIndex        =   73
               Top             =   1750
               Width           =   4335
            End
            Begin VB.TextBox txtMedType1 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   6
               Left            =   240
               TabIndex        =   74
               Top             =   2040
               Width           =   4335
            End
            Begin VB.TextBox txtMedDay1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   59
               Top             =   405
               Width           =   855
            End
            Begin VB.TextBox txtMedDay1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   61
               Top             =   690
               Width           =   855
            End
            Begin VB.TextBox txtMedDay1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   63
               Top             =   945
               Width           =   855
            End
            Begin VB.TextBox txtMedDay1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   65
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtMedDay1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   66
               Top             =   1485
               Width           =   855
            End
            Begin VB.TextBox txtMedDay1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   5
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   67
               Top             =   1740
               Width           =   855
            End
            Begin VB.TextBox txtMedDay1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   6
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   68
               Top             =   2030
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   5400
               MaxLength       =   3
               TabIndex        =   60
               Top             =   405
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   5400
               MaxLength       =   3
               TabIndex        =   62
               Top             =   690
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   5400
               MaxLength       =   3
               TabIndex        =   64
               Top             =   945
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   5400
               MaxLength       =   3
               TabIndex        =   78
               Top             =   1230
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   5400
               MaxLength       =   3
               TabIndex        =   83
               Top             =   1485
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   5
               Left            =   5400
               MaxLength       =   3
               TabIndex        =   84
               Top             =   1740
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   6
               Left            =   5400
               MaxLength       =   3
               TabIndex        =   85
               Top             =   2030
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   7
               Left            =   6240
               TabIndex        =   75
               Top             =   405
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   8
               Left            =   6240
               TabIndex        =   76
               Top             =   690
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   9
               Left            =   6240
               TabIndex        =   77
               Top             =   945
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   10
               Left            =   6240
               TabIndex        =   79
               Top             =   1230
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   11
               Left            =   6240
               TabIndex        =   80
               Top             =   1485
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   12
               Left            =   6240
               TabIndex        =   81
               Top             =   1740
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   13
               Left            =   6240
               TabIndex        =   82
               Top             =   2030
               Width           =   855
            End
            Begin VB.Label Label26 
               Caption         =   "7"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   95
               Top             =   2080
               Width           =   255
            End
            Begin VB.Label Labe26 
               Caption         =   "6"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   94
               Top             =   1800
               Width           =   255
            End
            Begin VB.Label Label72 
               Caption         =   "Qty"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5640
               TabIndex        =   93
               Top             =   150
               Width           =   375
            End
            Begin VB.Label Label71 
               Caption         =   "Days"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4800
               TabIndex        =   92
               Top             =   150
               Width           =   375
            End
            Begin VB.Label Label26 
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   91
               Top             =   1515
               Width           =   975
            End
            Begin VB.Label Label26 
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   90
               Top             =   1000
               Width           =   855
            End
            Begin VB.Label Label26 
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   89
               Top             =   1280
               Width           =   975
            End
            Begin VB.Label Label26 
               Caption         =   "1"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   88
               Top             =   480
               Width           =   135
            End
            Begin VB.Label Label26 
               Caption         =   "2"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   87
               Top             =   720
               Width           =   135
            End
            Begin VB.Label Label74 
               Caption         =   "Price"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   17
               Left            =   6480
               TabIndex        =   86
               Top             =   150
               Width           =   495
            End
         End
         Begin VB.TextBox txtPriDiag 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   113
            Top             =   2040
            Width           =   3015
         End
         Begin VB.Label lblPriDiag 
            Caption         =   "Preliminary Diag"
            Height          =   255
            Left            =   100
            TabIndex        =   112
            Top             =   2040
            Width           =   1155
         End
      End
      Begin VB.Frame FrameInv 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   8775
         Begin VB.TextBox txtBillNo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   20
            TabIndex        =   49
            Top             =   200
            Width           =   1305
         End
         Begin VB.TextBox txtCLNBillAmt 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3360
            MaxLength       =   20
            TabIndex        =   48
            Top             =   200
            Width           =   1305
         End
         Begin MSMask.MaskEdBox txtVisitDate 
            Height          =   300
            Left            =   5520
            TabIndex        =   50
            Top             =   200
            Width           =   1155
            _ExtentX        =   2032
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
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
         Begin MSMask.MaskEdBox txtVisitTime 
            Height          =   300
            Left            =   7440
            TabIndex        =   51
            Top             =   200
            Width           =   1155
            _ExtentX        =   2032
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label66 
            Caption         =   "Vst Time"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6720
            TabIndex        =   55
            Top             =   200
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "Vst Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4800
            TabIndex        =   54
            Top             =   200
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Bill No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   200
            Width           =   975
         End
         Begin VB.Label Label101 
            Caption         =   "Bill Amt"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   2280
            TabIndex        =   52
            Top             =   200
            Width           =   855
         End
      End
      Begin VB.Label Label2 
         Caption         =   "GL Documents"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   121
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Label lblInvoice 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Audit Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   19
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmAuditInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tmpPayto As String
Dim PictureFileName As String
Private Const SW_RESTORE        As Long = &H9&
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private WithEvents cFileCopy As clsFileCopy
Attribute cFileCopy.VB_VarHelpID = -1
Dim Auditcase As Boolean

Private Sub CmdUpdate_Click()
Dim StrSqlCRef As String
Const METHOD_NAME = "CmdUpdate_Click"
On Error GoTo debugs
Dim RecSaved
    If Len(Trim(txtCaseNo)) = "0" Or Auditcase = False Then
        MsgBox "Not Record Found ! ", vbOKOnly + vbCritical, DialogBoxTitle
        txtCaseNo.SetFocus
        Exit Sub
    End If
    
    RecSaved = MsgBox("Are you sure want update Audit Details? ", vbYesNo + vbExclamation)
    If Trim(txtquery) = "" Then
        MsgBox "Please Enter Query! ", vbCritical
        txtquery.SetFocus
    ElseIf Trim(txtauditby) = "" Then
            MsgBox "Please Enter Audit By! ", vbCritical
            txtauditby.SetFocus
    ElseIf Trim(txtauditdate) <> "" And (txtauditdate = "__/__/____" Or txtauditdate = "01/01/1900") Then
            MsgBox "Please Enter Audit Date! ", vbCritical
            txtauditdate.SetFocus
    ElseIf IsDate(txtauditdate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtauditdate.SetFocus
    ElseIf RecSaved = vbYes And Auditcase = True Then
    
        StrSqlCRef = ""
        StrSqlCRef = " Update CISDB..CISCas SET " & _
                     " AuditQuery = '" & Trim(Replace(txtquery, "'", "''")) & "'," & _
                     " AuditAnswer = '" & Trim(Replace(txtanswer, "'", "''")) & "'," & _
                     " AuditDiscountReceived = '" & Trim(txtdiscountreceived) & "'," & _
                     " AuditRemarks = '" & Trim(Replace(txtremarks, "'", "''")) & "'" & _
                     " ,AuditBy = '" & Trim(Replace(txtauditby, "'", "''")) & "'" & _
                     " ,AuditDate = '" & Format(txtauditdate, "MM/DD/yyyy") & "'" & _
                     " Where CASNumber = '" & Trim(txtCaseNo) & "'"
        MediConnCISDB.Execute StrSqlCRef
        MsgBox "Audit Details updated ! ", vbInformation
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "frmAuditInvoice")
    Auditcase = False
     Call ComboListing
End Sub
Private Sub txtdiscountreceived_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub txtauditdate_LostFocus()
    If IsDate(txtauditdate) Then
    Else
    If txtauditdate <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtauditdate.SetFocus
    End If
    End If
End Sub

Private Sub ComboListing()
Const METHOD_NAME = "ComboListing"
On Error GoTo debugs
    cboStatus.AddItem "--Select--"
    cboStatus.AddItem "O - Open"
    cboStatus.AddItem "C - Close"
    cboStatus.AddItem "D - Delete"
    
    cboClaimability.AddItem "--Select--"
    cboClaimability.AddItem "A - Accepted"
    cboClaimability.AddItem "R - Repudiated"
    cboClaimability.AddItem "C - Cancelled"
    cboClaimability.AddItem "S - Special Approve"
    
    cboPending.AddItem "Y - Yes"
    cboPending.AddItem "N - No"
    
    CboPayto.AddItem "C - CLINIC"
    CboPayto.AddItem "M - MEMBER"
    
debugs:
If Err.Description <> "" Then
Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub ClearEntries()
   On Error Resume Next
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Name <> "txtCaseNo" Then
               ctl.Text = ""
            End If
        ElseIf TypeOf ctl Is ComboBox Then
        ElseIf TypeOf ctl Is MaskEdBox Then
            If ctl.Name <> "txtInvDate" Then
                ctl.mask = ""
                ctl.Text = ""
                ctl.mask = "##/##/####"
            End If
            If ctl.Name = "txtVisitTime" Then
                ctl.mask = "##:##"
            End If
        End If
    Next ctl
    txtClinicName = ""
    
End Sub
    
    
Public Sub CmdGo_Click()
Const METHOD_NAME = "CmdGo_Click"
On Error GoTo debugs

Dim tmpDiagDesc As String
Dim CASSql As String
Set CISCAS1 = New ADODB.Recordset
SSTab1.Tab = 0
Call ClearEntries
CASSql = ""
CASSql = "Exec SP_CISAuditdetails '" & Trim(txtCaseNo) & "'"
CISCAS1.Open CASSql, MediConnCISDB, adOpenKeyset, adLockOptimistic

If CISCAS1.EOF = True Then
    Auditcase = False
    MsgBox "Not Record Found ! ", vbOKOnly + vbCritical, DialogBoxTitle
Else

            Auditcase = True
            If Trim(CISCAS1!CASStatus) = "O" Then
                cboStatus.ListIndex = 1
            ElseIf Trim(CISCAS1!CASStatus) = "C" Then
                cboStatus.ListIndex = 2
            ElseIf Trim(CISCAS1!CASStatus) = "D" Then
                cboStatus.ListIndex = 3
            End If
            
             If Len(CISCAS1!Claimability) Then
                If Trim(CISCAS1!Claimability) = "A" Then
                    cboClaimability.ListIndex = 1
                ElseIf Trim(CISCAS1!Claimability) = "R" Then
                    cboClaimability.ListIndex = 2
                ElseIf Trim(CISCAS1!Claimability) = "C" Then
                    cboClaimability.ListIndex = 3
                ElseIf Trim(CISCAS1!Claimability) = "S" Then
                    cboClaimability.ListIndex = 4
                End If
            End If
            
            If Trim(CISCAS1!Pending) = "Y" Then
                cboPending.Text = "Y - Yes"
            ElseIf Trim(CISCAS1!Pending) = "N" Then
                cboPending.Text = "N - No"
            End If
            
            txtClinicCode = IIf(Len(CISCAS1!ClinicCode), CISCAS1!ClinicCode, "")
            txtClinicName = IIf(Len(CISCAS1!CLNName), CISCAS1!CLNName, "")
            txtInvNo = IIf(Len(CISCAS1!InvoiceNo), CISCAS1!InvoiceNo, "")
            txtInvRecDate = IIf(Len(CISCAS1!InvoiceReceiptDate), CISCAS1!InvoiceReceiptDate, "__/__/____")
            txtInvDate = IIf(Len(CISCAS1!InvoiceDate), (CISCAS1!InvoiceDate), "__/__/____")
            cboProcMth = IIf(Len(CISCAS1!CLMProcMth), CISCAS1!CLMProcMth & "/" & CISCAS1!CLMProcYR, "")
            
            CboPayto.ListIndex = 0
            If Trim(CISCAS1!INVPayTo) = "M" Then
                CboPayto.ListIndex = 1
                tmpPayto = "M"
            Else
                tmpPayto = "C"
            End If
            CboInvRecMode = IIf(Len(CISCAS1!InvoiceReceiptMode), CISCAS1!InvoiceReceiptMode, "")
            txtMBMNo = CISCAS1!MBMNumber & "-" & CISCAS1!coverid
            txtMBMName = IIf(Len(CISCAS1!MBMName), CISCAS1!MBMName, "")
            txtPatientIC = IIf(Len(CISCAS1!MBMIcBcPp), CISCAS1!MBMIcBcPp, "")
            txtDept = IIf(Len(CISCAS1!MBMDepartment), CISCAS1!MBMDepartment, "")
            txtBillNo = IIf(Len(Trim(CISCAS1!BillNo)), Trim(CISCAS1!BillNo), "")
            txtCLNBillAmt = IIf(IsNumeric(CISCAS1!INVBillAmt), CISCAS1!INVBillAmt, 0)
            txtVisitDate = CISCAS1!DateVisit
            If Len(Trim(CISCAS1!TimeVisit)) Then
                txtVisitTime = Format(CISCAS1!TimeVisit, "hh:mm")
            End If
            
            txtDiagnosis1 = IIf(Len(Trim(CISCAS1!CASTreatReason1)), CISCAS1!CASTreatReason1, "")
            txtDiagnosis2 = IIf(Len(Trim(CISCAS1!CASTreatReason2)), CISCAS1!CASTreatReason2, "")
            txtDiagnosis3 = IIf(Len(Trim(CISCAS1!CASTreatReason3)), CISCAS1!CASTreatReason3, "")
            txtDiagnosis4 = IIf(Len(Trim(CISCAS1!CASTreatReason4)), CISCAS1!CASTreatReason4, "")
            TxtDiagDesc5.Visible = True
            TxtDiagnosis5 = IIf(Len(Trim(CISCAS1!CASTreatReason5)), CISCAS1!CASTreatReason5, "")
            If Trim(txtDiagnosis1) <> "" Then
                Call DiagDesc(txtDiagnosis1, tmpDiagDesc)
                TxtDiagDesc1 = tmpDiagDesc
            End If
            If Trim(txtDiagnosis2) <> "" Then
                Call DiagDesc(txtDiagnosis2, tmpDiagDesc)
                TxtDiagDesc2 = tmpDiagDesc
            End If
            If Trim(txtDiagnosis3) <> "" Then
                Call DiagDesc(txtDiagnosis3, tmpDiagDesc)
                TxtDiagDesc3 = tmpDiagDesc
            End If
            If Trim(txtDiagnosis4) <> "" Then
                Call DiagDesc(txtDiagnosis4, tmpDiagDesc)
                TxtDiagDesc4 = tmpDiagDesc
            End If
            If Trim(TxtDiagnosis5) <> "" Then
                Call DiagDesc(TxtDiagnosis5, tmpDiagDesc)
                If Trim(tmpDiagDesc) = "OTHERS" Then
                    TxtDiagDesc5 = IIf(Len(Trim(CISCAS1!CASDiagnosis5)), CISCAS1!CASDiagnosis5, "")
                End If
                If TxtDiagDesc5 = "" Then
                    TxtDiagDesc5 = tmpDiagDesc
                End If
            End If
            
            'If Trim(CISCAS1!CASDiagnosis5) <> "" Then
            'TxtDiagnosisOth = IIf(Len(Trim(CISCAS1!CASDiagnosis5)), CISCAS1!CASDiagnosis5, "")
            'End If
            txtPriDiag = Trim(CISCAS1!PrilimDiag)
            
            txtMedType1(0) = IIf(Len(Trim(CISCAS1!MedInjType1)), CISCAS1!MedInjType1, "")
            txtMedType1(1) = IIf(Len(Trim(CISCAS1!MedInjType2)), CISCAS1!MedInjType2, "")
            txtMedType1(2) = IIf(Len(Trim(CISCAS1!MedInjType3)), CISCAS1!MedInjType3, "")
            txtMedType1(3) = IIf(Len(Trim(CISCAS1!MedInjType4)), CISCAS1!MedInjType4, "")
            txtMedType1(4) = IIf(Len(Trim(CISCAS1!MedInjType5)), CISCAS1!MedInjType5, "")
            txtMedType1(5) = IIf(Len(Trim(CISCAS1!MedInjType6)), CISCAS1!MedInjType6, "")
            txtMedType1(6) = IIf(Len(Trim(CISCAS1!MedInjType7)), CISCAS1!MedInjType7, "")
            txtMedDay1(0) = IIf(Len(Trim(CISCAS1!MedInjDays1)), CISCAS1!MedInjDays1, "")
            txtMedDay1(1) = IIf(Len(Trim(CISCAS1!MedInjDays2)), CISCAS1!MedInjDays2, "")
            txtMedDay1(2) = IIf(Len(Trim(CISCAS1!MedInjDays3)), CISCAS1!MedInjDays3, "")
            txtMedDay1(3) = IIf(Len(Trim(CISCAS1!MedInjDays4)), CISCAS1!MedInjDays4, "")
            txtMedDay1(4) = IIf(Len(Trim(CISCAS1!MedInjDays5)), CISCAS1!MedInjDays5, "")
            txtMedDay1(5) = IIf(Len(Trim(CISCAS1!MedInjDays6)), CISCAS1!MedInjDays6, "")
            txtMedDay1(6) = IIf(Len(Trim(CISCAS1!MedInjDays7)), CISCAS1!MedInjDays7, "")
            txtMedQty1(0) = IIf(Len(Trim(CISCAS1!MedInjQty1)), CISCAS1!MedInjQty1, "")
            txtMedQty1(1) = IIf(Len(Trim(CISCAS1!MedInjQty2)), CISCAS1!MedInjQty2, "")
            txtMedQty1(2) = IIf(Len(Trim(CISCAS1!MedInjQty3)), CISCAS1!MedInjQty3, "")
            txtMedQty1(3) = IIf(Len(Trim(CISCAS1!MedInjQty4)), CISCAS1!MedInjQty4, "")
            txtMedQty1(4) = IIf(Len(Trim(CISCAS1!MedInjQty5)), CISCAS1!MedInjQty5, "")
            txtMedQty1(5) = IIf(Len(Trim(CISCAS1!MedInjQty6)), CISCAS1!MedInjQty6, "")
            txtMedQty1(6) = IIf(Len(Trim(CISCAS1!MedInjQty7)), CISCAS1!MedInjQty7, "")
            ''Price
            txtMedQty1(7) = IIf(Len(Trim(CISCAS1!MEDCHARGES1)), CISCAS1!MEDCHARGES1, "")
            txtMedQty1(8) = IIf(Len(Trim(CISCAS1!MEDCHARGES2)), CISCAS1!MEDCHARGES2, "")
            txtMedQty1(9) = IIf(Len(Trim(CISCAS1!MEDCHARGES3)), CISCAS1!MEDCHARGES3, "")
            txtMedQty1(10) = IIf(Len(Trim(CISCAS1!MEDCHARGES4)), CISCAS1!MEDCHARGES4, "")
            txtMedQty1(11) = IIf(Len(Trim(CISCAS1!MEDCHARGES5)), CISCAS1!MEDCHARGES5, "")
            txtMedQty1(12) = IIf(Len(Trim(CISCAS1!MEDCHARGES6)), CISCAS1!MEDCHARGES6, "")
            txtMedQty1(13) = IIf(Len(Trim(CISCAS1!MEDCHARGES7)), CISCAS1!MEDCHARGES7, "")
            
            txtquery = IIf(Len(Trim(CISCAS1!AuditQuery)), CISCAS1!AuditQuery, "")
            txtanswer = IIf(Len(Trim(CISCAS1!AuditAnswer)), CISCAS1!AuditAnswer, "")
            txtdiscountreceived = IIf(Len(Trim(CISCAS1!AuditDiscountReceived)), CISCAS1!AuditDiscountReceived, "")
            txtremarks = IIf(Len(Trim(CISCAS1!AuditRemarks)), CISCAS1!AuditRemarks, "")
            txtauditby = IIf(Len(Trim(CISCAS1!AuditBy)), CISCAS1!AuditBy, "")
            txtauditdate = IIf(Len(CISCAS1!AuditDate), CISCAS1!AuditDate, "__/__/____")
            CmdUpdate.Enabled = True
            
            
End If
CISCAS1.Close
Set CISCAS1 = Nothing

debugs:
    If Err.Description <> "" Then
        Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
    End If
End Sub


Private Sub CmdPrev_Click()
Const METHOD_NAME = "CmdPrev_Click"
On Error GoTo debugs
Dim CAS As New ADODB.Recordset
Dim sql As String
    sql = " select CASNumber from VIEW_Cas_MBM Where MBMCLNDataStatus <> 'D' " & _
          " AND MBMCLNStatus <> 'D' AND (HSCLNInd = 'C' OR HSCLNInd = 'M' OR HSCLNInd = 'B')" & _
          " AND CaseType='C' AND CASNumber < '" & RemStr(txtCaseNo) & "' Order by CASNumber DESC"
    
    CAS.MaxRecords = 1
    CAS.Open sql, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
    
    If CAS.EOF = False Then
        txtCaseNo.Text = Trim(CAS!CasNumber)
        CAS.Close
        Set CAS = Nothing
        CmdGo_Click
    Else
        MsgBox "First Record Reached! " & vbNewLine & "Not a valid Case Number! ", vbCritical + vbOKOnly
        CAS.Close
        Set CAS = Nothing
    End If
debugs:
If Err.Description <> "" Then
   Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub CmdNext_Click()
Const METHOD_NAME = "CmdNext_Click"
On Error GoTo debugs
Dim CAS As New ADODB.Recordset
Dim sql As String
    sql = " select CASNumber from VIEW_Cas_MBM Where MBMCLNDataStatus <> 'D' " & _
          " AND MBMCLNStatus <> 'D' AND (HSCLNInd = 'C' OR HSCLNInd = 'M' OR HSCLNInd = 'B')" & _
          " AND CaseType='C' AND CASNumber > '" & RemStr(txtCaseNo) & "' Order by CASNumber "
    
    CAS.MaxRecords = 1
    CAS.Open sql, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
    
    If CAS.EOF = False Then
        txtCaseNo.Text = Trim(CAS!CasNumber)
        CAS.Close
        Set CAS = Nothing
        CmdGo_Click
    Else
        MsgBox " Last Record Reached! ", vbCritical + vbOKOnly
        CAS.Close
        Set CAS = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdSearch_Click()
    frmCasList.Source = 1
    frmCasList.Show vbModal
End Sub


Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdClearAll_Click()
Call ComboListing
txtClinicCode = ""
txtClinicName = ""
txtInvNo = ""
txtInvRecDate = "__/__/____"
txtInvDate = "__/__/____"
cboProcMth = ""
CboInvRecMode = ""
txtMBMNo = ""
txtMBMName = ""
txtPatientIC = ""
txtDept = ""
txtBillNo = ""
txtCLNBillAmt = ""
txtVisitDate = "__/__/____"
txtVisitTime = "__:__"
txtDiagnosis1 = ""
txtDiagnosis2 = ""
txtDiagnosis3 = ""
txtDiagnosis4 = ""
TxtDiagnosis5 = ""
TxtDiagDesc1 = ""
TxtDiagDesc2 = ""
TxtDiagDesc3 = ""
TxtDiagDesc4 = ""
TxtDiagDesc5 = ""
TxtDiagDesc6 = ""
txtPriDiag = ""
txtMedType1(0) = ""
txtMedType1(1) = ""
txtMedType1(2) = ""
txtMedType1(3) = ""
txtMedType1(4) = ""
txtMedType1(5) = ""
txtMedType1(6) = ""
txtMedDay1(0) = ""
txtMedDay1(1) = ""
txtMedDay1(2) = ""
txtMedDay1(3) = ""
txtMedDay1(4) = ""
txtMedDay1(5) = ""
txtMedDay1(6) = ""
txtMedQty1(0) = ""
txtMedQty1(1) = ""
txtMedQty1(2) = ""
txtMedQty1(3) = ""
txtMedQty1(4) = ""
txtMedQty1(5) = ""
txtMedQty1(6) = ""
txtMedQty1(7) = ""
txtMedQty1(8) = ""
txtMedQty1(9) = ""
txtMedQty1(10) = ""
txtMedQty1(11) = ""
txtMedQty1(12) = ""
txtMedQty1(13) = ""
txtquery = ""
txtanswer = ""
txtdiscountreceived = ""
txtremarks = ""
txtCaseNo = ""
List1.Clear
CmdUpdate.Enabled = False
SSTab1.Tab = 0
txtauditdate = "__/__/____"
txtauditby = ""
End Sub


Private Sub DiagDesc(DIagCode As String, tmpDiagDesc As String)
Const METHOD_NAME = "DiagDesc"
On Error GoTo debugs
Dim DiagRec As New ADODB.Recordset
Dim StrDiagSql As String

    If Trim(DIagCode) <> "" Then
        StrDiagSql = " SELECT ReasonCode, ReasonDesc " & _
                 " from TreatmentReason where 0=0 " & _
                 " and ReasonCode='" & Trim(DIagCode) & "'"
        DiagRec.Open StrDiagSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        
        Do While Not DiagRec.EOF
            tmpDiagDesc = DiagRec!ReasonDesc
            DiagRec.MoveNext
        Loop
        
        DiagRec.Close
        Set DiagRec = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub List1_Click()
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String
    If Len(txtCaseNo) Then
    PictureFileName = DocPath & "Outpatient\" & "Guarantee Letter\" & txtCaseNo & ""
    '\Medical Rpt"
    crt PictureFileName & "\" & List1.List(List1.ListIndex)

    'If Index = 0 Then
        sFile = PictureFileName & "\" & List1.List(List1.ListIndex)
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
    End If
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

Sub addTheFiles()
Dim B As String
List1.Clear
PictureFileName = DocPath & "Outpatient\" & "Guarantee Letter\" & txtCaseNo & ""
'\Medical Rpt"
B = Dir(PictureFileName & "\*.*")
'b = Dir(Dir1.Path & "\*.*")
If B <> "" Then
    If Right(B, 4) = ".jpg" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".tif" Then
        List1.AddItem B
    End If
    'If Right(b, 4) = ".art" Then
    '    List1.AddItem b
    'End If
    If Right(B, 4) = ".gif" Then
        List1.AddItem B
    End If
    
    If Right(B, 4) = ".JPG" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".TIF" Then
        List1.AddItem B
    End If
    'If Right(b, 4) = ".art" Then
    '    List1.AddItem b
    'End If
    If Right(B, 4) = ".GIF" Then
        List1.AddItem B
    End If
    
    If Right(B, 4) = ".doc" Then
        List1.AddItem B
    End If

    If Right(B, 4) = ".pdf" Then
        List1.AddItem B
    End If
    
    If UCase(Right(B, 4)) = ".PNG" Then
        List1.AddItem B
    End If
    
    If Right(B, 4) = ".msg" Then
        List1.AddItem B
    End If
    If Right(B, 5) = ".xlsx" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".dot" Then
        List1.AddItem B
    End If
    If Right(B, 5) = ".docx" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".txt" Then
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
        'If Right(b, 4) = ".art" Then
        '    List1.AddItem b
        'End If
        If Right(B, 4) = ".gif" Then
            List1.AddItem B
        End If
        
        If Right(B, 4) = ".JPG" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".TIF" Then
            List1.AddItem B
        End If
        'If Right(b, 4) = ".art" Then
        '    List1.AddItem b
        'End If
        If Right(B, 4) = ".GIF" Then
            List1.AddItem B
        End If
        
        If Right(B, 4) = ".doc" Then
            List1.AddItem B
        End If
        
        If Right(B, 4) = ".pdf" Then
            List1.AddItem B
        End If
        If UCase(Right(B, 4)) = ".PNG" Then
            List1.AddItem B
        End If
        
         If Right(B, 4) = ".msg" Then
            List1.AddItem B
        End If
          If Right(B, 5) = ".xlsx" Then
        List1.AddItem B
        End If
        If Right(B, 4) = ".dot" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".docx" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".txt" Then
            List1.AddItem B
        End If
        
    Else
        Exit Do
    End If
Loop
'List1.Height = Me.Height - 2520 - 350
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
        If Len(txtCaseNo) Then
            Call LoadDoc
        End If
    End If
End Sub
Private Sub LoadDoc()
On Error Resume Next
Dim B As String, f As Integer, d As Integer, e As Integer
Call addTheFiles
Exit Sub
End Sub
Private Sub CmdBrowse_Click()
Dim Cnt As Integer
Dim tmpstring As String
Dim x As Boolean
    x = False
    x = CreatePath(DocPath & "Outpatient\" & "Guarantee Letter\" & txtCaseNo)
'Set filter to all files
    tmpstring = ""
    Cnt = 0
    CmnDlg.Filename = ""
    CmnDlg.CancelError = True
      On Error GoTo ErrHandler
      
      'CmnDlg.InitDir = DocPath & "FaxDocuments"
      
      CmnDlg.Flags = cdlOFNHideReadOnly
      CmnDlg.Filter = "All files|*.*"
    'Show the open file window
    CmnDlg.ShowOpen
    'Put the file bath in the source text
    TxtSrc.Text = CmnDlg.Filename
    
    Dim sArr() As String
    Dim fname As String
    Dim lenth As Integer
    sArr = Split(TxtSrc, "\")
    lenth = UBound(sArr) + 1
    fname = sArr(lenth - 1)
    tmpstring = DocPath & "Outpatient\" & "Guarantee Letter\"
    Cnt = Len(tmpstring)
    TxtDest.Text = DocPath & "Outpatient\" & "Guarantee Letter\" & txtCaseNo & "\" & fname & ""
ErrHandler:
      Exit Sub
End Sub
Public Function CreatePath(NewPath) As Boolean
    If Right(NewPath, 1) <> "\" Then
        NewPath = NewPath & "\"
    End If
    If MakeSureDirectoryPathExists(NewPath) <> 0 Then
        CreatePath = True
    End If
End Function
Private Sub CmdCopy_Click()
On Error GoTo CopyErr
Dim ErrCode As Long
Dim ErrMsg As String
Dim tmpstring As String

Dim Cnt As Integer
   ' Screen.MousePointer = vbHourglass
    tmpstring = ""
    Cnt = 0
    tmpstring = DocPath & "Outpatient\" & "Guarantee Letter\"
    Cnt = Len(tmpstring)
    If txtCaseNo = "" Then
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
        'ElseIf UCase(Mid(TxtSrc.Text, 1, Cnt)) <> UCase(tmpString) Then
        'MsgBox "Please check source file!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    Else
    Set cFileCopy = New clsFileCopy
    'Disable the copy button
    CmdCopy.Enabled = False
    'The source file the you want to copy
    cFileCopy.SourceFile = TxtSrc
    'The destination file name
    cFileCopy.TargetFile = TxtDest
    
    cFileCopy.ByteSize = 4096 '4kb
    
    If Not cFileCopy.copyfile(ErrCode, ErrMsg) Then
        GoTo CopyErr
    End If
    Call LoadDoc
    
    MsgBox "File Transfered Successful..."

        GoTo Ex
    End If
    
CopyErr:
    If ErrCode = 0 Then
        ErrCode = Err.Number
        ErrMsg = Err.Description
    End If
    
    MsgBox "Error: " & ErrCode & " - " & ErrMsg, vbCritical, "Error"
    
Ex:
    
    Set cFileCopy = Nothing
    TxtSrc = ""
    TxtDest = ""
    'Re-enable the copy button
    CmdCopy.Enabled = True
End Sub
