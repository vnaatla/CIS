VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSUNSystem 
   Caption         =   "SUN System - Report"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9990
   ScaleWidth      =   14790
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   0
      TabIndex        =   46
      Top             =   9240
      Width           =   14775
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13920
         TabIndex        =   49
         Top             =   255
         Width           =   735
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13080
         TabIndex        =   48
         Top             =   255
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   12240
         TabIndex        =   47
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   14655
      Begin VB.TextBox txtAllocRefTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   102
         Top             =   240
         Width           =   1545
      End
      Begin VB.TextBox txtAllocRefFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   101
         Top             =   240
         Width           =   1545
      End
      Begin VB.TextBox txtDB 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtT9To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   97
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtT9Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   96
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtT8To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   93
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtT8Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   92
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtT7To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   89
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtT7Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   88
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtT6To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   85
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtT6Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   84
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtT5To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   81
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtT5Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   80
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtT4To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   77
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtT4Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   76
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtT3To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   73
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtT3Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   72
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtT2To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   69
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtT2Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   68
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtT1To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   65
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtT1Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   64
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtJournalTypeTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         MaxLength       =   15
         TabIndex        =   61
         Top             =   2040
         Width           =   1785
      End
      Begin VB.TextBox txtJournalTypeFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   60
         Top             =   2040
         Width           =   1905
      End
      Begin VB.TextBox txtDRCR 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         MaxLength       =   15
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtAllocated 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtTranRefFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   13
         Top             =   1680
         Width           =   1905
      End
      Begin VB.TextBox txtTranRefTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         MaxLength       =   15
         TabIndex        =   14
         Top             =   1680
         Width           =   1785
      End
      Begin VB.TextBox txtAmtFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtAmtTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   12
         Top             =   600
         Width           =   1545
      End
      Begin VB.TextBox txtPeriodYrTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         MaxLength       =   4
         TabIndex        =   8
         Top             =   960
         Width           =   960
      End
      Begin VB.TextBox txtPeriodYrFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   6
         Top             =   960
         Width           =   1080
      End
      Begin VB.TextBox txtPeriodMthFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtPeriodMthTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         MaxLength       =   2
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox ChkBatchNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   30
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkPeriodMth 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   29
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox ChkPeriodYr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   28
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox ChkGrpComp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   27
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtACCCodeFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   3
         Top             =   600
         Width           =   1905
      End
      Begin VB.TextBox txtACCCodeTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         MaxLength       =   30
         TabIndex        =   4
         Top             =   600
         Width           =   1785
      End
      Begin VB.TextBox txtOPBatchNo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7800
         MaxLength       =   20
         TabIndex        =   26
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtOPBatchNo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8040
         MaxLength       =   20
         TabIndex        =   25
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox CboPayor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8160
         TabIndex        =   24
         Top             =   2520
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox txtJournalNoFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1320
         Width           =   1905
      End
      Begin VB.TextBox txtJournalNoTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1320
         Width           =   1785
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSUNSystem.frx":0000
         Left            =   7680
         List            =   "frmSUNSystem.frx":0002
         TabIndex        =   23
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSUNSystem.frx":0004
         Left            =   7200
         List            =   "frmSUNSystem.frx":0006
         TabIndex        =   22
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSUNSystem.frx":0008
         Left            =   7920
         List            =   "frmSUNSystem.frx":000A
         TabIndex        =   21
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSUNSystem.frx":000C
         Left            =   7680
         List            =   "frmSUNSystem.frx":000E
         TabIndex        =   20
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboBordxMonthfr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   19
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboBordxMonthTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         TabIndex        =   18
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboBordxYearfr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   17
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboBordxYearTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         TabIndex        =   16
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtBordxDateFr 
         Height          =   315
         Left            =   5880
         TabIndex        =   31
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateTo 
         Height          =   315
         Left            =   6720
         TabIndex        =   32
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
         Caption         =   "Allo. Ref"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   5640
         TabIndex        =   104
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   8040
         TabIndex        =   103
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Database"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "T9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   10080
         TabIndex        =   99
         Top             =   1695
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   12000
         TabIndex        =   98
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   10080
         TabIndex        =   95
         Top             =   1335
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   12000
         TabIndex        =   94
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   10080
         TabIndex        =   91
         Top             =   975
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   12000
         TabIndex        =   90
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   10080
         TabIndex        =   87
         Top             =   615
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   12000
         TabIndex        =   86
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   10080
         TabIndex        =   83
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   12000
         TabIndex        =   82
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   5880
         TabIndex        =   79
         Top             =   2055
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   8040
         TabIndex        =   78
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   5880
         TabIndex        =   75
         Top             =   1695
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   8040
         TabIndex        =   74
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5880
         TabIndex        =   71
         Top             =   1335
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   8040
         TabIndex        =   70
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   5880
         TabIndex        =   67
         Top             =   975
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   8040
         TabIndex        =   66
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Journal Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   63
         Top             =   2055
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   62
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "DR/CR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   59
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Allocated"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   58
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   57
         Top             =   1695
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Trans. Ref."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   56
         Top             =   1695
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8040
         TabIndex        =   55
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   54
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label38 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   53
         Top             =   975
         Width           =   135
      End
      Begin VB.Label Label38 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   52
         Top             =   975
         Width           =   135
      End
      Begin VB.Label Label38 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   51
         Top             =   975
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6600
         TabIndex        =   45
         Top             =   2565
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Bordx Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   44
         Top             =   2565
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   43
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Bordx Month"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   42
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Bordx Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   41
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label35 
         Caption         =   "OP Batch No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   40
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   39
         Top             =   615
         Width           =   375
      End
      Begin VB.Label Label30 
         Caption         =   "A/C Period Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   38
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label34 
         Caption         =   "A/C Period Month"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   37
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   36
         Top             =   1335
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Journal Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   1335
         Width           =   1455
      End
      Begin VB.Label Period 
         Caption         =   "Period"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   975
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Account Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   615
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lstSUNListing 
      Height          =   6435
      Left            =   120
      TabIndex        =   105
      Top             =   2760
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   11351
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
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
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   24
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Tran. Ref."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "DR/CR"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Alloc. Ind."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "A/C Period"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Description"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Journal Type"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Journal No."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Journal User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Line"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Alloc. Ref."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "Alloc. Date"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "T1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "T2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "T3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "T4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "T5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "T6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "T7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "T8"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "T9"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "  SUN System - Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   14805
   End
End
Attribute VB_Name = "frmSUNSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function SearchRecord()


Dim StrSql As String
Dim strAppend As String
Dim sql As String
Dim GrpCompanyName As String
Dim tmpPeriodFr As String
Dim tmpPeriodTo As String
    
    strAppend = ""


    If txtDB <> "" Then
        If Len(txtACCCodeFr) Then
            strAppend = strAppend + " and ACCNT_CODE >= '" & Trim(txtACCCodeFr) & "'"
        End If
            
        If Len(txtACCCodeTo) Then
            strAppend = strAppend + " and ACCNT_CODE <= '" & Trim(txtACCCodeTo) & "'"
        End If
        
        If Len(txtPeriodMthFr) Then
            If Len(txtPeriodMthFr) = 2 Then
                tmpPeriodFr = Trim(txtPeriodYrFr) & "0" & Trim(txtPeriodMthFr)
                strAppend = strAppend + " and PERIOD >= '" & tmpPeriodFr & "'"
            Else
                tmpPeriodFr = Trim(txtPeriodYrFr) & "00" & Trim(txtPeriodMthFr)
                strAppend = strAppend + " and Substring(PERIOD,7,1) >= '" & Trim(tmpPeriodFr) & "'"
            End If
        End If
            
        If Len(txtPeriodMthTo) Then
            If Len(txtPeriodMthTo) = 2 Then
                tmpPeriodTo = Trim(txtPeriodYrTo) & "0" & Trim(txtPeriodMthTo)
                strAppend = strAppend + " and PERIOD <= '" & tmpPeriodTo & "'"
            Else
                tmpPeriodTo = Trim(txtPeriodYrTo) & "00" & Trim(txtPeriodMthTo)
                strAppend = strAppend + " and PERIOD <= '" & Trim(tmpPeriodTo) & "'"
            End If
        End If
            
            
        'If Len(txtPeriodMthTo) Then
        '    If Len(txtPeriodMthFr) = 2 Then
        '        strAppend = strAppend + " and Substring(PERIOD,6,2) >= '" & Trim(txtPeriodMthTo) & "'"
        '    Else
        '        strAppend = strAppend + " and Substring(PERIOD,7,1) >= '" & Trim(txtPeriodMthTo) & "'"
        '    End If
        'End If
        
        'If Len(txtPeriodYrFr) Then
        '    strAppend = strAppend + " and Substring(PERIOD,1,4)  >= '" & Trim(txtPeriodYrFr) & "'"
        'End If
            
        'If Len(txtPeriodYrTo) Then
        '    strAppend = strAppend + " and Substring(PERIOD,1,4) <= '" & Trim(txtPeriodYrTo) & "'"
        'End If
        
        
        If Len(txtJournalNoFr) Then
            strAppend = strAppend + " and JRNAL_NO >= '" & Trim(txtJournalNoFr) & "'"
        End If
            
        If Len(txtJournalNoTo) Then
            strAppend = strAppend + " and JRNAL_NO <= '" & Trim(txtJournalNoTo) & "'"
        End If
       
        If Len(txtAmtFr) Then
            strAppend = strAppend + " and AMOUNT >= '" & Trim(txtAmtFr) & "'"
        End If
            
        If Len(txtAmtTo) Then
            strAppend = strAppend + " and AMOUNT <= '" & Trim(txtAmtTo) & "'"
        End If
            
        If Len(txtTranRefFr) Then
            strAppend = strAppend + " and TREFERENCE >= '" & Trim(txtTranRefFr) & "'"
        End If
            
        If Len(txtTranRefTo) Then
            strAppend = strAppend + " and TREFERENCE <= '" & Trim(txtTranRefTo) & "'"
        End If
        
       If Len(txtJournalTypeFr) Then
            strAppend = strAppend + " and JRNAL_TYPE >= '" & Trim(txtJournalTypeFr) & "'"
        End If
            
        If Len(txtJournalTypeTo) Then
            strAppend = strAppend + " and JRNAL_TYPE <= '" & Trim(txtJournalTypeTo) & "'"
        End If
        
        
       If Len(txtAllocRefFr) Then
            strAppend = strAppend + " and ALLOC_REF >= '" & Trim(txtAllocRefFr) & "'"
        End If
            
        If Len(txtAllocRefTo) Then
            strAppend = strAppend + " and ALLOC_REF <= '" & Trim(txtAllocRefTo) & "'"
        End If
    
        If Len(txtAllocated) Then
            If Trim(txtAllocated) = "Y" Then
                strAppend = strAppend + " and ALLOCATION = 'A'"
            ElseIf Trim(txtAllocated) = "N" Then
                strAppend = strAppend + " and ALLOCATION = ''"
            End If
        End If
           
        If Len(txtDRCR) Then
            strAppend = strAppend + " and D_C = '" & Trim(txtDRCR) & "'"
        End If
    
        If Len(txtT1Fr) Then
            strAppend = strAppend + " and ANAL_T1 >= '" & Trim(txtT1Fr) & "'"
        End If
            
        If Len(txtT1To) Then
            strAppend = strAppend + " and ANAL_T1 <= '" & Trim(txtT1To) & "'"
        End If
        
        If Len(txtT2Fr) Then
            strAppend = strAppend + " and ANAL_T2 >= '" & Trim(txtT2Fr) & "'"
        End If
            
        If Len(txtT2To) Then
            strAppend = strAppend + " and ANAL_T2 <= '" & Trim(txtT2To) & "'"
        End If
    
        If Len(txtT3Fr) Then
            strAppend = strAppend + " and ANAL_T3 >= '" & Trim(txtT3Fr) & "'"
        End If
            
        If Len(txtT3To) Then
            strAppend = strAppend + " and ANAL_T3 <= '" & Trim(txtT3To) & "'"
        End If
    
        If Len(txtT4Fr) Then
            strAppend = strAppend + " and ANAL_T4 >= '" & Trim(txtT4Fr) & "'"
        End If
            
        If Len(txtT4To) Then
            strAppend = strAppend + " and ANAL_T4 <= '" & Trim(txtT4To) & "'"
        End If
        
        If Len(txtT5Fr) Then
            strAppend = strAppend + " and ANAL_T5 >= '" & Trim(txtT5Fr) & "'"
        End If
            
        If Len(txtT5To) Then
            strAppend = strAppend + " and ANAL_T5 <= '" & Trim(txtT5To) & "'"
        End If
        
        If Len(txtT6Fr) Then
            strAppend = strAppend + " and ANAL_T6 >= '" & Trim(txtT6Fr) & "'"
        End If
            
        If Len(txtT6To) Then
            strAppend = strAppend + " and ANAL_T6 <= '" & Trim(txtT6To) & "'"
        End If
        
        If Len(txtT7Fr) Then
            strAppend = strAppend + " and ANAL_T7 >= '" & Trim(txtT7Fr) & "'"
        End If
            
        If Len(txtT7To) Then
            strAppend = strAppend + " and ANAL_T7 <= '" & Trim(txtT7To) & "'"
        End If
        
        If Len(txtT8Fr) Then
            strAppend = strAppend + " and ANAL_T8 >= '" & Trim(txtT8Fr) & "'"
        End If
            
        If Len(txtT8To) Then
            strAppend = strAppend + " and ANAL_T8 <= '" & Trim(txtT8To) & "'"
        End If
        
        If Len(txtT9Fr) Then
            strAppend = strAppend + " and ANAL_T9 >= '" & Trim(txtT9Fr) & "'"
        End If
            
        If Len(txtT9To) Then
            strAppend = strAppend + " and ANAL_T9 <= '" & Trim(txtT9To) & "'"
        End If
        
        Call SUNListing(strAppend)
    Else
        MsgBox "Please Enter database"
        txtDB.SetFocus
    End If
    'debugs:
    'If Err.Description <> "" Then
    '    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
    'End If
    
   
End Function

Private Sub SUNListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
Dim sql2 As String
Dim rst3 As New ADODB.Recordset
Dim tmpAccName As String

    total = 0
    Current = 0
    lstSUNListing.ListItems.Clear
    
    If Trim(txtDB) = "UHC" Then
        sql = "select * from SALFLDGUHC where 0 = 0 "
    ElseIf Trim(txtDB) = "FW" Then
        sql = "select * from SALFLDGFW where 0 = 0 "
    ElseIf Trim(txtDB) = "MED" Then
        sql = "select * from SALFLDGMED where 0 = 0 "
    ElseIf Trim(txtDB) = "GV" Then
        sql = "select * from SALFLDGGV where 0 = 0 "
    ElseIf Trim(txtDB) = "TL" Then
        sql = "select * from SALFLDGTL where 0 = 0 "
    ElseIf Trim(txtDB) = "UHI" Then
        sql = "select * from SALFLDGUHI where 0 = 0 "
    End If
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdClear.Enabled = True
        'CmdToPayor.Enabled = True
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
        Do Until rst2.EOF
            With rst2
                Set Record = lstSUNListing.ListItems.Add(, , IIf(Len(!ACCNT_CODE), Trim(!ACCNT_CODE), ""))
                
                If Trim(txtDB) = "UHC" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'UHC' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                ElseIf Trim(txtDB) = "FW" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'FW' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                ElseIf Trim(txtDB) = "MED" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'MED' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                ElseIf Trim(txtDB) = "GV" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'GV' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                ElseIf Trim(txtDB) = "TL" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'TL' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                ElseIf Trim(txtDB) = "UHI" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'UHI' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                End If
                    
                rst3.Open sql2, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
                    If rst3.EOF = False Then
                        tmpAccName = IIf(Len(rst3!ACCNT_NAME), rst3!ACCNT_NAME, "")
                    End If
                rst3.Close
                Set rst3 = Nothing
                
                If Len(tmpAccName) Then
                    Record.SubItems(1) = Trim(tmpAccName)
                End If
                
                
                If Len(!TRANS_DATE) Then
                    Record.SubItems(2) = Mid(!TRANS_DATE, 7, 2) & "/" & Mid(!TRANS_DATE, 5, 2) & "/" & Mid(!TRANS_DATE, 1, 4)
                End If
                If Len(!TREFERENCE) Then
                    Record.SubItems(3) = Trim(!TREFERENCE)
                End If
                If Len(!Amount) Then
                    'If !Amount < 0 Then
                    '    Record.SubItems(3) = !Amount * -1
                    'Else
                        Record.SubItems(4) = !Amount
                    'End If
                End If
                If Len(!D_C) Then
                    Record.SubItems(5) = !D_C
                End If
                If Len(!Allocation) Then
                    Record.SubItems(6) = !Allocation
                End If
                If Len(!Period) Then
                   Record.SubItems(7) = !Period
                End If
                If Len(!DESCRIPTN) Then
                    Record.SubItems(8) = !DESCRIPTN
                End If
                If Len(!JRNAL_TYPE) Then
                    Record.SubItems(9) = !JRNAL_TYPE
                End If
                
                If Len(!JRNAL_NO) Then
                    Record.SubItems(10) = !JRNAL_NO
                End If
        
                If Len(!JRNAL_SRCE) Then
                   Record.SubItems(11) = !JRNAL_SRCE
                End If
                
                If Len(!JRNAL_LINE) Then
                   Record.SubItems(12) = !JRNAL_LINE
                End If
                If Len(!ALLOC_REF) Then
                    Record.SubItems(13) = !ALLOC_REF
                End If
                If Len(!ALLOC_DATE) Then
                    If Trim(!ALLOC_DATE) <> "0" Then
                        Record.SubItems(14) = Mid(!ALLOC_DATE, 7, 2) & "/" & Mid(!ALLOC_DATE, 5, 2) & "/" & Mid(!ALLOC_DATE, 1, 4)
                    End If
                End If
                If Len(!ANAL_T1) Then
                    Record.SubItems(15) = !ANAL_T1
                End If
                If Len(!ANAL_T2) Then
                    Record.SubItems(16) = !ANAL_T2
                End If
                If Len(!ANAL_T3) Then
                    Record.SubItems(17) = !ANAL_T3
                End If
                If Len(!ANAL_T4) Then
                    Record.SubItems(18) = !ANAL_T4
                End If
                If Len(!ANAL_T5) Then
                    Record.SubItems(19) = !ANAL_T5
                End If
                
                If Len(!ANAL_T6) Then
                    Record.SubItems(20) = !ANAL_T6
                End If
                
                If Len(!ANAL_T7) Then
                    Record.SubItems(21) = !ANAL_T7
                End If
                If Len(!ANAL_T8) Then
                    Record.SubItems(22) = !ANAL_T8
                End If
            
                If Len(!ANAL_T9) Then
                    Record.SubItems(23) = !ANAL_T9
                End If
            End With
            rst2.MoveNext
        Loop
    End If

    rst2.Close
    Set rst2 = Nothing
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
    
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelBordxToPayor(lstSUNListing)
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSearch_Click()
    Screen.MousePointer = vbHourglass
        medixmasterSUNSys.Open "File Name=" & BOSPath & "Conn\MediConnSUNSystem.UDL"
            Call SearchRecord
        medixmasterSUNSys.Close
        Set medixmasterSUNSys = Nothing
    Screen.MousePointer = Default
End Sub

Private Sub Form_Load()
    txtDB.Text = "UHC"
End Sub

Private Sub txtACCCodeFr_LostFocus()
    txtACCCodeFr = UCase(txtACCCodeFr)
End Sub


Private Sub txtACCCodeTo_LostFocus()
    txtACCCodeTo = UCase(txtACCCodeTo)
End Sub

Private Sub txtAllocated_LostFocus()
    If Len(txtAllocated) Then
        txtAllocated = UCase(txtAllocated)
        If Trim(txtAllocated) <> "Y" And Trim(txtAllocated) <> "N" Then
            txtAllocated.SetFocus
            MsgBox "Please Enter Y or N"
        End If
    End If
End Sub

Private Sub txtAmtFr_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAmtTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtDB_LostFocus()
    If Len(txtDB) Then
        txtDB = UCase(txtDB)
        If Trim(txtDB) <> "FW" And Trim(txtDB) <> "MED" And Trim(txtDB) <> "GV" And Trim(txtDB) <> "TL" And Trim(txtDB) <> "UHI" And Trim(txtDB) <> "UHC" Then
            txtDB.SetFocus
            MsgBox "Please Enter a valid database"
        End If
    End If
End Sub

Private Sub txtDRCR_LostFocus()
    If Len(txtDRCR) Then
        txtDRCR = UCase(txtDRCR)
        If Trim(txtDRCR) <> "C" And Trim(txtDRCR) <> "D" Then
            txtDRCR.SetFocus
            MsgBox "Please Enter C or R"
        End If
    End If
End Sub

Private Sub txtJournalNoFr_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtJournalNoTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtJournalTypeFr_LostFocus()
    If Len(txtJournalTypeFr) Then
        txtJournalTypeFr = UCase(txtJournalTypeFr)
    End If
End Sub

Private Sub txtJournalTypeTo_LostFocus()
    If Len(txtJournalTypeTo) Then
        txtJournalTypeTo = UCase(txtJournalTypeTo)
    End If
End Sub

Private Sub txtPeriodMthFr_Change()
    If Len(txtPeriodMthFr) = 2 Then
        txtPeriodYrFr.SetFocus
    End If
End Sub

Private Sub txtPeriodMthFr_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPeriodMthTo_Change()
    If Len(txtPeriodMthTo) = 2 Then
        txtPeriodYrTo.SetFocus
    End If
End Sub

Private Sub txtPeriodMthTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPeriodYrFr_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPeriodYrTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtT1Fr_LostFocus()
    txtT1Fr = UCase(txtT1Fr)
End Sub

Private Sub txtT1To_LostFocus()
    txtT1To = UCase(txtT1To)
End Sub

Private Sub txtT2Fr_LostFocus()
    txtT2Fr = UCase(txtT2Fr)
End Sub

Private Sub txtT2To_LostFocus()
    txtT2To = UCase(txtT2To)
End Sub

Private Sub txtT3Fr_LostFocus()
    txtT3Fr = UCase(txtT3Fr)
End Sub

Private Sub txtT3To_LostFocus()
    txtT3To = UCase(txtT3To)
End Sub

Private Sub txtT4Fr_LostFocus()
    txtT4Fr = UCase(txtT4Fr)
End Sub

Private Sub txtT4To_LostFocus()
    txtT4To = UCase(txtT4To)
End Sub

Private Sub txtT5Fr_LostFocus()
    txtT5Fr = UCase(txtT5Fr)
End Sub

Private Sub txtT5To_LostFocus()
    txtT5To = UCase(txtT5To)
End Sub

Private Sub txtT6Fr_LostFocus()
    txtT6Fr = UCase(txtT6Fr)
End Sub

Private Sub txtT6To_LostFocus()
    txtT6To = UCase(txtT6To)
End Sub

Private Sub txtT7Fr_LostFocus()
    txtT7Fr = UCase(txtT7Fr)
End Sub

Private Sub txtT7To_LostFocus()
    txtT7To = UCase(txtT7To)
End Sub

Private Sub txtT8Fr_LostFocus()
    txtT8Fr = UCase(txtT8Fr)
End Sub

Private Sub txtT8To_LostFocus()
    txtT8To = UCase(txtT8To)
End Sub

Private Sub txtT9Fr_LostFocus()
    txtT9Fr = UCase(txtT9Fr)
End Sub

Private Sub txtT9To_LostFocus()
    txtT9To = UCase(txtT9To)
End Sub

Private Sub txtTranRefFr_LostFocus()
    txtTranRefFr = UCase(txtTranRefFr)
End Sub

Private Sub txtTranRefTo_LostFocus()
    txtTranRefTo = UCase(txtTranRefTo)
End Sub

