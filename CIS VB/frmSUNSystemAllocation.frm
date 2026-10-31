VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSUNSystemAllocation 
   Caption         =   "SUN System - Allocation"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   19008
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9600
   ScaleWidth      =   19008
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   0
      TabIndex        =   52
      Top             =   240
      Width           =   18855
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10920
         MaxLength       =   50
         TabIndex        =   131
         Top             =   2040
         Width           =   2865
      End
      Begin VB.TextBox txtAllocRefTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   15
         Top             =   240
         Width           =   1545
      End
      Begin VB.TextBox txtAllocRefFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   14
         Top             =   240
         Width           =   1545
      End
      Begin VB.TextBox txtDB 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
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
         Width           =   1905
      End
      Begin VB.TextBox txtT9To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   35
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtT9Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   34
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtT8To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   33
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtT8Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   32
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtT7To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   31
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtT7Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   30
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtT6To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   29
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtT6Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   28
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtT5To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12240
         MaxLength       =   15
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtT5Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtT4To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   25
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtT4Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   24
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtT3To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   23
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtT3Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   22
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtT2To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   21
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtT2Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   20
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtT1To 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   19
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtT1Fr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtJournalTypeTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         MaxLength       =   15
         TabIndex        =   13
         Top             =   2040
         Width           =   1785
      End
      Begin VB.TextBox txtJournalTypeFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   12
         Top             =   2040
         Width           =   1905
      End
      Begin VB.TextBox txtDRCR 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtAllocated 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         MaxLength       =   15
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtTranRefFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1680
         Width           =   1905
      End
      Begin VB.TextBox txtTranRefTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1680
         Width           =   1785
      End
      Begin VB.TextBox txtAmtFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   16
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtAmtTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   17
         Top             =   600
         Width           =   1545
      End
      Begin VB.TextBox txtPeriodYrTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         MaxLength       =   4
         TabIndex        =   7
         Top             =   960
         Width           =   960
      End
      Begin VB.TextBox txtPeriodYrFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   5
         Top             =   960
         Width           =   1080
      End
      Begin VB.TextBox txtPeriodMthFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtPeriodMthTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         MaxLength       =   2
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox ChkBatchNo 
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
         Left            =   7440
         TabIndex        =   78
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkPeriodMth 
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
         Left            =   6720
         TabIndex        =   77
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox ChkPeriodYr 
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
         Left            =   7920
         TabIndex        =   76
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox ChkGrpComp 
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
         Left            =   6720
         TabIndex        =   75
         Top             =   3240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtACCCodeFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   2
         Top             =   600
         Width           =   1665
      End
      Begin VB.TextBox txtACCCodeTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         MaxLength       =   30
         TabIndex        =   74
         Top             =   3360
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.TextBox txtOPBatchNo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7680
         MaxLength       =   20
         TabIndex        =   73
         Top             =   3360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtOPBatchNo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8160
         MaxLength       =   20
         TabIndex        =   72
         Top             =   3360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox CboPayor 
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
         Left            =   8040
         TabIndex        =   71
         Top             =   3360
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox txtJournalNoFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1320
         Width           =   1905
      End
      Begin VB.TextBox txtJournalNoTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1320
         Width           =   1785
      End
      Begin VB.ComboBox cboACPeriodMthFr 
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
         ItemData        =   "frmSUNSystemAllocation.frx":0000
         Left            =   7680
         List            =   "frmSUNSystemAllocation.frx":0002
         TabIndex        =   70
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodYRFr 
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
         ItemData        =   "frmSUNSystemAllocation.frx":0004
         Left            =   7200
         List            =   "frmSUNSystemAllocation.frx":0006
         TabIndex        =   69
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodMthTo 
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
         ItemData        =   "frmSUNSystemAllocation.frx":0008
         Left            =   7920
         List            =   "frmSUNSystemAllocation.frx":000A
         TabIndex        =   68
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodYRTo 
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
         ItemData        =   "frmSUNSystemAllocation.frx":000C
         Left            =   7680
         List            =   "frmSUNSystemAllocation.frx":000E
         TabIndex        =   67
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboBordxMonthfr 
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
         Left            =   4920
         TabIndex        =   66
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboBordxMonthTo 
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
         Left            =   7080
         TabIndex        =   65
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboBordxYearfr 
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
         Left            =   5040
         TabIndex        =   64
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboBordxYearTo 
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
         Left            =   6840
         TabIndex        =   63
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkAppendRcd 
         Caption         =   "Append Records"
         Height          =   255
         Left            =   14160
         TabIndex        =   62
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cmdLookup 
         Caption         =   ">"
         Height          =   255
         Left            =   3000
         TabIndex        =   61
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtImportFilename 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   14160
         TabIndex        =   60
         Top             =   480
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.CommandButton cmdFileLocation 
         Caption         =   " >"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   16440
         TabIndex        =   59
         Top             =   480
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton cmpUpload 
         Caption         =   "Upload Listing"
         Height          =   375
         Left            =   16920
         TabIndex        =   58
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAccountname 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   14160
         MaxLength       =   15
         TabIndex        =   57
         Top             =   1680
         Width           =   4305
      End
      Begin VB.CommandButton cmdUploadDirect 
         Caption         =   "Upload SUN"
         Height          =   375
         Left            =   16920
         TabIndex        =   56
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdFileLocationSUN 
         Caption         =   " >"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   16440
         TabIndex        =   55
         Top             =   1080
         Width           =   435
      End
      Begin VB.TextBox txtImportFilenameSUN 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   14160
         TabIndex        =   54
         Top             =   1080
         Width           =   2265
      End
      Begin VB.CheckBox ChkScan 
         Height          =   255
         Left            =   5520
         TabIndex        =   53
         Top             =   1680
         Width           =   255
      End
      Begin MSMask.MaskEdBox txtBordxDateFr 
         Height          =   315
         Left            =   6000
         TabIndex        =   79
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
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
         Left            =   6600
         TabIndex        =   80
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   18120
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label15 
         Caption         =   "Description"
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
         Index           =   17
         Left            =   10080
         TabIndex        =   132
         Top             =   2055
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Allo. Ref"
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
         Index           =   16
         Left            =   5640
         TabIndex        =   128
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Index           =   13
         Left            =   8040
         TabIndex        =   127
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Database"
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
         Index           =   15
         Left            =   120
         TabIndex        =   126
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "T9"
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
         Index           =   14
         Left            =   10080
         TabIndex        =   125
         Top             =   1695
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Index           =   12
         Left            =   12000
         TabIndex        =   124
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T8"
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
         Index           =   13
         Left            =   10080
         TabIndex        =   123
         Top             =   1335
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Index           =   11
         Left            =   12000
         TabIndex        =   122
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T7"
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
         Index           =   12
         Left            =   10080
         TabIndex        =   121
         Top             =   975
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Index           =   10
         Left            =   12000
         TabIndex        =   120
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T6"
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
         Index           =   11
         Left            =   10080
         TabIndex        =   119
         Top             =   615
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Index           =   9
         Left            =   12000
         TabIndex        =   118
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T5"
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
         Index           =   10
         Left            =   10080
         TabIndex        =   117
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Index           =   8
         Left            =   12000
         TabIndex        =   116
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T4"
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
         Index           =   9
         Left            =   5880
         TabIndex        =   115
         Top             =   2055
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Index           =   7
         Left            =   8040
         TabIndex        =   114
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T3"
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
         Index           =   8
         Left            =   5880
         TabIndex        =   113
         Top             =   1695
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Index           =   6
         Left            =   8040
         TabIndex        =   112
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T2"
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
         Index           =   7
         Left            =   5880
         TabIndex        =   111
         Top             =   1335
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Index           =   5
         Left            =   8040
         TabIndex        =   110
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "T1"
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
         Index           =   6
         Left            =   5880
         TabIndex        =   109
         Top             =   975
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Index           =   4
         Left            =   8040
         TabIndex        =   108
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Journal Type"
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
         Index           =   5
         Left            =   120
         TabIndex        =   107
         Top             =   2055
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Index           =   3
         Left            =   3360
         TabIndex        =   106
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "DR/CR"
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
         Index           =   4
         Left            =   3360
         TabIndex        =   105
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Allocated"
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
         Index           =   3
         Left            =   3360
         TabIndex        =   104
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Index           =   2
         Left            =   3360
         TabIndex        =   103
         Top             =   1695
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Trans. Ref."
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
         Index           =   2
         Left            =   120
         TabIndex        =   102
         Top             =   1695
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Index           =   1
         Left            =   8040
         TabIndex        =   101
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Amount"
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
         Index           =   1
         Left            =   5640
         TabIndex        =   100
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label38 
         Caption         =   "/"
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
         Index           =   3
         Left            =   4320
         TabIndex        =   99
         Top             =   975
         Width           =   135
      End
      Begin VB.Label Label38 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   98
         Top             =   975
         Width           =   135
      End
      Begin VB.Label Label38 
         Caption         =   "To"
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
         Index           =   1
         Left            =   3360
         TabIndex        =   97
         Top             =   975
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "To"
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
         Index           =   4
         Left            =   6720
         TabIndex        =   96
         Top             =   3240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Bordx Year"
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
         Index           =   3
         Left            =   3480
         TabIndex        =   95
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "To"
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
         Index           =   2
         Left            =   6600
         TabIndex        =   94
         Top             =   3240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Bordx Month"
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
         Index           =   1
         Left            =   3600
         TabIndex        =   93
         Top             =   3360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Bordx Date"
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
         Index           =   1
         Left            =   5880
         TabIndex        =   92
         Top             =   3240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label35 
         Caption         =   "OP Batch No"
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
         Left            =   7800
         TabIndex        =   91
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "To"
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
         Index           =   0
         Left            =   4800
         TabIndex        =   90
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label30 
         Caption         =   "A/C Period Year"
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
         Left            =   6240
         TabIndex        =   89
         Top             =   3240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label34 
         Caption         =   "A/C Period Month"
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
         Left            =   5640
         TabIndex        =   88
         Top             =   3360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Index           =   0
         Left            =   3360
         TabIndex        =   87
         Top             =   1335
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Journal Number"
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
         Index           =   0
         Left            =   120
         TabIndex        =   86
         Top             =   1335
         Width           =   1455
      End
      Begin VB.Label Period 
         Caption         =   "Period"
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
         TabIndex        =   85
         Top             =   975
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Account Code"
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
         TabIndex        =   84
         Top             =   615
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Filename (Listing)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   14160
         TabIndex        =   83
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Account Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14160
         TabIndex        =   82
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Filename (SUN)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   14160
         TabIndex        =   81
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdCheck 
      Caption         =   "Check All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   51
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   120
      TabIndex        =   36
      Top             =   8880
      Width           =   18855
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   18000
         TabIndex        =   40
         Top             =   255
         Width           =   735
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   16320
         TabIndex        =   39
         Top             =   255
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   15480
         TabIndex        =   38
         Top             =   255
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Allocate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   17160
         TabIndex        =   37
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "DR :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblDebit 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   49
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "CR :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   48
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblCrebit 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   47
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblTotalCR 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10320
         TabIndex        =   46
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Total CR :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9360
         TabIndex        =   45
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTotalDR 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   44
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Total DR :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   43
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblBalance 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   42
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Balance :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   41
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstSUNListing 
      Height          =   5715
      Left            =   120
      TabIndex        =   129
      Top             =   3120
      Width           =   18855
      _ExtentX        =   33253
      _ExtentY        =   10075
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
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   25
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account Code"
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
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "DR/CR"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Alloc. Ind."
         Object.Width           =   2558
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "A/C Period"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Description"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Journal Type"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
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
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Account Name"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "  SUN System - Allocation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   0
      TabIndex        =   130
      Top             =   0
      Width           =   19245
   End
End
Attribute VB_Name = "frmSUNSystemAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_blnDirection(1 To 28) As Boolean
Dim xl As New Excel.Application
Dim xlsheet As Excel.Worksheet
Dim xlwbook As Excel.Workbook
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim tmpPathName As String
Dim Check As Boolean
Dim listloop1 As Integer

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
            strAppend = strAppend + " and ACCNT_CODE = '" & Trim(txtACCCodeFr) & "'"
        End If
            
        'If Len(txtACCCodeTo) Then
        '    strAppend = strAppend + " and ACCNT_CODE <= '" & Trim(txtACCCodeTo) & "'"
        'End If
        
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
        
        If ChkScan.Value = 1 Then
            strAppend = strAppend + " and TREFERENCE = '" & Trim(txtTranRefFr) & "'"
        Else
            If Len(txtTranRefFr) Then
                strAppend = strAppend + " and TREFERENCE >= '" & Trim(txtTranRefFr) & "'"
            End If
                
            If Len(txtTranRefTo) Then
                strAppend = strAppend + " and TREFERENCE <= '" & Trim(txtTranRefTo) & "'"
            End If
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
            ElseIf Trim(txtAllocated) = "R" Then
                strAppend = strAppend + " and ALLOCATION = 'R'"
            ElseIf Trim(txtAllocated) = "0" Then
                strAppend = strAppend + " and ALLOCATION = '0'"
            ElseIf Trim(txtAllocated) = "1" Then
                strAppend = strAppend + " and ALLOCATION = '1'"
            ElseIf Trim(txtAllocated) = "2" Then
                strAppend = strAppend + " and ALLOCATION = '2'"
            ElseIf Trim(txtAllocated) = "3" Then
                strAppend = strAppend + " and ALLOCATION = '3'"
            ElseIf Trim(txtAllocated) = "4" Then
                strAppend = strAppend + " and ALLOCATION = '4'"
            ElseIf Trim(txtAllocated) = "5" Then
                strAppend = strAppend + " and ALLOCATION = '5'"
            ElseIf Trim(txtAllocated) = "6" Then
                strAppend = strAppend + " and ALLOCATION = '6'"
            ElseIf Trim(txtAllocated) = "7" Then
                strAppend = strAppend + " and ALLOCATION = '7'"
            ElseIf Trim(txtAllocated) = "8" Then
                strAppend = strAppend + " and ALLOCATION = '8'"
            ElseIf Trim(txtAllocated) = "9" Then
                strAppend = strAppend + " and ALLOCATION = '9'"
            ElseIf Trim(txtAllocated) = "U" Then
                strAppend = strAppend + " and ALLOCATION = '0' OR ALLOCATION = '1' OR ALLOCATION = '2' OR ALLOCATION = '3' OR ALLOCATION = '4' OR ALLOCATION = '5' OR ALLOCATION = '6' OR ALLOCATION = '7' OR ALLOCATION = '8' OR ALLOCATION = '9' OR ALLOCATION = ''  OR ALLOCATION = 'R'"
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
        
        If Len(txtDescription) Then
            strAppend = strAppend + " and DESCRIPTN = '" & Trim(txtDescription) & "'"
        End If
        
        
        
        lblDebit = 0
        lblCrebit = 0
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

Private Function SearchRecordSUNDB(tmpTREFERENCE As String)
Dim StrSql As String
Dim strAppend As String
Dim sql As String
Dim GrpCompanyName As String
Dim tmpPeriodFr As String
Dim tmpPeriodTo As String
    
    strAppend = ""


    If txtDB <> "" Then
        If Len(txtACCCodeFr) Then
            strAppend = strAppend + " and ACCNT_CODE = '" & Trim(txtACCCodeFr) & "'"
        End If
            
        'If Len(txtTranRefFr) Then
        '    strAppend = strAppend + " and TREFERENCE >= '" & Trim(txtTranRefFr) & "'"
        'End If
            
        'If Len(txtTranRefTo) Then
            strAppend = strAppend + " and TREFERENCE = '" & Trim(tmpTREFERENCE) & "'"
        'End If
            
        If Len(txtAllocated) Then
            If Trim(txtAllocated) = "Y" Then
                strAppend = strAppend + " and ALLOCATION = 'A'"
            ElseIf Trim(txtAllocated) = "N" Then
                strAppend = strAppend + " and ALLOCATION = ''"
            End If
        End If
           
        lblDebit = 0
        lblCrebit = 0
        Call SUNListingDirect(strAppend)
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
    lblTotalCR = 0
    lblTotalDR = 0
    
    If chkAppendRcd.Value = 0 Then
        If ChkScan.Value = 0 Then
            lstSUNListing.ListItems.Clear
        End If
    End If
    
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
    ElseIf Trim(txtDB) = "MS" Then
        sql = "select * from SALFLDGMS where 0 = 0 "
    End If
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdClear.Enabled = True
        'CmdToPayor.Enabled = True
        CmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
        Do Until rst2.EOF
            With rst2
                Set Record = lstSUNListing.ListItems.Add(, , "")
                
                'If Len(!Allocation) Then
                '    Record.SubItems(1) = !Allocation
                'End If
                If Len(!ACCNT_CODE) Then
                    Record.SubItems(1) = !ACCNT_CODE
                End If
                If Len(!TRANS_DATE) Then
                    Record.SubItems(2) = Mid(!TRANS_DATE, 7, 2) & "/" & Mid(!TRANS_DATE, 5, 2) & "/" & Mid(!TRANS_DATE, 1, 4)
                End If
                If Len(!TREFERENCE) Then
                    Record.SubItems(3) = Trim(!TREFERENCE)
                End If
                If Len(!Amount) Then
                    'If !Amount < 0 Then
                     '   Record.SubItems(4) = !Amount * -1
                    'Else
                        Record.SubItems(4) = Format(!Amount, "#,##0.00")
                    'End If
                    If !D_C = "C" Then
                        lblTotalCR = lblTotalCR + !Amount
                        lblTotalCR = Format(lblTotalCR, "#,##0.00")
                    Else
                        lblTotalDR = lblTotalDR + !Amount
                        lblTotalDR = Format(lblTotalDR, "#,##0.00")
                    End If
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
                ElseIf Trim(txtDB) = "MS" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'MS' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                End If
                    
                rst3.Open sql2, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
                    If rst3.EOF = False Then
                        tmpAccName = IIf(Len(rst3!ACCNT_NAME), rst3!ACCNT_NAME, "")
                    End If
                rst3.Close
                Set rst3 = Nothing
                
                If Len(tmpAccName) Then
                    Record.SubItems(24) = Trim(tmpAccName)
                End If
            
            End With
            rst2.MoveNext
        Loop
    End If

    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdExit.Enabled = True
    
End Sub
Private Sub SUNListingDirect(Optional strAppend As String)
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
    lblTotalCR = 0
    lblTotalDR = 0
    

    
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
    ElseIf Trim(txtDB) = "MS" Then
        sql = "select * from SALFLDGMS where 0 = 0 "
    End If
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
    
    If rst2.EOF = True Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdClear.Enabled = True
        'CmdToPayor.Enabled = True
        CmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
        Do Until rst2.EOF
            With rst2
                Set Record = lstSUNListing.ListItems.Add(, , "")
                
                'If Len(!Allocation) Then
                '    Record.SubItems(1) = !Allocation
                'End If
                If Len(!ACCNT_CODE) Then
                    Record.SubItems(1) = !ACCNT_CODE
                End If
                If Len(!TRANS_DATE) Then
                    Record.SubItems(2) = Mid(!TRANS_DATE, 7, 2) & "/" & Mid(!TRANS_DATE, 5, 2) & "/" & Mid(!TRANS_DATE, 1, 4)
                End If
                If Len(!TREFERENCE) Then
                    Record.SubItems(3) = Trim(!TREFERENCE)
                End If
                If Len(!Amount) Then
                    'If !Amount < 0 Then
                     '   Record.SubItems(4) = !Amount * -1
                    'Else
                        Record.SubItems(4) = Format(!Amount, "#,##0.00")
                    'End If
                    If !D_C = "C" Then
                        lblTotalCR = lblTotalCR + !Amount
                        lblTotalCR = Format(lblTotalCR, "#,##0.00")
                    Else
                        lblTotalDR = lblTotalDR + !Amount
                        lblTotalDR = Format(lblTotalDR, "#,##0.00")
                    End If
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
                ElseIf Trim(txtDB) = "MS" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'MS' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                End If
                    
                rst3.Open sql2, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
                    If rst3.EOF = False Then
                        tmpAccName = IIf(Len(rst3!ACCNT_NAME), rst3!ACCNT_NAME, "")
                    End If
                rst3.Close
                Set rst3 = Nothing
                
                If Len(tmpAccName) Then
                    Record.SubItems(24) = Trim(tmpAccName)
                End If
            
            End With
            rst2.MoveNext
        Loop
    End If

    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdExit.Enabled = True
    
End Sub

Private Sub chkCheckAll_Click()

End Sub

Private Sub ChkScan_Click()
    If ChkScan.Value = 1 Then
        Label16(2).Visible = False
        txtTranRefTo.Visible = False
        ChkScan.Left = 3280
        txtTranRefFr = ""
        txtTranRefTo = ""
        txtTranRefFr.SetFocus
        lstSUNListing.ListItems.Clear
    Else
        Label16(2).Visible = True
        txtTranRefTo.Visible = True
        ChkScan.Left = 5520
        txtTranRefFr.SetFocus
    End If
End Sub

Private Sub CmdCheck_Click()
    If lstSUNListing.ListItems.Count <> 0 Then
        If CmdCheck.Caption = "Check All" Then
            SetCheckAllItems True
            CmdCheck.Caption = "Uncheck All"
        Else
             SetCheckAllItems False
            CmdCheck.Caption = "Check All"
        End If
    Else
        MsgBox "Please select a record.", vbInformation
    End If
End Sub

Private Sub SetCheckAllItems(bState As Boolean)
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim r As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstSUNListing.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstSUNListing.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFileLocation_Click()
  
    If lstSUNListing.ListItems.Count > 0 Then
        CommonDialog1.CancelError = True
        On Error GoTo ErrHandler
        CommonDialog1.InitDir = "C:\"
        CommonDialog1.Flags = cdlOFNHideReadOnly
        CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Dat Files (*.Dat)|*.Dat"
        CommonDialog1.ShowOpen
        txtImportFilename = CommonDialog1.FileTitle
        Exit Sub
    Else
        MsgBox "Please select record!", vbOKOnly + vbCritical
    End If
  
ErrHandler:
  Exit Sub
End Sub

Private Sub cmdFileLocationSUN_Click()

    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.InitDir = "C:\"
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Dat Files (*.Dat)|*.Dat"
    CommonDialog1.ShowOpen
    txtImportFilenameSUN = CommonDialog1.FileTitle
    Exit Sub
  
ErrHandler:
  Exit Sub
End Sub

Private Sub cmdLookup_Click()
'SUN_DB
'ACCNT_CODE
'ACCNT_NAME
    'frmSearchICD.Source = 11
    If txtDB <> "" Then
        frmSUNSystemSearchAcc.Show vbModal
    Else
        MsgBox "Please enter DB Name", vbOKOnly + vbCritical
        txtDB.SetFocus
    End If
End Sub

Private Sub CmdNext_Click()

End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelBordxToPayor(lstSUNListing)
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSearch_Click()

    Screen.MousePointer = vbHourglass
        
    Dim listrecord As Integer
        
        medixmasterSUNSys.Open "File Name=" & BOSPath & "Conn\MediConnSUNSystem.UDL"
            
            If ChkScan.Value = 1 Then
                If txtTranRefFr = "" Then
                    MsgBox "Please Enter Transaction Ref", vbOKOnly + vbCritical, DialogBoxTitle
                    txtTranRefFr.SetFocus
                Else
                    listrecord = lstSUNListing.ListItems.Count
                    If listrecord <> 0 Then
                        
                        For listloop1 = 1 To listrecord
                            Call CaseNoCheck
                        Next listloop1
                    End If
                        If Check = True Then
                            MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                        ElseIf Check = False Then
                            Call SearchRecord
                        End If
                        Check = False
                    Screen.MousePointer = Default
                    txtTranRefFr = ""
                    txtTranRefFr.SetFocus
                End If
                txtTranRefFr.SetFocus
            Else
                Call SearchRecord
            End If
            
            'Call SearchRecord
        medixmasterSUNSys.Close
        Set medixmasterSUNSys = Nothing
    Screen.MousePointer = Default
End Sub

Private Function CaseNoCheck()
    If Trim(ChkStr(txtTranRefFr)) = Trim(lstSUNListing.ListItems(listloop1).SubItems(3)) Then
        Check = True
    End If
End Function

Private Sub cmdUpdate_Click()
Const METHOD_NAME = "CmdSave_Click"
'On Error GoTo debugs

'On Err GoTo ErrorSave
Dim PAYCode As String
Dim CaseNo As String
Dim PVStatus As String
Dim ii As Integer
Dim i As Integer
Dim Cnt As Integer
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
Dim tmpAmtPay2C As Double
Dim PVNumber As String
Dim PrevCLN, CurrCLN As String
Dim DiffCLN As Boolean
Dim CLNCnt As Integer
Dim strlsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim strsqlpymt As String
Dim pymt As New ADODB.Recordset
Dim PYMTSTATUS As Boolean
Dim pymtcnt As Integer
Dim chqcode As String
Dim tmpchqno As String
Dim tmpPVNoFr As String
Dim tmpPVNoto As String
Dim startTrans As Boolean
Dim CAQ As New ADODB.Recordset
Dim StrSql As String
Dim CAQNumber As String
Dim tmpCAQNumber As String

    Screen.MousePointer = vbHourglass
    'On Error GoTo HANDLE_ERROR
    tmpPVNoFr = ""
    tmpPVNoto = ""
    pymtcnt = 0
    PYMTSTATUS = False
    Cnt = 0
    ItemChecked = False
    TmpPVCode = ""
    
    For i = 1 To lstSUNListing.ListItems.Count
        If lstSUNListing.ListItems(i).Checked = True Then
            ItemChecked = True
            Exit For
        End If
    Next i
    CLNCnt = 0
    DiffCLN = False

    If ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    Else
        update = MsgBox("Do you want to update the record ? ", vbYesNo + vbExclamation)
        If update = vbYes Then
        
            medixmasterSUNSys.Open "File Name=" & BOSPath & "Conn\MediConnSUNSystem.UDL"
    
            StrSql = "Select * from Tbl_Alloc_Ref"
            CAQ.Open StrSql, medixmasterSUNSys, adOpenStatic, adLockOptimistic
           
                CAQ!ALLOC_REF = Format(CAQ!ALLOC_REF + 1, "000000")
                tmpCAQNumber = CAQ!ALLOC_REF
                CAQ.update
       
            CAQ.Close
            Set CAQ = Nothing

            For i = 1 To lstSUNListing.ListItems.Count
                If lstSUNListing.ListItems(i).Checked = True Then
                    Call UpdateRecord(Trim(lstSUNListing.ListItems(i).SubItems(3)), Trim(lstSUNListing.ListItems(i).SubItems(4)), Trim(lstSUNListing.ListItems(i).SubItems(1)), tmpCAQNumber, Trim(lstSUNListing.ListItems(i).SubItems(12)), Trim(lstSUNListing.ListItems(i).SubItems(10)))
                    lstSUNListing.ListItems(i).Checked = False
                    lstSUNListing.ListItems(i).Text = ""
                    lstSUNListing.ListItems(i).SubItems(6) = "A"
                    lstSUNListing.ListItems(i).SubItems(13) = tmpCAQNumber
                    lstSUNListing.ListItems(i).SubItems(14) = Date
                End If
            Next i
            lblCrebit = 0
            lblDebit = 0
            medixmasterSUNSys.Close
            Set medixmasterSUNSys = Nothing
            CmdCheck.Caption = "Check All"
            MsgBox "Record Allocated...", vbOKOnly + vbInformation, DialogBoxTitle
        
        End If
    End If
    
    Screen.MousePointer = Default
    'Exit Function
'ErrorSave:
'    Screen.MousePointer = vbDefault
'        If startTrans = True Then
'            MediConnCISPayment.RollbackTrans
'        End If
'
'    If Err.Description <> "" Then
'        Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
'    End If
'
'debugs:
'        Screen.MousePointer = vbDefault
'        If startTrans = True Then
'            MediConnCISPayment.RollbackTrans
'        End If
'        If Err.Description <> "" Then
'            Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
'        End If

End Sub

Private Sub UpdateRecord(CaseNo As String, PaymentAmt As Double, AccCode As String, AllocRef As String, JournalLine As String, JournalNo As String)

Const METHOD_NAME = "UpdateRecord"
On Error GoTo debugs

Dim update
Dim StrSql As String
Dim PAY As New ADODB.Recordset
Dim strsqlPay As String
Dim CLM As New ADODB.Recordset
Dim strsqlCLM As String
Dim ZZ As String
Dim TempBank As String
Dim tmpUpdWording1 As String
Dim tmpUpdWording2 As String
Dim strupdatepymt As String
Dim tmpPeriodFr As String
    tmpPeriodFr = ""
    
        If Month(Date) = "1" Then
            tmpPeriodFr = Year(Date) & "001"
        ElseIf Month(Date) = "2" Then
            tmpPeriodFr = Year(Date) & "002"
        ElseIf Month(Date) = "3" Then
            tmpPeriodFr = Year(Date) & "003"
        ElseIf Month(Date) = "4" Then
            tmpPeriodFr = Year(Date) & "004"
        ElseIf Month(Date) = "5" Then
            tmpPeriodFr = Year(Date) & "005"
        ElseIf Month(Date) = "6" Then
            tmpPeriodFr = Year(Date) & "006"
        ElseIf Month(Date) = "7" Then
            tmpPeriodFr = Year(Date) & "007"
        ElseIf Month(Date) = "8" Then
            tmpPeriodFr = Year(Date) & "008"
        ElseIf Month(Date) = "9" Then
            tmpPeriodFr = Year(Date) & "009"
        ElseIf Month(Date) = "10" Then
            tmpPeriodFr = Year(Date) & "010"
        ElseIf Month(Date) = "11" Then
            tmpPeriodFr = Year(Date) & "011"
        ElseIf Month(Date) = "12" Then
            tmpPeriodFr = Year(Date) & "012"
        End If
        'PAY.Close
        'Set PAY = Nothing
        ' strsqlINV = "select  * from ClaimInvoice where  INVPayTo = 'H' And CASNumber = '" & Trim(txtCASNumber) & "'"
        ' INV.Open strsqlINV, medixmasterconnWS, adOpenKeyset, adLockOptimistic



    strsqlPay = ""
    
    
    If Trim(txtDB) = "UHC" Then
        strsqlPay = " Select * from SALFLDGUHC Where TREFERENCE = '" & Trim(CaseNo) & "' And AMOUNT = '" & PaymentAmt & "'  And ACCNT_CODE = '" & AccCode & "' And JRNAL_LINE = '" & JournalLine & "' And JRNAL_NO = '" & JournalNo & "'"
    ElseIf Trim(txtDB) = "FW" Then
        strsqlPay = " Select * from SALFLDGFW Where TREFERENCE = '" & Trim(CaseNo) & "' And AMOUNT = '" & PaymentAmt & "'  And ACCNT_CODE = '" & AccCode & "' And JRNAL_LINE = '" & JournalLine & "' And JRNAL_NO = '" & JournalNo & "'"
    ElseIf Trim(txtDB) = "MED" Then
        strsqlPay = " Select * from SALFLDGMED Where TREFERENCE = '" & Trim(CaseNo) & "' And AMOUNT = '" & PaymentAmt & "'  And ACCNT_CODE = '" & AccCode & "' And JRNAL_LINE = '" & JournalLine & "' And JRNAL_NO = '" & JournalNo & "'"
    ElseIf Trim(txtDB) = "GV" Then
        strsqlPay = " Select * from SALFLDGGV Where TREFERENCE = '" & Trim(CaseNo) & "' And AMOUNT = '" & PaymentAmt & "'  And ACCNT_CODE = '" & AccCode & "' And JRNAL_LINE = '" & JournalLine & "' And JRNAL_NO = '" & JournalNo & "'"
    ElseIf Trim(txtDB) = "TL" Then
        strsqlPay = " Select * from SALFLDGTL Where TREFERENCE = '" & Trim(CaseNo) & "' And AMOUNT = '" & PaymentAmt & "'  And ACCNT_CODE = '" & AccCode & "' And JRNAL_LINE = '" & JournalLine & "' And JRNAL_NO = '" & JournalNo & "'"
    ElseIf Trim(txtDB) = "UHI" Then
        strsqlPay = " Select * from SALFLDGUHI Where TREFERENCE = '" & Trim(CaseNo) & "' And AMOUNT = '" & PaymentAmt & "'  And ACCNT_CODE = '" & AccCode & "' And JRNAL_LINE = '" & JournalLine & "' And JRNAL_NO = '" & JournalNo & "'"
    ElseIf Trim(txtDB) = "MS" Then
        strsqlPay = " Select * from SALFLDGMS Where TREFERENCE = '" & Trim(CaseNo) & "' And AMOUNT = '" & PaymentAmt & "'  And ACCNT_CODE = '" & AccCode & "' And JRNAL_LINE = '" & JournalLine & "' And JRNAL_NO = '" & JournalNo & "'"
    End If
    PAY.Open strsqlPay, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
    
    If PAY.RecordCount > 0 Then
    'Else
    '    PAY.AddNew
    'End If
    With PAY
        !ALLOC_REF = AllocRef
        !Allocation = "A"
        !ALLOC_DATE = Format(Date, "yyyymmdd")
        !ALLOC_PERIOD = tmpPeriodFr
        
    End With
    PAY.update
    
    End If
    PAY.Close
    Set PAY = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
 
End Sub


Private Sub cmdUploadDirect_Click()
Dim AccNo As String
Dim TransRef As String
Dim AppAmt As String
Dim AccLine As String
Dim tmpJournalNo As String
Dim commitTrans As Boolean
Dim Val1 As Integer
Dim error As Boolean
Dim Record As ListItem
Dim lSize As Long
Dim FileExists As Integer
Dim SendDate As String
Dim tmpDebitAmt As Double
Dim tmpCreditAmt As Double
    
    Screen.MousePointer = vbHourglass
        lstSUNListing.ListItems.Clear
        ChkScan.Value = 0
    If txtDB <> "" And txtAllocated <> "" And txtACCCodeFr <> "" Then
        If txtImportFilenameSUN <> "" Then
                Val1 = 2
                error = False
                cmpUpload.Enabled = False
                FileExists = 0
            
                lSize = -1
                lSize = FileLen(CommonDialog1.Filename)
                
                If lSize <= 0 Then
                    FileExists = 0
                ElseIf lSize > 0 Then
                    FileExists = 1
                End If
                
                lblDebit = 0
                lblCrebit = 0
                'FileLen (dlgOpen.Filename)
                txtImportFilenameSUN = CommonDialog1.Filename
                If FileExists = 1 Then
                    
        'lSize = -1
        'lSize = FileLen(CommonDialog1.Filename)
        'If lSize <= 0 Then
        '    FileExists = 0
        'ElseIf lSize > 0 Then
        '    FileExists = 1
        'End If
         '   'FileLen (dlgOpen.Filename)
        
       '     If FileExists = 1 Then
       '         Set xlwbook = xl.Workbooks.Open(CommonDialog1.Filename)
       '         Set xlsheet = xlwbook.Sheets.Item(1)
            
                
                    
                    medixmasterSUNSys.Open "File Name=" & BOSPath & "Conn\MediConnSUNSystem.UDL"
                    
                    'Set xlwbook = xl.Workbooks.Open(txtImportFilenameSUN)
                    Set xlwbook = xl.Workbooks.Open(CommonDialog1.Filename)
                    Set xlsheet = xlwbook.Sheets.Item(1)
                    
                    Do Until Val1 = 60000
                        If xlsheet.Cells(Val1, 1) <> "" Then
                            error = False
                            TransRef = xlsheet.Cells(Val1, 1)
                            
                            Call SearchRecordSUNDB(TransRef)
                        Else
                            medixmasterSUNSys.Close
                            Set medixmasterSUNSys = Nothing
                            Screen.MousePointer = vbDefault
                            cmpUpload.Enabled = True
                            Exit Sub
                        End If
                        Val1 = Val1 + 1
                    Loop
                Else
                    Screen.MousePointer = vbDefault
                    cmpUpload.Enabled = True
                    MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
                
                End If
        Else
            Screen.MousePointer = vbDefault
            cmpUpload.Enabled = True
            MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
    
        End If
    
        
        Else
            MsgBox "Please enter Database, Account Code and Allocation Indicator!", vbOKOnly + vbCritical, DialogBoxTitle
            Screen.MousePointer = vbDefault
        End If

    'Screen.MousePointer = vbDefault

End Sub

Private Sub cmpUpload_Click()
Dim AccNo As String
Dim TransRef As String
Dim AppAmt As String
Dim AccLine As String
Dim tmpJournalNo As String
Dim commitTrans As Boolean
Dim Val1 As Integer
Dim error As Boolean
Dim Record As ListItem
Dim lSize As Long
Dim FileExists As Integer
Dim SendDate As String
Dim tmpDebitAmt As Double
Dim tmpCreditAmt As Double
    
    Screen.MousePointer = vbHourglass
    If txtImportFilename <> "" Then
            Val1 = 2
            error = False
            cmpUpload.Enabled = False
            FileExists = 0
        
            lSize = -1
            lSize = FileLen(CommonDialog1.Filename)
            If lSize <= 0 Then
                FileExists = 0
            ElseIf lSize > 0 Then
                FileExists = 1
            End If
            lblDebit = 0
            lblCrebit = 0
            'FileLen (dlgOpen.Filename)
            txtImportFilename = CommonDialog1.Filename
            If FileExists = 1 Then
                Set xlwbook = xl.Workbooks.Open(txtImportFilename)
                Set xlsheet = xlwbook.Sheets.Item(1)
                
                Do Until Val1 = 60000
                    If xlsheet.Cells(Val1, 1) <> "" Then
                        error = False
                        AccNo = xlsheet.Cells(Val1, 1)
                        TransRef = xlsheet.Cells(Val1, 2)
                        AppAmt = Format(xlsheet.Cells(Val1, 3), "#,##0.00")
                        AppAmt = Format(Trim(AppAmt), "#,###0.00")
                        tmpJournalNo = xlsheet.Cells(Val1, 4)
                        AccLine = xlsheet.Cells(Val1, 5)

                            Dim objItem As ListItem
                            Dim dAcctNo As String
                            Dim dTransRef As String
                            Dim dAppAmt As String
                            Dim dAccLine As String
                            Dim dJournalNo As String
                            For Each objItem In lstSUNListing.ListItems
                                dAcctNo = Trim(objItem.SubItems(1))
                                dTransRef = Trim(objItem.SubItems(3))
                                dAppAmt = Trim(Format(objItem.SubItems(4), "#,##0.00"))
                                dAccLine = Trim(objItem.SubItems(12))
                                dJournalNo = Trim(objItem.SubItems(10))
                                
                                'Format(Trim(AppAmt), "#,##0.00"))
                                If (Trim(AccNo) = dAcctNo And Trim(TransRef) = dTransRef And Trim(AppAmt) = dAppAmt And Trim(AccLine) = dAccLine And dJournalNo = Trim(tmpJournalNo)) Then
                                
                                    
                                    If Trim(objItem.SubItems(6)) = "" Then
                                        objItem.Checked = True
                                        objItem.Text = "Allocate"
                                        If objItem.SubItems(5) = "D" Then
                                            lblDebit = CDbl(lblDebit) + CDbl(objItem.SubItems(4))
                                        Else
                                            lblCrebit = CDbl(lblCrebit) + CDbl(objItem.SubItems(4))
                                        End If
                                        If lblDebit <> 0 And lblCrebit <> 0 Then
                                             tmpDebitAmt = IIf(lblDebit < 0, lblDebit * -1, lblDebit)
                                             tmpCreditAmt = IIf(lblCrebit < 0, lblCrebit * -1, lblCrebit)
                                             
                                             lblBalance = Format(tmpDebitAmt, "#,##0.00") - Format(tmpCreditAmt, "#,##0.00")
                                             lblBalance = Format(lblBalance, "#,##0.00")
                                             If tmpDebitAmt = tmpCreditAmt Then
                                                 cmdUpdate.Enabled = True
                                             Else
                                                 cmdUpdate.Enabled = False
                                             End If
                                         Else
                                         
                                            cmdUpdate.Enabled = False
                                             tmpDebitAmt = IIf(lblDebit < 0, lblDebit * -1, lblDebit)
                                             tmpCreditAmt = IIf(lblCrebit < 0, lblCrebit * -1, lblCrebit)
                                             lblBalance = tmpDebitAmt - tmpCreditAmt
                                             lblBalance = Format(lblBalance, "#,##0.00")
                                         End If
                                    Else
                                        MsgBox "Record has been allocated!", vbOKOnly + vbCritical, DialogBoxTitle
                                    End If
                                End If
                                                                
                            Next
                    Else
                        Screen.MousePointer = vbDefault
                        cmpUpload.Enabled = True
                        Exit Sub
                    End If
                    Val1 = Val1 + 1
                Loop
            Else
                Screen.MousePointer = vbDefault
                cmpUpload.Enabled = True
                MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
            End If
        Else
            Screen.MousePointer = vbDefault
            cmpUpload.Enabled = True
            MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    'Screen.MousePointer = vbDefault

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "frmSUNSystemAllocation")
    txtDB.Text = "UHC"
End Sub

Private Sub lstSUNListing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
   

    Select Case ColumnHeader.Index
    Case 1
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstSUNListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstSUNListing, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstSUNListing, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
    Case 15
        SortListView lstSUNListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(15)
    Case 16
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 19
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(19)
    Case 20
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(21)
    Case 22
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(22)
    Case 23
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(23)
    Case 24
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(24)
    Case 25
        SortListView lstSUNListing, ColumnHeader.Index, ldtString, m_blnDirection(24)

    End Select
End Sub

Private Sub lstSUNListing_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ProcessCheckedItem(Item)
End Sub





Private Sub txtACCCodeFr_KeyPress(KeyAscii As Integer)
    Dim rst2 As New ADODB.Recordset
    Dim ReasonMaster As ListItem
    Dim sql As String
    Dim status As Boolean
    
    If KeyAscii = 13 Then  ' The ENTER key.
        
        medixmasterSUNSys.Open "File Name=" & BOSPath & "Conn\MediConnSUNSystem.UDL"
        
        If Trim(frmSUNSystemAllocation.txtDB) = "UHC" Then
            sql = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'UHC' And ACCNT_CODE > '" & Trim(txtACCCodeFr) & "'"
        ElseIf Trim(frmSUNSystemAllocation.txtDB) = "FW" Then
            sql = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'FW' And ACCNT_CODE > '" & Trim(txtACCCodeFr) & "'"
        ElseIf Trim(frmSUNSystemAllocation.txtDB) = "MED" Then
            sql = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'MED' And ACCNT_CODE > '" & Trim(txtACCCodeFr) & "'"
        ElseIf Trim(frmSUNSystemAllocation.txtDB) = "GV" Then
            sql = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'GV' And ACCNT_CODE > '" & Trim(txtACCCodeFr) & "'"
        ElseIf Trim(frmSUNSystemAllocation.txtDB) = "TL" Then
            sql = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'TL' And ACCNT_CODE > '" & Trim(txtACCCodeFr) & "'"
        ElseIf Trim(frmSUNSystemAllocation.txtDB) = "UHI" Then
            sql = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'UHI' And ACCNT_CODE > '" & Trim(txtACCCodeFr) & "'"
        ElseIf Trim(frmSUNSystemAllocation.txtDB) = "MS" Then
            sql = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'MS' And ACCNT_CODE > '" & Trim(txtACCCodeFr) & "'"
        End If
        
        'If Len(Trim(strAppend)) Then
        '    sql = sql + strAppend
        'End If
    
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        status = False
        If rst2.RecordCount = 0 Then
            MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
        Else
            
            txtACCCodeFr = Trim(rst2!ACCNT_CODE)
            txtAccountname = Trim(rst2!ACCNT_NAME)
            status = True
            rst2.Close
            Set rst2 = Nothing
        End If
        medixmasterSUNSys.Close
        Set medixmasterSUNSys = Nothing
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        'sql = " select MBMNumber from MBM where MBMnumber > '" & RemStr(txtMBMNumber) & "' " & _
        '    " Order by MBMnumber "
        'MBM.MaxRecords = 1
        'MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
       '
        'If MBM.RecordCount Then
        '    MBMNumber = MBM!MBMNumber
        '    status = True
        'End If
        'MBM.Close
        'Set MBM = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
        If status = False Then
            MsgBox "Last Record Reached! " & vbNewLine & "Not a valid Membership Number Format! ", vbCritical + vbOKOnly
        End If
        KeyAscii = 0 ' Ignore this key.
    End If
End Sub

Private Sub txtACCCodeFr_LostFocus()
    txtACCCodeFr = UCase(txtACCCodeFr)
End Sub

Private Sub txtAllocated_LostFocus()
    If Len(txtAllocated) Then
        txtAllocated = UCase(txtAllocated)
        'If Trim(txtAllocated) <> "Y" And Trim(txtAllocated) <> "N" Then
        '    txtAllocated.SetFocus
        '    MsgBox "Please Enter Y or N"
        'End If
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
        If Trim(txtDB) <> "FW" And Trim(txtDB) <> "MED" And Trim(txtDB) <> "GV" And Trim(txtDB) <> "TL" And Trim(txtDB) <> "UHI" And Trim(txtDB) <> "UHC" And Trim(txtDB) <> "MS" Then
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
            MsgBox "Please Enter C or D"
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


Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
'Const METHOD_NAME = "ProcessCheckedItem"
'On Error GoTo debugs
Dim tmpAllocAmt As Double
Dim tmpDebitAmt As Double
Dim tmpCreditAmt As Double

    If Item.Checked = True Then
        
        If Trim(Item.SubItems(6)) <> "" Then
            MsgBox "Record has been allocated!", vbCritical
            Item.Checked = False
            tmpAllocAmt = 0
        Else
            'Item.SubItems(9) = CDbl(Item.SubItems(8))
            Item.Text = "Allocate"
            If Item.SubItems(5) = "D" Then
                lblDebit = lblDebit + CDbl(Format(Item.SubItems(4), "#,##0.00"))
                lblDebit = Format(lblDebit, "#,##0.00")
                tmpAllocAmt = Item.SubItems(4)
            Else
                lblCrebit = lblCrebit + CDbl(Format(Item.SubItems(4), "#,##0.00"))
                lblCrebit = Format(lblCrebit, "#,##0.00")
                tmpAllocAmt = Item.SubItems(4)
            End If
        End If

    Else
        'LblAmt = LblAmt - CDbl(Format(Item.SubItems(9), "#,##0.00"))
        'Item.SubItems(9) = ""
        Item.Text = ""
        'If tmpAllocAmt <> 0 Then
            If Item.SubItems(5) = "D" Then
                lblDebit = lblDebit - CDbl(Format(Item.SubItems(4), "#,##0.00"))
                lblDebit = Format(lblDebit, "#,##0.00")
            Else
                lblCrebit = lblCrebit - CDbl(Format(Item.SubItems(4), "#,##0.00"))
                lblCrebit = Format(lblCrebit, "#,##0.00")
            End If
        'End If
    End If
    
    If lblDebit <> 0 And lblCrebit <> 0 Then
        tmpDebitAmt = IIf(lblDebit < 0, lblDebit * -1, lblDebit)
        tmpCreditAmt = IIf(lblCrebit < 0, lblCrebit * -1, lblCrebit)
        
        lblBalance = Format(tmpDebitAmt, "#,##0.00") - Format(tmpCreditAmt, "#,##0.00")
        lblBalance = Format(lblBalance, "#,##0.00")
        If tmpDebitAmt = tmpCreditAmt Then
            cmdUpdate.Enabled = True
        Else
            cmdUpdate.Enabled = False
        End If
    Else
    
       cmdUpdate.Enabled = False
        tmpDebitAmt = IIf(lblDebit < 0, lblDebit * -1, lblDebit)
        tmpCreditAmt = IIf(lblCrebit < 0, lblCrebit * -1, lblCrebit)
        lblBalance = tmpDebitAmt - tmpCreditAmt
        lblBalance = Format(lblBalance, "#,##0.00")
    End If
'debugs:
'If ERR.Description <> "" Then
'    Call ErrorLog(Me.Caption, METHOD_NAME, ERR.Description)
'End If
End Function









