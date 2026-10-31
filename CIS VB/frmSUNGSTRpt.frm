VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSUNGSTRpt 
   Caption         =   "SUN GST Report"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1155
   ScaleWidth      =   9345
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   9135
      Begin VB.CommandButton cmdPrintSum 
         Caption         =   "Print Summary"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3120
         TabIndex        =   62
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6000
         TabIndex        =   61
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Print Full"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4920
         TabIndex        =   60
         Top             =   240
         Width           =   1095
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
         TabIndex        =   29
         Top             =   1920
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
         Left            =   5040
         TabIndex        =   28
         Top             =   1920
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
         TabIndex        =   27
         Top             =   1665
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
         Left            =   5040
         TabIndex        =   26
         Top             =   1665
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
         ItemData        =   "frmSUNGSTRpt.frx":0000
         Left            =   7680
         List            =   "frmSUNGSTRpt.frx":0002
         TabIndex        =   25
         Top             =   1680
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
         ItemData        =   "frmSUNGSTRpt.frx":0004
         Left            =   7920
         List            =   "frmSUNGSTRpt.frx":0006
         TabIndex        =   24
         Top             =   1680
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
         ItemData        =   "frmSUNGSTRpt.frx":0008
         Left            =   7320
         List            =   "frmSUNGSTRpt.frx":000A
         TabIndex        =   23
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
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
         ItemData        =   "frmSUNGSTRpt.frx":000C
         Left            =   7680
         List            =   "frmSUNGSTRpt.frx":000E
         TabIndex        =   22
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
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
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   21
         Top             =   1365
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   20
         Top             =   1365
         Visible         =   0   'False
         Width           =   1815
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
         Left            =   8400
         TabIndex        =   19
         Top             =   1650
         Visible         =   0   'False
         Width           =   3615
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
         Left            =   8400
         MaxLength       =   20
         TabIndex        =   18
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
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
         Left            =   8040
         MaxLength       =   20
         TabIndex        =   17
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
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
         Left            =   6600
         MaxLength       =   30
         TabIndex        =   16
         Top             =   1260
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   4440
         MaxLength       =   30
         TabIndex        =   15
         Top             =   1260
         Visible         =   0   'False
         Width           =   1785
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
         TabIndex        =   14
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
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
         Left            =   8040
         TabIndex        =   13
         Top             =   1680
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
         Left            =   7080
         TabIndex        =   12
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
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
         Left            =   7560
         TabIndex        =   11
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
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
         Left            =   3840
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtPeriodMthFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   960
         MaxLength       =   2
         TabIndex        =   0
         Top             =   240
         Width           =   720
      End
      Begin VB.TextBox txtPeriodYrFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   9
         Top             =   240
         Width           =   1080
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
         Left            =   4560
         MaxLength       =   30
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   960
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
         Left            =   12000
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1260
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1260
         Visible         =   0   'False
         Width           =   1815
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
         Left            =   9720
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   7560
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txtBordxDateFr 
         Height          =   315
         Left            =   6120
         TabIndex        =   30
         Top             =   1680
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
         Left            =   6960
         TabIndex        =   31
         Top             =   1680
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
      Begin VB.Label Label18 
         Caption         =   "Group Company"
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
         Left            =   6960
         TabIndex        =   57
         Top             =   1725
         Visible         =   0   'False
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
         Left            =   3000
         TabIndex        =   56
         Top             =   1275
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Period 
         Caption         =   "Period"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   320
         Width           =   855
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
         TabIndex        =   54
         Top             =   1395
         Visible         =   0   'False
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
         Index           =   0
         Left            =   3480
         TabIndex        =   53
         Top             =   1395
         Visible         =   0   'False
         Width           =   375
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
         Left            =   5760
         TabIndex        =   52
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label32 
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
         Left            =   8040
         TabIndex        =   51
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label31 
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
         Left            =   8640
         TabIndex        =   50
         Top             =   1680
         Visible         =   0   'False
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
         TabIndex        =   49
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   8280
         TabIndex        =   48
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
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
         Left            =   6360
         TabIndex        =   47
         Top             =   1275
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label4 
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
         Left            =   7680
         TabIndex        =   46
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
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
         TabIndex        =   45
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label19 
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
         Left            =   7200
         TabIndex        =   44
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
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
         TabIndex        =   43
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
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
         Left            =   3480
         TabIndex        =   42
         Top             =   1695
         Visible         =   0   'False
         Width           =   1455
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
         TabIndex        =   41
         Top             =   1695
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
         Left            =   3480
         TabIndex        =   40
         Top             =   1965
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
         Index           =   4
         Left            =   6600
         TabIndex        =   39
         Top             =   1965
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   3600
         TabIndex        =   38
         Top             =   1335
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label38 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   37
         Top             =   285
         Width           =   255
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
         Left            =   4440
         TabIndex        =   36
         Top             =   1320
         Visible         =   0   'False
         Width           =   135
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
         Left            =   8640
         TabIndex        =   35
         Top             =   1275
         Visible         =   0   'False
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
         Left            =   9480
         TabIndex        =   34
         Top             =   915
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Transaction Ref."
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
         Left            =   6360
         TabIndex        =   33
         Top             =   1335
         Visible         =   0   'False
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
         Index           =   2
         Left            =   9480
         TabIndex        =   32
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   9240
      Visible         =   0   'False
      Width           =   13335
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
         Left            =   11640
         TabIndex        =   2
         Top             =   255
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView lstSUNListing 
      Height          =   7275
      Left            =   120
      TabIndex        =   58
      Top             =   1920
      Visible         =   0   'False
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   12832
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Tran. Ref."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DR/CR"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Alloc. Ind."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "A/C Period"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Description"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Journal Type"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Journal No."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Line"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Alloc. Ref."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Alloc. Date"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "  SUN GST Report"
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
      TabIndex        =   59
      Top             =   0
      Width           =   9405
   End
End
Attribute VB_Name = "frmSUNGSTRpt"
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
    
    If txtPeriodMthFr <> "" And txtPeriodYrFr <> "" Then
        
            'tmpPeriodFr = Trim(txtPeriodYrFr) & "0" & Trim(txtPeriodMthFr)
            'strAppend = strAppend + " and PERIOD = '" & tmpPeriodFr & "'"
            'Call SUNListingSum(strAppend)
        
            tmpPeriodFr = Trim(txtPeriodYrFr) & "0" & Trim(txtPeriodMthFr)
            strAppend = strAppend + " and PERIOD = '" & tmpPeriodFr & "'"
            Call SUNListing(strAppend)
            
    Else
        MsgBox "Please Enter Accouting Period, vbCritical"
    End If
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
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim Cnt As Double
Dim tmpPeriodFr As String
Dim TotalAmt9 As Double
Dim tmpPeriodFr2 As String
Dim linecnt As Double
Dim TotalAmt1000 As Double
Dim TotalAmt1002 As Double
Dim TotalAmt1010 As Double
Dim TotalAmt1020 As Double
Dim TotalAmt1100 As Double
Dim TotalAmt1600 As Double
Dim TotalAmt8830 As Double
Dim TotalAmt8834 As Double
Dim TotalAmt8839 As Double
Dim Total8830CPAmt As Double
    TotalAmt1000 = 0
    TotalAmt1002 = 0
    TotalAmt1010 = 0
    TotalAmt1020 = 0
    TotalAmt1100 = 0
    TotalAmt1600 = 0
    TotalAmt8830 = 0
    TotalAmt8834 = 0
    TotalAmt8839 = 0
    
    linecnt = 1
    TotalAmt9 = 0
    TotalAmt9 = 0
    Cnt = 0
    total = 0
    Current = 0
    lstSUNListing.ListItems.Clear
    
    
        If txtPeriodMthFr = "01" Then
            tmpPeriodFr = Trim(txtPeriodYrFr - 1) & "012"
        ElseIf txtPeriodMthFr = "02" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "001"
        ElseIf txtPeriodMthFr = "03" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "002"
        ElseIf txtPeriodMthFr = "04" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "003"
        ElseIf txtPeriodMthFr = "05" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "004"
        ElseIf txtPeriodMthFr = "06" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "005"
        ElseIf txtPeriodMthFr = "07" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "006"
        ElseIf txtPeriodMthFr = "08" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "007"
        ElseIf txtPeriodMthFr = "09" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "008"
        ElseIf txtPeriodMthFr = "10" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "009"
        ElseIf txtPeriodMthFr = "11" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "010"
        ElseIf txtPeriodMthFr = "12" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "011"
        End If

    sql = "select * from VIEW_GST_P1 where 0 = 0 "
    
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
    
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "PaymentVoucher\SUNGSTRptV3.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        objWorksheet.Cells(2, "B") = Trim(txtPeriodMthFr) & "/" & Trim(txtPeriodYrFr)
        objWorksheet.Cells(5, "E") = Trim(txtPeriodMthFr) & "/" & Trim(txtPeriodYrFr)
        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "1000" Then
                objWorksheet.Cells(6, "E") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1002" Then
                objWorksheet.Cells(7, "E") = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1010" Then
                objWorksheet.Cells(8, "E") = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1020" Then
                objWorksheet.Cells(9, "E") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1100" Then
                objWorksheet.Cells(10, "E") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1600" Then
                objWorksheet.Cells(11, "E") = rst2!Amount
            End If
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing

        objWorksheet.Cells(5, "G") = Mid(tmpPeriodFr, 6, 2) & "/" & Mid(tmpPeriodFr, 1, 4)
        sql = "select * from VIEW_GST_P1 where 0 = 0 "
                
        'If Len(Trim(strAppend)) Then
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        'End If
        
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic

        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "1000" Then
                objWorksheet.Cells(6, "G") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1002" Then
                objWorksheet.Cells(7, "G") = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1010" Then
                objWorksheet.Cells(8, "G") = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1020" Then
                objWorksheet.Cells(9, "G") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1100" Then
                objWorksheet.Cells(10, "G") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1600" Then
                objWorksheet.Cells(11, "G") = rst2!Amount
            End If
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing

        sql = "select * from VIEW_GST_P2 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!JRNAL_TYPE) = "GJ" Then
                objWorksheet.Cells(19, "E") = rst2!Amount
            End If
            
            
            If Trim(rst2!JRNAL_TYPE) = "SI" Then
                objWorksheet.Cells(20, "E") = rst2!Amount
            End If
        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing


        sql = "select * from VIEW_GST_P2 where 0 = 0 "
        
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!JRNAL_TYPE) = "GJ" Then
                objWorksheet.Cells(19, "G") = rst2!Amount
            End If
            
            
            If Trim(rst2!JRNAL_TYPE) = "SI" Then
                objWorksheet.Cells(20, "G") = rst2!Amount
            End If
        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing

        sql = "select * from VIEW_GST_P3 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(24, "E") = rst2!Amount
            End If
            
            
            If Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(25, "E") = rst2!Amount
            End If
        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_P3 where 0 = 0 "
        
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(24, "G") = rst2!Amount
            End If
            
            
            If Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(25, "G") = rst2!Amount
            End If
        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_P4 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            If Trim(rst2!JRNAL_TYPE) = "CP" Then
                objWorksheet.Cells(29, "E") = rst2!Amount
            End If
            
            If Trim(rst2!JRNAL_TYPE) = "RC" Then
                objWorksheet.Cells(30, "E") = rst2!Amount
            End If
                        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        sql = "select * from VIEW_GST_P4 where 0 = 0 "
        
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            If Trim(rst2!JRNAL_TYPE) = "CP" Then
                objWorksheet.Cells(29, "G") = rst2!Amount
            End If
            
            If Trim(rst2!JRNAL_TYPE) = "RC" Then
                objWorksheet.Cells(30, "G") = rst2!Amount
            End If
                        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        sql = "select * from VIEW_GST_P5 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            If Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(34, "E") = rst2!Amount
            End If
            
            If Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(35, "E") = rst2!Amount
            End If
                        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        sql = "select * from VIEW_GST_P5 where 0 = 0 "
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            If Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(34, "G") = rst2!Amount
            End If
            
            If Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(35, "G") = rst2!Amount
            End If
                        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
    
        sql = "select * from VIEW_GST_P6 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!JRNAL_TYPE) = "CP" Then
                objWorksheet.Cells(39, "E") = rst2!Amount
            End If
                        
            If Trim(rst2!JRNAL_TYPE) = "RC" Then
                objWorksheet.Cells(40, "E") = rst2!Amount
            End If
            
            If Trim(rst2!JRNAL_TYPE) = "SI" Then
                objWorksheet.Cells(41, "E") = rst2!Amount
            End If
                       
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        sql = "select * from VIEW_GST_P6 where 0 = 0 "
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!JRNAL_TYPE) = "CP" Then
                objWorksheet.Cells(39, "G") = rst2!Amount
            End If
                        
            If Trim(rst2!JRNAL_TYPE) = "RC" Then
                objWorksheet.Cells(40, "G") = rst2!Amount
            End If
            
            If Trim(rst2!JRNAL_TYPE) = "SI" Then
                objWorksheet.Cells(41, "G") = rst2!Amount
            End If
                       
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_P7_OSR where 0 = 0 "
        If Len(Trim(strAppend)) Then
            sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            objWorksheet.Cells(45, "E") = rst2!Amount
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        If txtPeriodMthFr = "01" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr - 1) & "011"
        ElseIf txtPeriodMthFr = "02" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr - 1) & "012"
        ElseIf txtPeriodMthFr = "03" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "001"
        ElseIf txtPeriodMthFr = "04" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "002"
        ElseIf txtPeriodMthFr = "05" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "003"
        ElseIf txtPeriodMthFr = "06" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "004"
        ElseIf txtPeriodMthFr = "07" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "005"
        ElseIf txtPeriodMthFr = "08" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "006"
        ElseIf txtPeriodMthFr = "09" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "007"
        ElseIf txtPeriodMthFr = "10" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "008"
        ElseIf txtPeriodMthFr = "11" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "009"
        ElseIf txtPeriodMthFr = "12" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "010"
        End If

    
        sql = "select * from VIEW_GST_P7_OSR where 0 = 0 "
        sql = sql + " and PERIOD = '" & tmpPeriodFr2 & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            objWorksheet.Cells(45, "G") = rst2!Amount
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        sql = "select * from VIEW_GST_P7 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(47, "E") = rst2!Amount
            End If
                        
            If Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(48, "E") = rst2!Amount
            End If
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        sql = "select * from VIEW_GST_P7 where 0 = 0 "
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(47, "G") = rst2!Amount
            End If
                        
            If Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(48, "G") = rst2!Amount
            End If
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        sql = "select * from VIEW_GST_P9 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!JRNAL_TYPE) = "CP" And Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(52, "E") = rst2!Amount
                TotalAmt9 = TotalAmt9 + rst2!Amount
            End If
                        
            If Trim(rst2!JRNAL_TYPE) = "RC" And Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(53, "E") = rst2!Amount
                TotalAmt9 = TotalAmt9 + rst2!Amount
            End If
            
            If Trim(rst2!JRNAL_TYPE) = "SI" And Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(54, "E") = rst2!Amount
                TotalAmt9 = TotalAmt9 + rst2!Amount
            End If
            rst2.MoveNext
        Loop
            
                objWorksheet.Cells(56, "E") = TotalAmt9
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_P9 where 0 = 0 "
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        TotalAmt9 = 0
        Do Until rst2.EOF
            If Trim(rst2!JRNAL_TYPE) = "CP" And Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(52, "G") = rst2!Amount
                TotalAmt9 = TotalAmt9 + rst2!Amount
            End If
                        
            If Trim(rst2!JRNAL_TYPE) = "RC" And Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(53, "G") = rst2!Amount
                TotalAmt9 = TotalAmt9 + rst2!Amount
            End If
            
            If Trim(rst2!JRNAL_TYPE) = "SI" And Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(54, "G") = rst2!Amount
                TotalAmt9 = TotalAmt9 + rst2!Amount
            End If
            rst2.MoveNext
        Loop
            
                objWorksheet.Cells(56, "G") = TotalAmt9
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_P8 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "8830" Then
                objWorksheet.Cells(15, "E") = rst2!Amount
            End If
                        
            If Trim(rst2!ACCNT_CODE) = "8834" Then
                objWorksheet.Cells(16, "E") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "8839" Then
                objWorksheet.Cells(17, "E") = rst2!Amount
            End If
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_P8 where 0 = 0 "
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "8830" Then
                objWorksheet.Cells(15, "G") = rst2!Amount
            End If
                        
            If Trim(rst2!ACCNT_CODE) = "8834" Then
                objWorksheet.Cells(16, "G") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "8839" Then
                objWorksheet.Cells(17, "G") = rst2!Amount
            End If
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
                
        
        'Listing
        Set objWorksheet = objWorkbook.Sheets(4)
        Set objWorksheet = objWorkbook.Worksheets.Item(4)
   
        sql = "select * from VIEW_GST_1000_Listing where 0 = 0 "
   
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic

        Do Until rst2.EOF
    
            Cnt = CDbl(Cnt) + 1
            objWorksheet.Cells(1 + CDbl(Cnt), "A") = Trim(rst2!ACCNT_CODE)

            objWorksheet.Cells(1 + CDbl(Cnt), "B") = Trim(rst2!TREFERENCE)
 
            objWorksheet.Cells(1 + CDbl(Cnt), "C") = Trim(rst2!Amount)

            objWorksheet.Cells(1 + CDbl(Cnt), "D") = Trim(rst2!D_C)
            
            

            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        
        Cnt = 0
        sql = "select * from VIEW_GST_8830_Listing where 0 = 0 "
 
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
 
        Do Until rst2.EOF
            Cnt = CDbl(Cnt) + 1
            objWorksheet.Cells(1 + CDbl(Cnt), "F") = Trim(rst2!ACCNT_CODE)
 
            objWorksheet.Cells(1 + CDbl(Cnt), "G") = Trim(rst2!TREFERENCE)
 
            objWorksheet.Cells(1 + CDbl(Cnt), "H") = Trim(rst2!Amount)
  
            objWorksheet.Cells(1 + CDbl(Cnt), "I") = Trim(rst2!D_C)
            
            If objWorksheet.Cells(1 + CDbl(Cnt), "C") > 0 Then
                objWorksheet.Cells(1 + CDbl(Cnt), "J") = objWorksheet.Cells(1 + CDbl(Cnt), "H") / objWorksheet.Cells(1 + CDbl(Cnt), "C")
            End If
            
            If Round(objWorksheet.Cells(1 + CDbl(Cnt), "J") * 100, 0) <> "6" Then
                objWorksheet.Cells(1 + CDbl(Cnt), "J").Font.Bold = True
                objWorksheet.Cells(1 + CDbl(Cnt), "J").Font.Color = vbRed
            End If
            
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        Cnt = 0
        Set objWorksheet = objWorkbook.Sheets(5)
        Set objWorksheet = objWorkbook.Worksheets.Item(5)
   
        sql = "select * from VIEW_GST_1002_Listing where 0 = 0 "
   
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic

        Do Until rst2.EOF
    
            Cnt = CDbl(Cnt) + 1
            objWorksheet.Cells(1 + CDbl(Cnt), "A") = Trim(rst2!ACCNT_CODE)

            objWorksheet.Cells(1 + CDbl(Cnt), "B") = Trim(rst2!TREFERENCE)
 
            objWorksheet.Cells(1 + CDbl(Cnt), "C") = Trim(rst2!Amount)

            objWorksheet.Cells(1 + CDbl(Cnt), "D") = Trim(rst2!D_C)
            
            

            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        
        Cnt = 0
        sql = "select * from VIEW_GST_8830RC_Listing where 0 = 0 "
 
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
 
        Do Until rst2.EOF
            Cnt = CDbl(Cnt) + 1
            objWorksheet.Cells(1 + CDbl(Cnt), "F") = Trim(rst2!ACCNT_CODE)
 
            objWorksheet.Cells(1 + CDbl(Cnt), "G") = Trim(rst2!TREFERENCE)
 
            objWorksheet.Cells(1 + CDbl(Cnt), "H") = Trim(rst2!Amount)
  
            objWorksheet.Cells(1 + CDbl(Cnt), "I") = Trim(rst2!D_C)
            
            objWorksheet.Cells(1 + CDbl(Cnt), "J") = objWorksheet.Cells(1 + CDbl(Cnt), "H") / objWorksheet.Cells(1 + CDbl(Cnt), "C")

            If Round(objWorksheet.Cells(1 + CDbl(Cnt), "J") * 100, 0) <> "6" Then
                objWorksheet.Cells(1 + CDbl(Cnt), "J").Font.Bold = True
                objWorksheet.Cells(1 + CDbl(Cnt), "J").Font.Color = vbRed
            End If
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        
        
        
        
        
        Set objWorksheet = objWorkbook.Sheets(3)
        Set objWorksheet = objWorkbook.Worksheets.Item(3)
        Cnt = 0
        
        sql = "select * from VIEW_GST_ACC2kTO5999_Listing where 0 = 0 "
 
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
 
        Do Until rst2.EOF
            Cnt = CDbl(Cnt) + 1
            objWorksheet.Cells(1 + CDbl(Cnt), "A") = Trim(rst2!ACCNT_CODE)
 
            objWorksheet.Cells(1 + CDbl(Cnt), "B") = Trim(rst2!TREFERENCE)
 
            objWorksheet.Cells(1 + CDbl(Cnt), "C") = Trim(rst2!Amount)
  
            objWorksheet.Cells(1 + CDbl(Cnt), "D") = Trim(rst2!D_C)
            
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        
        'Set objWorksheet = objWorkbook.Sheets(5)
        'Set objWorksheet = objWorkbook.Worksheets.Item(5)
        Cnt = 0
        
        sql = "select * from VIEW_GST_ACC6810_Listing where 0 = 0 "
 
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
 
        Do Until rst2.EOF
            Cnt = CDbl(Cnt) + 1
            objWorksheet.Cells(1 + CDbl(Cnt), "F") = Trim(rst2!ACCNT_CODE)
 
            objWorksheet.Cells(1 + CDbl(Cnt), "G") = Trim(rst2!TREFERENCE)
 
            objWorksheet.Cells(1 + CDbl(Cnt), "H") = Trim(rst2!Amount)
  
            objWorksheet.Cells(1 + CDbl(Cnt), "I") = Trim(rst2!D_C)
            
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        
        
'Sheet 2
'Sheet 2
    Set objWorksheet = objWorkbook.Sheets(2)
    Set objWorksheet = objWorkbook.Worksheets.Item(2)

    sql = "select * from VIEW_GST_P1 where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic

        
        objWorksheet.Cells(2, "B") = Trim(txtPeriodMthFr) & "/" & Trim(txtPeriodYrFr)
        objWorksheet.Cells(5, "E") = Trim(txtPeriodMthFr) & "/" & Trim(txtPeriodYrFr)
        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "1000" Then
                objWorksheet.Cells(6, "E") = rst2!Amount
                TotalAmt1000 = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1002" Then
                objWorksheet.Cells(7, "E") = rst2!Amount
                TotalAmt1002 = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1010" Then
                objWorksheet.Cells(8, "E") = rst2!Amount
                TotalAmt1010 = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1020" Then
                objWorksheet.Cells(9, "E") = rst2!Amount
                TotalAmt1020 = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1100" Then
                objWorksheet.Cells(10, "E") = rst2!Amount
                TotalAmt1100 = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1600" Then
                objWorksheet.Cells(11, "E") = rst2!Amount
                TotalAmt1600 = rst2!Amount
            End If
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
       
       
        
        sql = "select * from VIEW_GST_P8 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "8830" Then
                objWorksheet.Cells(15, "E") = rst2!Amount
                TotalAmt8830 = rst2!Amount
            End If
                        
            If Trim(rst2!ACCNT_CODE) = "8834" Then
                objWorksheet.Cells(16, "E") = rst2!Amount
                TotalAmt8834 = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "8839" Then
                objWorksheet.Cells(17, "E") = rst2!Amount
                TotalAmt8839 = rst2!Amount
            End If
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_ACC6810_Summary where 0 = 0 "
        
            If Len(Trim(strAppend)) Then
                sql = sql + strAppend
            End If
            rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
            
            Do Until rst2.EOF
                objWorksheet.Cells(18, "E") = rst2!Amount
                rst2.MoveNext
            Loop
            rst2.Close
            Set rst2 = Nothing
       
        linecnt = 20
        TotalAmt = 0
        sql = "select * from VIEW_GST_P2 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            objWorksheet.Cells(linecnt, "A") = rst2!JRNAL_TYPE
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        'objWorksheet.Cells(7 + CDbl(PVCnt), "J").Borders.LineStyle = True
        'objWorksheet.Cells(7 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
         'ws.Cells(1, X).Font.Bold = True
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 1000 BY JT"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
        objWorksheet.Cells(linecnt, "E").Font.Bold = True
        
        If TotalAmt1000 = TotalAmt Then
            objWorksheet.Cells(linecnt, "F") = "Agree"
            objWorksheet.Cells(linecnt, "F").Font.Bold = True
        Else
            objWorksheet.Cells(linecnt, "F") = "Not Agree " & TotalAmt1000 - TotalAmt
        End If
        linecnt = CDbl(linecnt) + 2
        TotalAmt = 0
        
        sql = "select * from VIEW_GST_P3 where 0 = 0 "
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            objWorksheet.Cells(linecnt, "A") = IIf(Trim(rst2!ANAL_T7) = "", "EMPTY", rst2!ANAL_T7)
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 1000 BY T7"
        objWorksheet.Cells(linecnt, "E") = TotalAmt

        
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
        objWorksheet.Cells(linecnt, "E").Font.Bold = True
        
        If TotalAmt1000 = TotalAmt Then
            objWorksheet.Cells(linecnt, "F") = "Agree"
            objWorksheet.Cells(linecnt, "F").Font.Bold = True
        Else
            objWorksheet.Cells(linecnt, "F") = "Not Agree " & TotalAmt1000 - TotalAmt
        End If
        linecnt = CDbl(linecnt) + 2
        TotalAmt = 0
        
        sql = "select * from VIEW_GST_P4 where 0 = 0 "
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            objWorksheet.Cells(linecnt, "A") = rst2!JRNAL_TYPE
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        
        
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 1002 BY JT"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
        objWorksheet.Cells(linecnt, "E").Font.Bold = True
        
        If TotalAmt1002 = TotalAmt Then
            objWorksheet.Cells(linecnt, "F") = "Agree"
            objWorksheet.Cells(linecnt, "F").Font.Bold = True
        Else
            objWorksheet.Cells(linecnt, "F") = "Not Agree " & TotalAmt1002 - TotalAmt
        End If
        
        linecnt = CDbl(linecnt) + 2
        TotalAmt = 0
        
        sql = "select * from VIEW_GST_P5 where 0 = 0 "
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            objWorksheet.Cells(linecnt, "A") = IIf(Trim(rst2!ANAL_T7) = "", "EMPTY", rst2!ANAL_T7)
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 1002 BY T7"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        
        
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
        objWorksheet.Cells(linecnt, "E").Font.Bold = True
        
        If TotalAmt1002 = TotalAmt Then
            objWorksheet.Cells(linecnt, "F") = "Agree"
            objWorksheet.Cells(linecnt, "F").Font.Bold = True
        Else
            objWorksheet.Cells(linecnt, "F") = "Not Agree " & TotalAmt1002 - TotalAmt
        End If
        
        linecnt = CDbl(linecnt) + 2
        TotalAmt = 0
    
    
        sql = "select * from VIEW_GST_P6 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        Total8830CPAmt = 0
        Do Until rst2.EOF
        
            objWorksheet.Cells(linecnt, "A") = rst2!JRNAL_TYPE
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            If Trim(rst2!JRNAL_TYPE) = "CP" Then
                Total8830CPAmt = rst2!Amount
            End If
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 8830 BY JT"
        
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
        objWorksheet.Cells(linecnt, "E").Font.Bold = True
        
        If TotalAmt8830 = TotalAmt Then
            objWorksheet.Cells(linecnt, "F") = "Agree"
            objWorksheet.Cells(linecnt, "F").Font.Bold = True
        Else
            objWorksheet.Cells(linecnt, "F") = "Not Agree " & TotalAmt8830 - TotalAmt
        End If
        
        linecnt = CDbl(linecnt) + 2
        TotalAmt = 0
        
        
        sql = "select * from VIEW_GST_P7_OSR where 0 = 0 "
        If Len(Trim(strAppend)) Then
            sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            objWorksheet.Cells(linecnt, "A") = "AC 8830 BY T7 = OSR (Prior Month)"
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            If Total8830CPAmt = rst2!Amount Then
                objWorksheet.Cells(linecnt, "F") = "Agree"
                objWorksheet.Cells(linecnt, "F").Font.Bold = True
            Else
                objWorksheet.Cells(linecnt, "F").Font.Bold = True
                objWorksheet.Cells(linecnt, "F") = "Not Agree (" & Total8830CPAmt - rst2!Amount & ")"
            End If
            
            objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(linecnt, "E").Font.Bold = True

            linecnt = CDbl(linecnt) + 2
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        TotalAmt = 0
 
    
        sql = "select * from VIEW_GST_P7 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            objWorksheet.Cells(linecnt, "A") = IIf(Trim(rst2!ANAL_T7) = "", "EMPTY", rst2!ANAL_T7)
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 8830 BY T7"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
        objWorksheet.Cells(linecnt, "E").Font.Bold = True
        
        linecnt = CDbl(linecnt) + 2
        TotalAmt = 0
    
    
        sql = "select * from VIEW_GST_P9 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
        
            objWorksheet.Cells(linecnt, "A") = Trim(rst2!JRNAL_TYPE)
            objWorksheet.Cells(linecnt, "B") = IIf(Trim(rst2!ANAL_T7) = "", "EMPTY", rst2!ANAL_T7)
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 8830 BY T7"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
        objWorksheet.Cells(linecnt, "E").Font.Bold = True
        
        linecnt = CDbl(linecnt) + 2
        TotalAmt = 0
        
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
        
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        
        MsgBox "Report is ready!"
    
    End If



    Exit Sub
End Sub

Private Function SearchRecordSum()
Dim StrSql As String
Dim strAppend As String
Dim sql As String
Dim GrpCompanyName As String
Dim tmpPeriodFr As String
Dim tmpPeriodTo As String
    
    strAppend = ""
    
    If txtPeriodMthFr <> "" And txtPeriodYrFr <> "" Then
        
            tmpPeriodFr = Trim(txtPeriodYrFr) & "0" & Trim(txtPeriodMthFr)
            strAppend = strAppend + " and PERIOD = '" & tmpPeriodFr & "'"
            Call SUNListingSum(strAppend)
    Else
        MsgBox "Please Enter Accouting Period, vbCritical"
    End If
End Function

Private Sub SUNListingSum(Optional strAppend As String)
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
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim Cnt As Double
Dim tmpPeriodFr As String
Dim TotalAmt9 As Double
Dim tmpPeriodFr2 As String
Dim linecnt As Double
Dim TotalAmt1000 As Double
Dim TotalAmt1002 As Double
Dim TotalAmt1010 As Double
Dim TotalAmt1020 As Double
Dim TotalAmt1100 As Double
Dim TotalAmt1600 As Double
Dim TotalAmt8830 As Double
Dim TotalAmt8834 As Double
Dim TotalAmt8839 As Double
Dim Total8830CPAmt As Double

    TotalAmt1000 = 0
    TotalAmt1002 = 0
    TotalAmt1010 = 0
    TotalAmt1020 = 0
    TotalAmt1100 = 0
    TotalAmt1600 = 0
    TotalAmt8830 = 0
    TotalAmt8834 = 0
    TotalAmt8839 = 0
    
    linecnt = 1
    TotalAmt9 = 0
    Cnt = 0
    total = 0
    Current = 0
    lstSUNListing.ListItems.Clear
    
    
        'tmpPeriodFr = Trim(txtPeriodYrFr) & "0" & Trim(txtPeriodMthFr)
        'strAppend = strAppend + " and PERIOD = '" & tmpPeriodFr & "'"
        If txtPeriodMthFr = "01" Then
            tmpPeriodFr = Trim(txtPeriodYrFr - 1) & "012"
        ElseIf txtPeriodMthFr = "02" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "001"
        ElseIf txtPeriodMthFr = "03" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "002"
        ElseIf txtPeriodMthFr = "04" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "003"
        ElseIf txtPeriodMthFr = "05" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "004"
        ElseIf txtPeriodMthFr = "06" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "005"
        ElseIf txtPeriodMthFr = "07" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "006"
        ElseIf txtPeriodMthFr = "08" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "007"
        ElseIf txtPeriodMthFr = "09" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "008"
        ElseIf txtPeriodMthFr = "10" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "009"
        ElseIf txtPeriodMthFr = "11" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "010"
        ElseIf txtPeriodMthFr = "12" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "011"
        End If

    sql = "select * from VIEW_GST_P1 where 0 = 0 "
    
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
    
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "PaymentVoucher\SUNGSTRptV3.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        
        objWorksheet.Cells(2, "B") = Trim(txtPeriodMthFr) & "/" & Trim(txtPeriodYrFr)
        'objWorksheet.Cells(5, "E") = Trim(txtPeriodMthFr) & "/" & Trim(txtPeriodYrFr)
        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "1000" Then
                objWorksheet.Cells(6, "E") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1002" Then
                objWorksheet.Cells(7, "E") = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1010" Then
                objWorksheet.Cells(8, "E") = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1020" Then
                objWorksheet.Cells(9, "E") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1100" Then
                objWorksheet.Cells(10, "E") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1600" Then
                objWorksheet.Cells(11, "E") = rst2!Amount
            End If
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing

        sql = "select * from VIEW_GST_P1 where 0 = 0 "
                
        If Len(Trim(strAppend)) Then
            sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        End If
        
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
       
        objWorksheet.Cells(5, "G") = Mid(tmpPeriodFr, 6, 2) & "/" & Mid(tmpPeriodFr, 1, 4)
       
        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "1000" Then
                objWorksheet.Cells(6, "G") = rst2!Amount
            End If
       
            If Trim(rst2!ACCNT_CODE) = "1002" Then
                objWorksheet.Cells(7, "G") = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1010" Then
                objWorksheet.Cells(8, "G") = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1020" Then
                objWorksheet.Cells(9, "G") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1100" Then
                objWorksheet.Cells(10, "G") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1600" Then
                objWorksheet.Cells(11, "G") = rst2!Amount
            End If
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing

        sql = "select * from VIEW_GST_P2 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!JRNAL_TYPE) = "GJ" Then
                objWorksheet.Cells(19, "E") = rst2!Amount
            End If
            
            
            If Trim(rst2!JRNAL_TYPE) = "SI" Then
                objWorksheet.Cells(20, "E") = rst2!Amount
            End If
        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing


        sql = "select * from VIEW_GST_P2 where 0 = 0 "
        
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!JRNAL_TYPE) = "GJ" Then
                objWorksheet.Cells(19, "G") = rst2!Amount
            End If
            
            
            If Trim(rst2!JRNAL_TYPE) = "SI" Then
                objWorksheet.Cells(20, "G") = rst2!Amount
            End If
        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing

        sql = "select * from VIEW_GST_P3 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(24, "E") = rst2!Amount
            End If
            
            
            If Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(25, "E") = rst2!Amount
            End If
        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_P3 where 0 = 0 "
        
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(24, "G") = rst2!Amount
            End If
            
            
            If Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(25, "G") = rst2!Amount
            End If
        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_P4 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            If Trim(rst2!JRNAL_TYPE) = "CP" Then
                objWorksheet.Cells(29, "E") = rst2!Amount
            End If
            
            If Trim(rst2!JRNAL_TYPE) = "RC" Then
                objWorksheet.Cells(30, "E") = rst2!Amount
            End If
                        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        sql = "select * from VIEW_GST_P4 where 0 = 0 "
        
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            If Trim(rst2!JRNAL_TYPE) = "CP" Then
                objWorksheet.Cells(29, "G") = rst2!Amount
            End If
            
            If Trim(rst2!JRNAL_TYPE) = "RC" Then
                objWorksheet.Cells(30, "G") = rst2!Amount
            End If
                        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        sql = "select * from VIEW_GST_P5 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            If Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(34, "E") = rst2!Amount
            End If
            
            If Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(35, "E") = rst2!Amount
            End If
                        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        sql = "select * from VIEW_GST_P5 where 0 = 0 "
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            If Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(34, "G") = rst2!Amount
            End If
            
            If Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(35, "G") = rst2!Amount
            End If
                        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
    
        sql = "select * from VIEW_GST_P6 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!JRNAL_TYPE) = "CP" Then
                objWorksheet.Cells(39, "E") = rst2!Amount
            End If
                        
            If Trim(rst2!JRNAL_TYPE) = "RC" Then
                objWorksheet.Cells(40, "E") = rst2!Amount
            End If
            
            If Trim(rst2!JRNAL_TYPE) = "SI" Then
                objWorksheet.Cells(41, "E") = rst2!Amount
            End If
                       
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        sql = "select * from VIEW_GST_P6 where 0 = 0 "
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!JRNAL_TYPE) = "CP" Then
                objWorksheet.Cells(39, "G") = rst2!Amount
            End If
                        
            If Trim(rst2!JRNAL_TYPE) = "RC" Then
                objWorksheet.Cells(40, "G") = rst2!Amount
            End If
            
            If Trim(rst2!JRNAL_TYPE) = "SI" Then
                objWorksheet.Cells(41, "G") = rst2!Amount
            End If
                       
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_P7_OSR where 0 = 0 "
        If Len(Trim(strAppend)) Then
            sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            objWorksheet.Cells(45, "E") = rst2!Amount
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        If txtPeriodMthFr = "01" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr - 1) & "011"
        ElseIf txtPeriodMthFr = "02" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr - 1) & "012"
        ElseIf txtPeriodMthFr = "03" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "001"
        ElseIf txtPeriodMthFr = "04" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "002"
        ElseIf txtPeriodMthFr = "05" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "003"
        ElseIf txtPeriodMthFr = "06" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "004"
        ElseIf txtPeriodMthFr = "07" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "005"
        ElseIf txtPeriodMthFr = "08" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "006"
        ElseIf txtPeriodMthFr = "09" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "007"
        ElseIf txtPeriodMthFr = "10" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "008"
        ElseIf txtPeriodMthFr = "11" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "009"
        ElseIf txtPeriodMthFr = "12" Then
            tmpPeriodFr2 = Trim(txtPeriodYrFr) & "010"
        End If

    
        sql = "select * from VIEW_GST_P7_OSR where 0 = 0 "
        sql = sql + " and PERIOD = '" & tmpPeriodFr2 & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            objWorksheet.Cells(45, "G") = rst2!Amount
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        sql = "select * from VIEW_GST_P7 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(47, "E") = rst2!Amount
            End If
                        
            If Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(48, "E") = rst2!Amount
            End If
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        sql = "select * from VIEW_GST_P7 where 0 = 0 "
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(47, "G") = rst2!Amount
            End If
                        
            If Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(48, "G") = rst2!Amount
            End If
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    
        sql = "select * from VIEW_GST_P9 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!JRNAL_TYPE) = "CP" And Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(52, "E") = rst2!Amount
                TotalAmt9 = TotalAmt9 + rst2!Amount
            End If
                        
            If Trim(rst2!JRNAL_TYPE) = "RC" And Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(53, "E") = rst2!Amount
                TotalAmt9 = TotalAmt9 + rst2!Amount
            End If
            
            If Trim(rst2!JRNAL_TYPE) = "SI" And Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(54, "E") = rst2!Amount
                TotalAmt9 = TotalAmt9 + rst2!Amount
            End If
            rst2.MoveNext
        Loop
            
                objWorksheet.Cells(56, "E") = TotalAmt9
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_P9 where 0 = 0 "
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        TotalAmt9 = 0
        Do Until rst2.EOF
            If Trim(rst2!JRNAL_TYPE) = "CP" And Trim(rst2!ANAL_T7) = "" Then
                objWorksheet.Cells(52, "G") = rst2!Amount
                TotalAmt9 = TotalAmt9 + rst2!Amount
            End If
                        
            If Trim(rst2!JRNAL_TYPE) = "RC" And Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(53, "G") = rst2!Amount
                TotalAmt9 = TotalAmt9 + rst2!Amount
            End If
            
            If Trim(rst2!JRNAL_TYPE) = "SI" And Trim(rst2!ANAL_T7) = "OSR" Then
                objWorksheet.Cells(54, "G") = rst2!Amount
                TotalAmt9 = TotalAmt9 + rst2!Amount
            End If
            rst2.MoveNext
        Loop
            
                objWorksheet.Cells(56, "G") = TotalAmt9
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_P8 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "8830" Then
                objWorksheet.Cells(15, "E") = rst2!Amount
            End If
                        
            If Trim(rst2!ACCNT_CODE) = "8834" Then
                objWorksheet.Cells(16, "E") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "8839" Then
                objWorksheet.Cells(17, "E") = rst2!Amount
            End If
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_P8 where 0 = 0 "
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "8830" Then
                objWorksheet.Cells(15, "G") = rst2!Amount
            End If
                        
            If Trim(rst2!ACCNT_CODE) = "8834" Then
                objWorksheet.Cells(16, "G") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "8839" Then
                objWorksheet.Cells(17, "G") = rst2!Amount
            End If
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_ACC6810_Summary where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
                objWorksheet.Cells(18, "E") = rst2!Amount
                        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_ACC6810_Summary where 0 = 0 "
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            objWorksheet.Cells(18, "G") = rst2!Amount
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        
        
        
        Set objWorksheet = objWorkbook.Sheets(3)
        Set objWorksheet = objWorkbook.Worksheets.Item(3)
        Cnt = 0
        
        sql = "select * from VIEW_GST_ACC2kTO5999_Listing where 0 = 0 "
 
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
 
        Do Until rst2.EOF
            Cnt = CDbl(Cnt) + 1
            objWorksheet.Cells(1 + CDbl(Cnt), "A") = Trim(rst2!ACCNT_CODE)
 
            objWorksheet.Cells(1 + CDbl(Cnt), "B") = Trim(rst2!TREFERENCE)
 
            objWorksheet.Cells(1 + CDbl(Cnt), "C") = Trim(rst2!Amount)
  
            objWorksheet.Cells(1 + CDbl(Cnt), "D") = Trim(rst2!D_C)
            
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        
        'Set objWorksheet = objWorkbook.Sheets(5)
        'Set objWorksheet = objWorkbook.Worksheets.Item(5)
        Cnt = 0
        
        sql = "select * from VIEW_GST_ACC6810_Listing where 0 = 0 "
 
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
 
        Do Until rst2.EOF
            Cnt = CDbl(Cnt) + 1
            objWorksheet.Cells(1 + CDbl(Cnt), "F") = Trim(rst2!ACCNT_CODE)
 
            objWorksheet.Cells(1 + CDbl(Cnt), "G") = Trim(rst2!TREFERENCE)
 
            objWorksheet.Cells(1 + CDbl(Cnt), "H") = Trim(rst2!Amount)
  
            objWorksheet.Cells(1 + CDbl(Cnt), "I") = Trim(rst2!D_C)
            
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        
        
'Sheet 2
    Set objWorksheet = objWorkbook.Sheets(2)
    Set objWorksheet = objWorkbook.Worksheets.Item(2)

    sql = "select * from VIEW_GST_P1 where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic

        
        objWorksheet.Cells(2, "B") = Trim(txtPeriodMthFr) & "/" & Trim(txtPeriodYrFr)
        objWorksheet.Cells(5, "E") = Trim(txtPeriodMthFr) & "/" & Trim(txtPeriodYrFr)
        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "1000" Then
                objWorksheet.Cells(6, "E") = rst2!Amount
                TotalAmt1000 = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1002" Then
                objWorksheet.Cells(7, "E") = rst2!Amount
                TotalAmt1002 = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1010" Then
                objWorksheet.Cells(8, "E") = rst2!Amount
                TotalAmt1010 = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1020" Then
                objWorksheet.Cells(9, "E") = rst2!Amount
                TotalAmt1020 = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1100" Then
                objWorksheet.Cells(10, "E") = rst2!Amount
                TotalAmt1100 = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1600" Then
                objWorksheet.Cells(11, "E") = rst2!Amount
                TotalAmt1600 = rst2!Amount
            End If
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
       
       
        
        sql = "select * from VIEW_GST_P8 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "8830" Then
                objWorksheet.Cells(15, "E") = rst2!Amount
                TotalAmt8830 = rst2!Amount
            End If
                        
            If Trim(rst2!ACCNT_CODE) = "8834" Then
                objWorksheet.Cells(16, "E") = rst2!Amount
                TotalAmt8834 = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "8839" Then
                objWorksheet.Cells(17, "E") = rst2!Amount
                TotalAmt8839 = rst2!Amount
            End If
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_ACC6810_Summary where 0 = 0 "
        
            If Len(Trim(strAppend)) Then
                sql = sql + strAppend
            End If
            rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
            
            Do Until rst2.EOF
                objWorksheet.Cells(18, "E") = rst2!Amount
                rst2.MoveNext
            Loop
            rst2.Close
            Set rst2 = Nothing
       
        linecnt = 20
        TotalAmt = 0
        sql = "select * from VIEW_GST_P2 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            objWorksheet.Cells(linecnt, "A") = rst2!JRNAL_TYPE
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        'objWorksheet.Cells(7 + CDbl(PVCnt), "J").Borders.LineStyle = True
        'objWorksheet.Cells(7 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
         'ws.Cells(1, X).Font.Bold = True
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 1000 BY JT"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
        objWorksheet.Cells(linecnt, "E").Font.Bold = True
        
        If TotalAmt1000 = TotalAmt Then
            objWorksheet.Cells(linecnt, "F") = "Agree"
            objWorksheet.Cells(linecnt, "F").Font.Bold = True
        Else
            objWorksheet.Cells(linecnt, "F") = "Not Agree " & TotalAmt1000 - TotalAmt
        End If
        linecnt = CDbl(linecnt) + 2
        TotalAmt = 0
        
        sql = "select * from VIEW_GST_P3 where 0 = 0 "
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            objWorksheet.Cells(linecnt, "A") = IIf(Trim(rst2!ANAL_T7) = "", "EMPTY", rst2!ANAL_T7)
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 1000 BY T7"
        objWorksheet.Cells(linecnt, "E") = TotalAmt

        
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
        objWorksheet.Cells(linecnt, "E").Font.Bold = True
        
        If TotalAmt1000 = TotalAmt Then
            objWorksheet.Cells(linecnt, "F") = "Agree"
            objWorksheet.Cells(linecnt, "F").Font.Bold = True
        Else
            objWorksheet.Cells(linecnt, "F") = "Not Agree " & TotalAmt1000 - TotalAmt
        End If
        linecnt = CDbl(linecnt) + 2
        TotalAmt = 0
        
        sql = "select * from VIEW_GST_P4 where 0 = 0 "
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            objWorksheet.Cells(linecnt, "A") = rst2!JRNAL_TYPE
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        
        
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 1002 BY JT"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
        objWorksheet.Cells(linecnt, "E").Font.Bold = True
        
        If TotalAmt1002 = TotalAmt Then
            objWorksheet.Cells(linecnt, "F") = "Agree"
            objWorksheet.Cells(linecnt, "F").Font.Bold = True
        Else
            objWorksheet.Cells(linecnt, "F") = "Not Agree " & TotalAmt1002 - TotalAmt
        End If
        
        linecnt = CDbl(linecnt) + 2
        TotalAmt = 0
        
        sql = "select * from VIEW_GST_P5 where 0 = 0 "
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            objWorksheet.Cells(linecnt, "A") = IIf(Trim(rst2!ANAL_T7) = "", "EMPTY", rst2!ANAL_T7)
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 1002 BY T7"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        
        
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
        objWorksheet.Cells(linecnt, "E").Font.Bold = True
        
        If TotalAmt1002 = TotalAmt Then
            objWorksheet.Cells(linecnt, "F") = "Agree"
            objWorksheet.Cells(linecnt, "F").Font.Bold = True
        Else
            objWorksheet.Cells(linecnt, "F") = "Not Agree " & TotalAmt1002 - TotalAmt
        End If
        
        linecnt = CDbl(linecnt) + 2
        TotalAmt = 0
    
    
        sql = "select * from VIEW_GST_P6 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        Total8830CPAmt = 0
        Do Until rst2.EOF
        
            objWorksheet.Cells(linecnt, "A") = rst2!JRNAL_TYPE
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            If Trim(rst2!JRNAL_TYPE) = "CP" Then
                Total8830CPAmt = rst2!Amount
            End If
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 8830 BY JT"
        
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
        objWorksheet.Cells(linecnt, "E").Font.Bold = True
        
        If TotalAmt8830 = TotalAmt Then
            objWorksheet.Cells(linecnt, "F") = "Agree"
            objWorksheet.Cells(linecnt, "F").Font.Bold = True
        Else
            objWorksheet.Cells(linecnt, "F") = "Not Agree " & TotalAmt8830 - TotalAmt
        End If
        
        linecnt = CDbl(linecnt) + 2
        TotalAmt = 0
        
        
        sql = "select * from VIEW_GST_P7_OSR where 0 = 0 "
        If Len(Trim(strAppend)) Then
            sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            objWorksheet.Cells(linecnt, "A") = "AC 8830 BY T7 = OSR (Prior Month)"
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            If Total8830CPAmt = rst2!Amount Then
                objWorksheet.Cells(linecnt, "F") = "Agree"
                objWorksheet.Cells(linecnt, "F").Font.Bold = True
            Else
                objWorksheet.Cells(linecnt, "F").Font.Bold = True
                objWorksheet.Cells(linecnt, "F") = "Not Agree (" & Total8830CPAmt - rst2!Amount & ")"
            End If
            
            objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(linecnt, "E").Font.Bold = True

            linecnt = CDbl(linecnt) + 2
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        TotalAmt = 0
 
    
        sql = "select * from VIEW_GST_P7 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            objWorksheet.Cells(linecnt, "A") = IIf(Trim(rst2!ANAL_T7) = "", "EMPTY", rst2!ANAL_T7)
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 8830 BY T7"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
        objWorksheet.Cells(linecnt, "E").Font.Bold = True
        
        linecnt = CDbl(linecnt) + 2
        TotalAmt = 0
    
    
        sql = "select * from VIEW_GST_P9 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
        
            objWorksheet.Cells(linecnt, "A") = Trim(rst2!JRNAL_TYPE)
            objWorksheet.Cells(linecnt, "B") = IIf(Trim(rst2!ANAL_T7) = "", "EMPTY", rst2!ANAL_T7)
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 8830 BY T7"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
        objWorksheet.Cells(linecnt, "E").Borders(xlEdgeBottom).LineStyle = xlDouble
        objWorksheet.Cells(linecnt, "E").Font.Bold = True
        
        linecnt = CDbl(linecnt) + 2
        TotalAmt = 0

        

        cmdPrint.Enabled = True
        CmdExit.Enabled = True
        
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        
        MsgBox "Report is ready!"
    
    End If

    Exit Sub
End Sub


Private Sub SUNListingSum2(Optional strAppend As String)
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
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim Cnt As Double
Dim tmpPeriodFr As String
Dim TotalAmt9 As Double
Dim tmpPeriodFr2 As String
Dim linecnt As Double
'Dim TotalAmt As Double

    TotalAmt = 0
    linecnt = 0
    TotalAmt9 = 0
    Cnt = 0
    total = 0
    Current = 0
    lstSUNListing.ListItems.Clear
    
    
        'tmpPeriodFr = Trim(txtPeriodYrFr) & "0" & Trim(txtPeriodMthFr)
        'strAppend = strAppend + " and PERIOD = '" & tmpPeriodFr & "'"
        If txtPeriodMthFr = "01" Then
            tmpPeriodFr = Trim(txtPeriodYrFr - 1) & "012"
        ElseIf txtPeriodMthFr = "02" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "001"
        ElseIf txtPeriodMthFr = "03" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "002"
        ElseIf txtPeriodMthFr = "04" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "003"
        ElseIf txtPeriodMthFr = "05" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "004"
        ElseIf txtPeriodMthFr = "06" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "005"
        ElseIf txtPeriodMthFr = "07" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "006"
        ElseIf txtPeriodMthFr = "08" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "007"
        ElseIf txtPeriodMthFr = "09" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "008"
        ElseIf txtPeriodMthFr = "10" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "009"
        ElseIf txtPeriodMthFr = "11" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "010"
        ElseIf txtPeriodMthFr = "12" Then
            tmpPeriodFr = Trim(txtPeriodYrFr) & "011"
        End If

    sql = "select * from VIEW_GST_P1 where 0 = 0 "
    
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
    
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "PaymentVoucher\SUNGSTRptV2.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        objWorksheet.Cells(2, "B") = Trim(txtPeriodMthFr) & "/" & Trim(txtPeriodYrFr)
        objWorksheet.Cells(5, "E") = Trim(txtPeriodMthFr) & "/" & Trim(txtPeriodYrFr)
        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "1000" Then
                objWorksheet.Cells(6, "E") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1002" Then
                objWorksheet.Cells(7, "E") = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1010" Then
                objWorksheet.Cells(8, "E") = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1020" Then
                objWorksheet.Cells(9, "E") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1100" Then
                objWorksheet.Cells(10, "E") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1600" Then
                objWorksheet.Cells(11, "E") = rst2!Amount
            End If
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing

        sql = "select * from VIEW_GST_P1 where 0 = 0 "
                
        'If Len(Trim(strAppend)) Then
        sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        'End If
        
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        objWorksheet.Cells(5, "G") = Mid(tmpPeriodFr, 6, 2) & "/" & Mid(tmpPeriodFr, 1, 4)
        
        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "1000" Then
                objWorksheet.Cells(6, "G") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1002" Then
                objWorksheet.Cells(7, "G") = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1010" Then
                objWorksheet.Cells(8, "G") = rst2!Amount
            End If
        
            If Trim(rst2!ACCNT_CODE) = "1020" Then
                objWorksheet.Cells(9, "G") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1100" Then
                objWorksheet.Cells(10, "G") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "1600" Then
                objWorksheet.Cells(11, "G") = rst2!Amount
            End If
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        linecnt = 20
        TotalAmt = 0
        sql = "select * from VIEW_GST_P2 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
                
                objWorksheet.Cells(linecnt, "A") = rst2!JRNAL_TYPE
                objWorksheet.Cells(linecnt, "E") = rst2!Amount
                TotalAmt = TotalAmt + rst2!Amount
                linecnt = CDbl(linecnt) + 1
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 1000 BY JT"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        linecnt = CDbl(linecnt) + 1
        TotalAmt = 0
        
        sql = "select * from VIEW_GST_P3 where 0 = 0 "
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            objWorksheet.Cells(linecnt, "A") = IIf(Trim(rst2!ANAL_T7) = "", "EMPTY", rst2!ANAL_T7)
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
        
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 1000 BY T7"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        linecnt = CDbl(linecnt) + 1
        TotalAmt = 0
        
        sql = "select * from VIEW_GST_P4 where 0 = 0 "
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            objWorksheet.Cells(linecnt, "A") = rst2!JRNAL_TYPE
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 1002 BY JT"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        linecnt = CDbl(linecnt) + 1
        TotalAmt = 0
        
        sql = "select * from VIEW_GST_P5 where 0 = 0 "
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            
            objWorksheet.Cells(linecnt, "A") = IIf(Trim(rst2!ANAL_T7) = "", "EMPTY", rst2!ANAL_T7)
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 1002 BY T7"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        linecnt = CDbl(linecnt) + 1
        TotalAmt = 0
    
    
        sql = "select * from VIEW_GST_P6 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
        
            objWorksheet.Cells(linecnt, "A") = rst2!JRNAL_TYPE
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 8830 BY JT"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        linecnt = CDbl(linecnt) + 1
        TotalAmt = 0
        
        
        sql = "select * from VIEW_GST_P7_OSR where 0 = 0 "
        If Len(Trim(strAppend)) Then
            sql = sql + " and PERIOD = '" & tmpPeriodFr & "'"
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            objWorksheet.Cells(linecnt, "A") = "AC 8830 BY T7 = OSR (Prior Month)"
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            linecnt = CDbl(linecnt) + 1
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        TotalAmt = 0
 
    
        sql = "select * from VIEW_GST_P7 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            objWorksheet.Cells(linecnt, "A") = IIf(Trim(rst2!ANAL_T7) = "", "EMPTY", rst2!ANAL_T7)
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 8830 BY T7"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        linecnt = CDbl(linecnt) + 1
        TotalAmt = 0
    
    
        sql = "select * from VIEW_GST_P9 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
        
            objWorksheet.Cells(linecnt, "A") = Trim(rst2!JRNAL_TYPE)
            objWorksheet.Cells(linecnt, "A") = IIf(Trim(rst2!ANAL_T7) = "", "EMPTY", rst2!ANAL_T7)
            objWorksheet.Cells(linecnt, "E") = rst2!Amount
            TotalAmt = TotalAmt + rst2!Amount
            linecnt = CDbl(linecnt) + 1
            
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        objWorksheet.Cells(linecnt, "A") = "T1 = AC 8830 BY T7"
        objWorksheet.Cells(linecnt, "E") = TotalAmt
        linecnt = CDbl(linecnt) + 1
        TotalAmt = 0
        
        sql = "select * from VIEW_GST_P8 where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
        rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            If Trim(rst2!ACCNT_CODE) = "8830" Then
                objWorksheet.Cells(15, "E") = rst2!Amount
            End If
                        
            If Trim(rst2!ACCNT_CODE) = "8834" Then
                objWorksheet.Cells(16, "E") = rst2!Amount
            End If
            
            If Trim(rst2!ACCNT_CODE) = "8839" Then
                objWorksheet.Cells(17, "E") = rst2!Amount
            End If
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        sql = "select * from VIEW_GST_ACC6810_Summary where 0 = 0 "
        
            If Len(Trim(strAppend)) Then
                sql = sql + strAppend
            End If
            rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
            
            Do Until rst2.EOF
                objWorksheet.Cells(18, "E") = rst2!Amount
                rst2.MoveNext
            Loop
            rst2.Close
            Set rst2 = Nothing
            
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
        
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        
        MsgBox "Report is ready!"
    
    End If

    Exit Sub
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelBordxToPayor(lstSUNListing)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrintSum_Click()

    Screen.MousePointer = vbHourglass

        medixmasterSUNSys.Open "File Name=" & BOSPath & "Conn\MediConnSUNSystem.UDL"
            Call SearchRecordSum
        medixmasterSUNSys.Close
        Set medixmasterSUNSys = Nothing

    Screen.MousePointer = Default
End Sub

Private Sub CmdSearch_Click()
    Screen.MousePointer = vbHourglass
        medixmasterSUNSys.Open "File Name=" & BOSPath & "Conn\MediConnSUNSystem.UDL"
            Call SearchRecord
        medixmasterSUNSys.Close
        Set medixmasterSUNSys = Nothing
    Screen.MousePointer = Default
End Sub

Private Sub txtPeriodMthFr_Change()
    If Len(txtPeriodMthFr) = 2 Then
        txtPeriodYrFr.SetFocus
    End If
End Sub
