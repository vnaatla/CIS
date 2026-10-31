VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaintenanceOthers 
   Caption         =   "Others"
   ClientHeight    =   8148
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11808
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8148
   ScaleWidth      =   11808
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame10 
      Height          =   705
      Left            =   360
      TabIndex        =   41
      Top             =   6840
      Width           =   11055
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   4080
         TabIndex        =   10
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Default         =   -1  'True
         Height          =   315
         Left            =   7320
         TabIndex        =   11
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   9000
         TabIndex        =   13
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   8160
         TabIndex        =   12
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   8
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   720
         TabIndex        =   6
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   315
         Left            =   3240
         TabIndex        =   9
         Top             =   285
         Width           =   840
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Others Maintenance"
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
         Top             =   120
         Width           =   3285
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7125
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   11685
      _ExtentX        =   20616
      _ExtentY        =   12573
      _Version        =   393216
      Tabs            =   5
      TabHeight       =   520
      TabCaption(0)   =   "Treatment Reason"
      TabPicture(0)   =   "frmMaintenanceOthers.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label21"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstReasonList"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Procedures"
      TabPicture(1)   =   "frmMaintenanceOthers.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstProceduresList"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Drugs"
      TabPicture(2)   =   "frmMaintenanceOthers.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lstDrugList"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label7"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Diagnosis Not Payable"
      TabPicture(3)   =   "frmMaintenanceOthers.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(3)"
      Tab(3).Control(1)=   "lstDiagnosisList"
      Tab(3).Control(2)=   "Label17"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Repudiation"
      TabPicture(4)   =   "frmMaintenanceOthers.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "lstRepudiationList"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label23"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   4
         Left            =   -74760
         TabIndex        =   60
         Top             =   720
         Width           =   10860
         Begin VB.TextBox txtREPCode 
            Height          =   285
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   22
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtREPDesc 
            Height          =   1095
            Left            =   2040
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   720
            Width           =   8535
         End
         Begin VB.Label Label20 
            Caption         =   "Repudiation Code   *"
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
            Left            =   120
            TabIndex        =   62
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label19 
            Caption         =   "Repudiation Description  *"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   720
            Width           =   1890
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   3
         Left            =   -74760
         TabIndex        =   55
         Top             =   720
         Width           =   10860
         Begin VB.TextBox txtDiagCode 
            Height          =   285
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   21
            Top             =   720
            Width           =   3495
         End
         Begin VB.TextBox txtDiagPLNCode 
            Height          =   285
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   20
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label10 
            Caption         =   "Plan Code    *"
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
            TabIndex        =   57
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label9 
            Caption         =   "Diagnosis Code *"
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
            TabIndex        =   56
            Top             =   720
            Width           =   1770
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   2
         Left            =   -74760
         TabIndex        =   51
         Top             =   720
         Width           =   10860
         Begin VB.TextBox txtPROdesc 
            Height          =   1095
            Left            =   2040
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   720
            Width           =   8535
         End
         Begin VB.TextBox txtPRO 
            Height          =   285
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   14
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "Procedure Description *"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   720
            Width           =   1770
         End
         Begin VB.Label Label12 
            Caption         =   "Procedure Code    *"
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
            TabIndex        =   52
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1335
         Index           =   1
         Left            =   -74760
         TabIndex        =   44
         Top             =   650
         Width           =   11175
         Begin VB.TextBox txtDrugValue 
            Height          =   315
            Left            =   5760
            MaxLength       =   8
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   1245
         End
         Begin VB.TextBox txtDrug 
            Height          =   285
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   16
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtDrugDesc 
            Height          =   600
            Left            =   1680
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   600
            Width           =   5325
         End
         Begin VB.TextBox txtDrugUnit 
            Height          =   315
            Left            =   3360
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label Label8 
            Caption         =   "Drug Value   *"
            Height          =   255
            Left            =   4680
            TabIndex        =   54
            Top             =   240
            Width           =   1650
         End
         Begin VB.Label Label6 
            Caption         =   "Description     *"
            Height          =   255
            Left            =   360
            TabIndex        =   46
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Drug Code   *"
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
            Left            =   360
            TabIndex        =   33
            Top             =   285
            Width           =   1740
         End
         Begin VB.Label Label2 
            Caption         =   "Drug Unit    *"
            Height          =   255
            Left            =   2280
            TabIndex        =   45
            Top             =   240
            Width           =   1650
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2175
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   660
         Width           =   11175
         Begin VB.TextBox txtReason 
            Height          =   285
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   3
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtReaDesc 
            Height          =   600
            Left            =   2160
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   500
            Width           =   8685
         End
         Begin VB.TextBox txtReaRemark 
            Height          =   960
            Left            =   2160
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   1080
            Width           =   8685
         End
         Begin VB.Label Label16 
            Caption         =   "Remark                     *"
            Height          =   255
            Left            =   405
            TabIndex        =   32
            Top             =   1200
            Width           =   1650
         End
         Begin VB.Label Label15 
            Caption         =   "Treatment Code   *"
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
            Left            =   360
            TabIndex        =   31
            Top             =   285
            Width           =   1740
         End
         Begin VB.Label Label14 
            Caption         =   "Description                 *"
            Height          =   255
            Left            =   360
            TabIndex        =   30
            Top             =   645
            Width           =   1695
         End
      End
      Begin VB.Frame Frame 
         Height          =   1215
         Left            =   -74640
         TabIndex        =   24
         Top             =   960
         Width           =   6375
         Begin VB.TextBox txtOCPDesc1 
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
            Left            =   1995
            MaxLength       =   50
            TabIndex        =   26
            Top             =   675
            Width           =   3300
         End
         Begin VB.TextBox txtOCPcode1 
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
            Left            =   1995
            MaxLength       =   2
            TabIndex        =   25
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Occupation Description"
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
            Left            =   240
            TabIndex        =   28
            Top             =   675
            Width           =   1770
         End
         Begin VB.Label Label1 
            Caption         =   "Occupation Code"
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
            Left            =   255
            TabIndex        =   27
            Top             =   360
            Width           =   1320
         End
      End
      Begin MSComctlLib.ListView lstReasonList 
         Height          =   3060
         Left            =   240
         TabIndex        =   34
         Top             =   3120
         Width           =   11115
         _ExtentX        =   19600
         _ExtentY        =   5398
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
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Treatment Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Remark"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstOccupationList1 
         Height          =   3120
         Left            =   -74640
         TabIndex        =   35
         Top             =   2640
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   5503
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
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Occupation Code"
            Object.Width           =   2611
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descriptions"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstDrugList 
         Height          =   4100
         Left            =   -74760
         TabIndex        =   47
         Top             =   2200
         Width           =   11115
         _ExtentX        =   19600
         _ExtentY        =   7218
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
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Drug Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Drug Unit"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Drug Value"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstProceduresList 
         Height          =   3060
         Left            =   -74760
         TabIndex        =   49
         Top             =   3120
         Width           =   11115
         _ExtentX        =   19600
         _ExtentY        =   5398
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
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Procedures Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Procedure Description"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView lstDiagnosisList 
         Height          =   4020
         Left            =   -74760
         TabIndex        =   58
         Top             =   2280
         Width           =   11115
         _ExtentX        =   19600
         _ExtentY        =   7091
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
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Plan Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Diagnosis Code"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView lstRepudiationList 
         Height          =   3060
         Left            =   -74760
         TabIndex        =   63
         Top             =   3120
         Width           =   11115
         _ExtentX        =   19600
         _ExtentY        =   5398
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
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Repudiation Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Repudiation Description"
            Object.Width           =   14111
         EndProperty
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "REPUDIATION LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74760
         TabIndex        =   64
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "DIAGNOSIS LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74760
         TabIndex        =   59
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "PROCEDURE LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74760
         TabIndex        =   50
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "DRUG LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74760
         TabIndex        =   48
         Top             =   2000
         Width           =   1155
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "OCCUPATIOP LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74640
         TabIndex        =   37
         Top             =   2400
         Width           =   1635
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "TREATMENT LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   360
         TabIndex        =   36
         Top             =   2880
         Width           =   1515
      End
   End
   Begin VB.Label Label30 
      Caption         =   "1 - Based on Actual"
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
      Left            =   7440
      TabIndex        =   43
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Label Label29 
      Caption         =   "2 - Based on Limit"
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
      Left            =   9480
      TabIndex        =   42
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label txtTotalRecord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5520
      TabIndex        =   40
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label32 
      Caption         =   "TOTAL RECORDS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4080
      TabIndex        =   39
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "Bold - Searchable        * - Mandatory Fields"
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
      Left            =   240
      TabIndex        =   38
      Top             =   7800
      Width           =   4215
   End
End
Attribute VB_Name = "frmMaintenanceOthers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim EditFlag As Boolean
Public REASONCriteria As String
Public PROCriteria As String
Public DRUGCriteria As String
Public REPCriteria As String
Public DIAGCriteria As String
Private m_blnDirection(1 To 4) As Boolean

Private Sub Form_Activate()
'    SSTab1.Tab = 0
End Sub

Private Sub cmdAdd_Click()
    EditFlag = False
    Call ClearForm(Me)
    Call DisableCmdButton(Me)
    Call EntriesEnable(True)
    CmdSave.Enabled = True
    CmdAdd.Enabled = True
   
    Select Case SSTab1.Tab
        Case 0
            txtReason.SetFocus
        Case 1
            txtPRO.SetFocus
        Case 2
            txtDrug.SetFocus
        Case 3
            txtDiagPLNCode.SetFocus
        Case 4
            txtREPCode.SetFocus
    End Select
End Sub
Private Sub Form_Load()
Const METHOD_NAME = "Form_Load"
On Error GoTo debugs

Call ActionLogUpdate(12, "frmMaintenanceOthers")


debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    Call EntriesEnable(True)
    CmdSave.Enabled = True
    CmdDelete.Enabled = False
     
    If txtReason <> "" Then
       txtReason.Enabled = False
       txtReaDesc.Enabled = True
       txtReaDesc.SetFocus
       txtReaRemark.Enabled = True
       txtReaRemark.SetFocus
    End If

    If txtPRO <> "" Then
       txtPRO.Enabled = False
       txtPROdesc.Enabled = True
       txtPROdesc.SetFocus
    End If
    
    If txtDrug <> "" Then
       txtDrug.Enabled = False
       txtDrugUnit.Enabled = True
       txtDrugUnit.SetFocus
       txtDrugValue.Enabled = True
       txtDrugValue.SetFocus
       txtDrugDesc.Enabled = True
       txtDrugDesc.SetFocus
     End If

    If txtDiagPLNCode <> "" Then
       txtDiagPLNCode.Enabled = False
       txtDiagCode.Enabled = True
       txtDiagCode.SetFocus
    End If
        
    If txtREPCode <> "" Then
       txtREPCode.Enabled = False
       txtREPDesc.Enabled = True
       txtREPDesc.SetFocus
    End If
End Sub

Private Sub cmdprint_Click()
    Screen.MousePointer = vbHourglass
    
   ' MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
  '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        
        Select Case SSTab1.Tab
            Case 0
                Call PrintREASON
            Case 1
                Call PrintPRO
            Case 2
                Call PrintDrug
            Case 3
                Call PrintDiagnosis
            Case 4
                Call PrintRepudiation
         
         End Select
        CmdAdd.Enabled = True
        CmdEdit.Enabled = False
        
   ' MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' MediConnCISWS.Close
   ' Set MediConnCISWS = Nothing
   ' MediConnCISPayment.Close
   ' Set MediConnCISPayment = Nothing
    Screen.MousePointer = Default
End Sub

Private Sub CmdRefresh_Click()
'    Select Case SSTab1.Tab
'        Case 1
'            Call ClinicBNFcombolisting
'        Case 2
'            Call ClnBNFCombolisting
'    End Select
End Sub

Private Sub CmdSave_Click()
  Screen.MousePointer = vbHourglass
        
   ' MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        Select Case SSTab1.Tab
            Case 0
                Call SaveReason
            Case 1
                Call SavePRO
            Case 2
                Call SaveDrug
            Case 3
                Call SaveDiagnosis
            Case 4
                Call SaveRepudiation
        End Select
        CmdAdd.Enabled = True
        CmdEdit.Enabled = False
 '   MediConnCISDB.Close
  '  Set MediConnCISDB = Nothing
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  MediConnCISWS.Close
  '  Set MediConnCISWS = Nothing
  '  MediConnCISPayment.Close
  '  Set MediConnCISPayment = Nothing
    Screen.MousePointer = Default
End Sub

Private Sub cmdDelete_Click()
Screen.MousePointer = vbHourglass
   ' MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
        Select Case SSTab1.Tab
           Case 0
               Call DeleteREASON
           Case 1
               Call DeletePRO
            Case 2
                Call DeleteDrug
            Case 3
                Call DeleteDiagnosis
            Case 4
                Call DeleteRepudiation
        End Select
        CmdAdd.Enabled = True
        CmdEdit.Enabled = False
       
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
   ' medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
    'MediConnCISPayment.Close
    'Set MediConnCISPayment = Nothing
    Screen.MousePointer = Default
End Sub

Private Sub cmdSearch_Click()
Screen.MousePointer = vbHourglass
    'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"

        Select Case SSTab1.Tab
        Case 0
            Call SearchREASON
            Call ClearForm(Me)
            txtReason.Enabled = True
            txtReaDesc.Enabled = True
            txtReaRemark.Enabled = True
            txtReason.SetFocus
        Case 1
            Call SearchPRO
            Call ClearForm(Me)
            txtPRO.Enabled = True
            txtPROdesc.Enabled = True
            txtPRO.SetFocus
        Case 2
            Call SearchDrug
            Call ClearForm(Me)
            txtDrug.Enabled = True
            txtDrugUnit.Enabled = True
            txtDrugValue.Enabled = True
            txtDrugDesc.Enabled = True
            txtDrug.SetFocus
        Case 3
            Call SearchDiagnosis
            Call ClearForm(Me)
            txtDiagPLNCode.Enabled = True
            txtDiagCode.Enabled = True
            txtDiagPLNCode.SetFocus
        Case 4
            Call SearchRepudiation
            Call ClearForm(Me)
            txtREPCode.Enabled = True
            txtREPDesc.Enabled = True
            txtREPCode.SetFocus
        End Select
        
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

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    CmdAdd.Enabled = True
    txtTotalRecord = 0
    lstReasonList.ListItems.Clear
    lstProceduresList.ListItems.Clear
    lstDrugList.ListItems.Clear
    lstDiagnosisList.ListItems.Clear
    lstRepudiationList.ListItems.Clear
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    Call EntriesEnable(True)
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


'************************Start EntriesEnable*********************************

Private Sub EntriesEnable(status As Boolean)
     
    txtReason.Enabled = status
    txtReaDesc.Enabled = status
    txtReaRemark.Enabled = status
    txtPRO.Enabled = status
    txtPROdesc.Enabled = status
    txtDrug.Enabled = status
    txtDrugUnit.Enabled = status
    txtDrugValue.Enabled = status
    txtDrugDesc.Enabled = status
    txtDiagPLNCode.Enabled = status
    txtDiagCode.Enabled = status
    txtREPCode.Enabled = status
    txtREPDesc.Enabled = status
    CmdSearch.Enabled = status
    
End Sub
'************************End EntriesEnable*********************************
    
'####### @@@ START SAVE, EDIT, ADD, DELETE REASON @@@ #######'

Private Sub DeleteREASON()
Const METHOD_NAME = "DeleteREASON"
On Error GoTo debugs
Dim rsDeleteREA As New ADODB.Recordset
Dim rsCheckREA As New ADODB.Recordset
Dim strsql, sqlREA As String
Dim DeleteRecord
    
EditFlag = False
       
If txtReason = "" Then
   MsgBox "Please Enter Reason Code", vbOKOnly + vbCritical, DialogBoxTitle
ElseIf lstReasonList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlREA = "Select Reason Code From TreatmentReason where ReasonCode = '" & Trim(lstReasonList.SelectedItem.Text) & "'"
    
'   rsCheckCOP.MaxRecords = 1
'   rsCheckCOP.Open sqlCOP, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
'    If rsCheckCOP.RecordCount Then
'        MsgBox "This record cannot be deleted because it is used in Calculation Table!", vbCritical, "Error in deleting record..."
'        Call ClearForm(Me)
'        CmdDelete.Enabled = False
'        CmdEdit.Enabled = False
    
'    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
             strsql = "Delete From TreatmentReason Where ReasonCode = '" & ChkStr(lstReasonList.SelectedItem.Text) & "'"
             rsDeleteREA.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call ClearForm(Me)
             Call REASONListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             txtReason.SetFocus
        Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

    
Private Sub SearchREASON()
Dim strAppend As String
    
  Screen.MousePointer = vbHourglass
  strAppend = ""
  REASONCriteria = ""
  
  If Len(Trim(txtReason)) Then
    strAppend = strAppend + " And ReasonCode = '" & Trim(Mid(txtReason, 1, 3)) & "'"
  'HLTCode like '%" & Mid(ChkStr(txtHLTcode), 1, 2) & "%'"
  End If
    
  REASONCriteria = strAppend
  Call REASONListing(strAppend)
  Screen.MousePointer = Default
  
End Sub

Private Sub PrintREASON()
    RptReasonList.Show
End Sub

Private Sub SaveReason()
Const METHOD_NAME = "SaveReason"
On Error GoTo debugs
   'On Error Resume Next
   Dim RecordSaved
   Dim REA As New ADODB.Recordset
   Dim SaveREA As New ADODB.Recordset
   Dim strsql As String
   
    If txtReason = "" Then
        MsgBox "Please Enter Reason Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtReason.SetFocus
    ElseIf Len(txtReason) < 3 Then
        MsgBox "Reason Code must be 3 integer!", vbOKOnly + vbCritical, DialogBoxTitle
        txtReason.SetFocus
    ElseIf txtReaDesc = "" Then
        MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtReaDesc.SetFocus
'    ElseIf txtCoRemark = "" Then
'        MsgBox "Please Enter Remarks", vbOKOnly + vbCritical, DialogBoxTitle
'        txtCoRemark.SetFocus
    Else
        strsql = "Select * From TreatmentReason where ReasonCode =  '" & Trim(txtReason) & "'"
        REA.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If REA.RecordCount > 0 And EditFlag = False Then
           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           REA.Close
           txtReason.SetFocus
           txtReason.SelStart = 0
           txtReason.SelLength = Len(txtReason)
           Set REA = Nothing
           Exit Sub
        Else
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           If RecordSaved = vbYes Then
              If EditFlag = False Then
                 SaveREA.Open "TreatmentReason", MediConnCISDB, adOpenKeyset, adLockOptimistic
                 SaveREA.AddNew
              Else
                 strsql = "Select * From TreatmentReason Where ReasonCode = '" & ChkStr(txtReason) & "'"
                 SaveREA.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
              End If
              
              With SaveREA
                !ReasonCode = ChkStr(txtReason)
                !ReasonDesc = ChkStr(txtReaDesc)
                !ReasonRemark = ChkStr(txtReaRemark)
                !ReasonLastUpdateUser = Logon
                !ReasonLastUpdateDate = Date
                
              End With
              SaveREA.update
            
             Call ClearForm(Me)
             'Call SearchCOP
             Call REASONListing(REASONCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             REA.Close
             Set SaveREA = Nothing
        End If
        EditFlag = False
     End If
   End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub


Private Sub REASONListing(Optional strAppend As String)
Const METHOD_NAME = "REASONListing"
On Error GoTo debugs

 'On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim REA As ListItem
    Dim sql As String
    Dim TotalREACount As Integer
   
    TotalREACount = 0
    lstReasonList.ListItems.Clear
    
    sql = " select ReasonCode, ReasonDesc, ReasonRemark " & _
      " From TreatmentReason where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            With rst2
            Set REA = lstReasonList.ListItems.Add(, , rst2!ReasonCode)
                REA.SubItems(1) = rst2!ReasonDesc
                REA.SubItems(2) = rst2!ReasonRemark
                rst2.MoveNext
            End With
        TotalREACount = TotalREACount + 1
        txtTotalRecord = TotalREACount
        Loop
    End If
    'REASONCriteria = ""
    rst2.Close
    Set rst2 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Frame9_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lstReasonList_Click()
Dim strsql As String
    
    If lstReasonList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtReason = lstReasonList.SelectedItem.Text
        txtReason.Enabled = False
        txtReaDesc = lstReasonList.SelectedItem.SubItems(1)
        txtReaDesc.Enabled = False
        txtReaRemark = lstReasonList.SelectedItem.SubItems(2)
        txtReaRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstReasonList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstReasonList.SortKey = ColumnHeader.Index - 1
    lstReasonList.Sorted = True
    If lstReasonList.SortOrder = lvwAscending Then
       lstReasonList.SortOrder = lvwDescending
    Else
       lstReasonList.SortOrder = lvwAscending
    End If
    
    txtReason.Enabled = False
End Sub

Private Sub lstReasonList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstReasonList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtReason = lstReasonList.SelectedItem.Text
        txtReason.Enabled = False
        txtReaDesc = lstReasonList.SelectedItem.SubItems(1)
        txtReaDesc.Enabled = False
        txtReaRemark = lstReasonList.SelectedItem.SubItems(2)
        txtReaRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstReasonList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstReasonList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtReason = lstReasonList.SelectedItem.Text
        txtReason.Enabled = False
        txtReaDesc = lstReasonList.SelectedItem.SubItems(1)
        txtReaDesc.Enabled = False
        txtReaRemark = lstReasonList.SelectedItem.SubItems(2)
        txtReaRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub
'####### @@@ END SAVE, EDIT , ADD, DELETE REASON @@@ #######'


Private Sub SSTab1_Click(PreviousTab As Integer)
    Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
    Case 0
        Call ClearForm(Me)
        lstReasonList.ListItems.Clear
'        CmdRefresh.Enabled = False
        CmdEdit.Enabled = False
        CmdSave.Enabled = True
        CmdDelete.Enabled = False
        txtTotalRecord = 0
        txtReason.SetFocus
    Case 1
        Call ClearForm(Me)
        lstProceduresList.ListItems.Clear
'        CmdRefresh.Enabled = False
        CmdEdit.Enabled = False
        CmdSave.Enabled = True
        CmdDelete.Enabled = False
        txtTotalRecord = 0
        txtPRO.SetFocus
    Case 2
        Call ClearForm(Me)
        lstDrugList.ListItems.Clear
'        CmdRefresh.Enabled = False
        CmdEdit.Enabled = False
        CmdSave.Enabled = True
        CmdDelete.Enabled = False
        txtTotalRecord = 0
'        txtDrug.SetFocus
    Case 3
        Call ClearForm(Me)
        lstDiagnosisList.ListItems.Clear
        CmdEdit.Enabled = False
        CmdSave.Enabled = True
        CmdDelete.Enabled = False
        txtTotalRecord = 0
        txtDiagPLNCode.SetFocus
    End Select
    Screen.MousePointer = Default


End Sub

'####### @@@ START SAVE, EDIT , ADD, DELETE PROCEDURES @@@ #######'
Private Sub DeletePRO()
Const METHOD_NAME = "DeletePRO"
On Error GoTo debugs
Dim rsDeletePRO As New ADODB.Recordset
Dim rsCheckPRO As New ADODB.Recordset
Dim strsql, sqlPRO As String
Dim DeleteRecord
    
EditFlag = False
       
If txtPRO = "" Then
   MsgBox "Please Enter Procedure Code", vbOKOnly + vbCritical, DialogBoxTitle
ElseIf lstProceduresList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlPRO = "Select ProcedureCode From Procedures where ProcedureCode = '" & Trim(lstProceduresList.SelectedItem.Text) & "'"
    
   DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
   If DeleteRecord = vbYes Then
      strsql = "Delete From Procedures Where ProcedureCode = '" & ChkStr(lstProceduresList.SelectedItem.Text) & "'"
        rsDeletePRO.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
            CmdSave.Enabled = True
            Call ClearForm(Me)
            Call PROListing
            Call DisableCmdButton(Me)
            Call EntriesEnable(True)
            txtPRO.SetFocus
   Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
   End If
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

    
Private Sub SearchPRO()
Dim strAppend As String
    
  Screen.MousePointer = vbHourglass
  strAppend = ""
  PROCriteria = ""
  
  If Len(Trim(txtPRO)) Then
    strAppend = strAppend + " And ProcedureCode = '" & Trim(Mid(txtPRO, 1, 3)) & "'"
  'HLTCode like '%" & Mid(ChkStr(txtHLTcode), 1, 2) & "%'"
  End If
    
  PROCriteria = strAppend
  Call PROListing(strAppend)
  Screen.MousePointer = Default
  
End Sub

Private Sub PrintPRO()
    rptProcedureslist.Show
End Sub

Private Sub SavePRO()
Const METHOD_NAME = "SavePRO"
On Error GoTo debugs
   'On Error Resume Next
   Dim RecordSaved
   Dim PRO As New ADODB.Recordset
   Dim SavePRO As New ADODB.Recordset
   Dim strsql As String
   
    If txtPRO = "" Then
        MsgBox "Please Enter Procedure Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtPRO.SetFocus
    'ElseIf Len(txtPRO) < 3 Then
    '    MsgBox "Procedure Code must be 3 integer!", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtPRO.SetFocus
    ElseIf txtPROdesc = "" Then
        MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtPROdesc.SetFocus
'    ElseIf txtCoRemark = "" Then
'        MsgBox "Please Enter Remarks", vbOKOnly + vbCritical, DialogBoxTitle
'        txtCoRemark.SetFocus
    Else
        strsql = "Select * From Procedures where ProcedureCode =  '" & Trim(txtPRO) & "'"
        PRO.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If PRO.RecordCount > 0 And EditFlag = False Then
           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           PRO.Close
           txtPRO.SetFocus
           txtPRO.SelStart = 0
           txtPRO.SelLength = Len(txtPRO)
           Set PRO = Nothing
           Exit Sub
        Else
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           If RecordSaved = vbYes Then
              If EditFlag = False Then
                 SavePRO.Open "Procedures", MediConnCISDB, adOpenKeyset, adLockOptimistic
                 SavePRO.AddNew
              Else
                 strsql = "Select * From Procedures Where ProcedureCode = '" & ChkStr(txtPRO) & "'"
                 SavePRO.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
              End If
              
              With SavePRO
                !ProcedureCode = ChkStr(txtPRO)
                !ProcedureDesc = ChkStr(txtPROdesc)
                !ProLastUpdateUser = Logon
                !ProLastUpdateDate = Date
                
              End With
              SavePRO.update
            
             Call ClearForm(Me)
             'Call SearchCOP
             Call PROListing(PROCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             PRO.Close
             Set SavePRO = Nothing
        End If
        EditFlag = False
     End If
   End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub PROListing(Optional strAppend As String)
Const METHOD_NAME = "PROListing"
On Error GoTo debugs

 'On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim PRO As ListItem
    Dim sql As String
    Dim TotalPROCount As Integer
   
    TotalPROCount = 0
    lstProceduresList.ListItems.Clear
    
    sql = " Select ProcedureCode, ProcedureDesc " & _
      " From Procedures where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            With rst2
            Set PRO = lstProceduresList.ListItems.Add(, , rst2!ProcedureCode)
                PRO.SubItems(1) = rst2!ProcedureDesc
                'PRO.SubItems(2) = rst2!ReasonRemark
                rst2.MoveNext
            End With
        TotalPROCount = TotalPROCount + 1
        txtTotalRecord = TotalPROCount
        Loop
    End If
    'PROCriteria = ""
    rst2.Close
    Set rst2 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstProceduresList_Click()
Dim strsql As String
    
    If lstProceduresList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtPRO = lstProceduresList.SelectedItem.Text
        txtPRO.Enabled = False
        txtPROdesc = lstProceduresList.SelectedItem.SubItems(1)
        txtPROdesc.Enabled = False
        'txtReaRemark = lstProceduresList.SelectedItem.SubItems(2)
        'txtReaRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstProceduresList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstProceduresList.SortKey = ColumnHeader.Index - 1
    lstProceduresList.Sorted = True
    If lstProceduresList.SortOrder = lvwAscending Then
       lstProceduresList.SortOrder = lvwDescending
    Else
       lstProceduresList.SortOrder = lvwAscending
    End If
    
    txtPRO.Enabled = False
End Sub

Private Sub lstProceduresList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstProceduresList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtPRO = lstProceduresList.SelectedItem.Text
        txtPRO.Enabled = False
        txtPROdesc = lstProceduresList.SelectedItem.SubItems(1)
        txtPROdesc.Enabled = False
        'txtReaRemark = lstProceduresList.SelectedItem.SubItems(2)
        'txtReaRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstProceduresList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstProceduresList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtPRO = lstProceduresList.SelectedItem.Text
        txtPRO.Enabled = False
        txtPROdesc = lstProceduresList.SelectedItem.SubItems(1)
        txtPROdesc.Enabled = False
        'txtReaRemark = lstProceduresList.SelectedItem.SubItems(2)
        'txtReaRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabl0ed = False
    End If
End Sub
'####### @@@ END SAVE, EDIT , ADD, DELETE PROCEDURES @@@ #######'

'####### @@@ START SAVE, EDIT , ADD, DELETE DRUG @@@ #######'
Private Sub DeleteDrug()
Const METHOD_NAME = "DeleteDrug"
On Error GoTo debugs
Dim rsDeleteDG As New ADODB.Recordset
Dim rsCheckDG As New ADODB.Recordset
Dim strsql, sqlDG As String
Dim DeleteRecord
    
EditFlag = False
       
If txtDrug = "" Then
   MsgBox "Please Enter Drug Code", vbOKOnly + vbCritical, DialogBoxTitle
ElseIf lstDrugList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlDG = "Select DrugCode From Drug where DrugCode = '" & Trim(lstDrugList.SelectedItem.Text) & "'"
    
   DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
   If DeleteRecord = vbYes Then
      strsql = "Delete From Drug Where DrugCode = '" & ChkStr(lstDrugList.SelectedItem.Text) & "'"
        rsDeleteDG.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
            CmdSave.Enabled = True
            Call ClearForm(Me)
            Call DrugListing
            Call DisableCmdButton(Me)
            Call EntriesEnable(True)
            txtDrug.SetFocus
   Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
   End If
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

    
Private Sub SearchDrug()
Dim strAppend As String
    
  Screen.MousePointer = vbHourglass
  strAppend = ""
  DRUGCriteria = ""
  
  If Len(Trim(txtDrug)) Then
    strAppend = strAppend + " And drugCode = '" & Trim(Mid(txtDrug, 1, 3)) & "'"
  'HLTCode like '%" & Mid(ChkStr(txtHLTcode), 1, 2) & "%'"
  End If
    
  DRUGCriteria = strAppend
  Call DrugListing(strAppend)
  Screen.MousePointer = Default
  
End Sub

Private Sub PrintDrug()
    rptDruglist.Show
End Sub

Private Sub SaveDrug()
Const METHOD_NAME = "SaveDrug"
On Error GoTo debugs
   'On Error Resume Next
   Dim RecordSaved
   Dim DG As New ADODB.Recordset
   Dim SaveDG As New ADODB.Recordset
   Dim strsql As String
   
    If txtDrug = "" Then
        MsgBox "Please Enter Drug Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtDrug.SetFocus
    'ElseIf Len(txtDrug) < 3 Then
    '    MsgBox "Drug Code must be 3 integer!", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtDrug.SetFocus
    ElseIf txtDrugUnit = "" Then
        MsgBox "Please Enter Drug Unit", vbOKOnly + vbCritical, DialogBoxTitle
        txtDrugUnit.SetFocus
    ElseIf txtDrugValue = "" Then
        MsgBox "Please Enter Drug Value", vbOKOnly + vbCritical, DialogBoxTitle
        txtDrugValue.SetFocus
    'ElseIf txtDrugDesc = "" Then
    '    MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtDrugDesc.SetFocus
    Else
        strsql = "Select * From Drug where DrugCode =  '" & Trim(txtDrug) & "'"
        DG.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If DG.RecordCount > 0 And EditFlag = False Then
           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           DG.Close
           txtDrug.SetFocus
           txtDrug.SelStart = 0
           txtDrug.SelLength = Len(txtDrug)
           Set DG = Nothing
           Exit Sub
        Else
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           If RecordSaved = vbYes Then
              If EditFlag = False Then
                 SaveDG.Open "Drug", MediConnCISDB, adOpenKeyset, adLockOptimistic
                 SaveDG.AddNew
              Else
                 strsql = "Select * From Drug Where DrugCode = '" & ChkStr(txtDrug) & "'"
                 SaveDG.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
              End If
              
              With SaveDG
                !DrugCode = ChkStr(txtDrug)
                !DrugUnit = txtDrugUnit
                !DrugValue = Format(txtDrugValue, "#,##0.00")
                !DrugDesc = ChkStr(txtDrugDesc)
                !DrugLastUpdateUser = Logon
                !DrugLastUpdateDate = Date
                
              End With
              SaveDG.update
            
             Call ClearForm(Me)
             'Call SearchCOP
             Call DrugListing(DRUGCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             DG.Close
             Set SaveDG = Nothing
        End If
        EditFlag = False
     End If
   End If
   
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub


Private Sub DrugListing(Optional strAppend As String)
Const METHOD_NAME = "DrugListing"
On Error GoTo debugs

 'On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim DG As ListItem
    Dim sql As String
    Dim TotalDGCount As Integer
   
    TotalDGCount = 0
    lstDrugList.ListItems.Clear
    
    sql = " Select DrugCode, DrugUnit, DrugValue, DrugDesc" & _
      " From Drug where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            With rst2
            Set DG = lstDrugList.ListItems.Add(, , rst2!DrugCode)
                DG.SubItems(1) = rst2!DrugUnit
                DG.SubItems(2) = rst2!DrugValue
                DG.SubItems(3) = rst2!DrugDesc
                rst2.MoveNext
            End With
        TotalDGCount = TotalDGCount + 1
        txtTotalRecord = TotalDGCount
        Loop
    End If
    'DRUGCriteria = ""
    rst2.Close
    Set rst2 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstDrugList_Click()
Dim strsql As String
    
    If lstDrugList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtDrug = lstDrugList.SelectedItem.Text
        txtDrug.Enabled = False
        txtDrugUnit = lstDrugList.SelectedItem.SubItems(1)
        txtDrugUnit.Enabled = False
        txtDrugValue = lstDrugList.SelectedItem.SubItems(2)
        txtDrugValue.Enabled = False
        txtDrugDesc = lstDrugList.SelectedItem.SubItems(3)
        txtDrugDesc.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstDrugList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstDrugList.SortKey = ColumnHeader.Index - 1
    lstDrugList.Sorted = True
    If lstDrugList.SortOrder = lvwAscending Then
       lstDrugList.SortOrder = lvwDescending
    Else
       lstDrugList.SortOrder = lvwAscending
    End If
    
    txtDrug.Enabled = False
End Sub

Private Sub lstDrugList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstDrugList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtDrug = lstDrugList.SelectedItem.Text
        txtDrug.Enabled = False
        txtDrugUnit = lstDrugList.SelectedItem.SubItems(1)
        txtDrugUnit.Enabled = False
        txtDrugValue = lstDrugList.SelectedItem.SubItems(2)
        txtDrugValue.Enabled = False
        txtDrugDesc = lstDrugList.SelectedItem.SubItems(3)
        txtDrugDesc.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstDrugList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstDrugList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtDrug = lstDrugList.SelectedItem.Text
        txtDrug.Enabled = False
        txtDrugUnit = lstDrugList.SelectedItem.SubItems(1)
        txtDrugUnit.Enabled = False
        txtDrugValue = lstDrugList.SelectedItem.SubItems(2)
        txtDrugValue.Enabled = False
        txtDrugDesc = lstDrugList.SelectedItem.SubItems(3)
        txtDrugDesc.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub
'####### @@@ END SAVE, EDIT , ADD, DELETE PROCEDURES @@@ #######'

'####### @@@ START SAVE, EDIT , ADD, DELETE DIAGNOSIS @@@ #######'
Private Sub DeleteDiagnosis()
Const METHOD_NAME = "DeleteDiagnosis"
On Error GoTo debugs

Dim rsDeleteDiag As New ADODB.Recordset
Dim rsCheckDiag As New ADODB.Recordset
Dim strsql, sqlDiag As String
Dim DeleteRecord
    
EditFlag = False
       
If txtDiagPLNCode = "" Then
   MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
ElseIf lstDiagnosisList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlDiag = "Select PLNCode From DiagNotPayable where PLNCode = '" & Trim(lstDiagnosisList.SelectedItem.Text) & "'"
    
   DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
   If DeleteRecord = vbYes Then
      strsql = "Delete From DiagNotPayable Where PLNCode = '" & ChkStr(lstDiagnosisList.SelectedItem.Text) & "'"
        rsDeleteDiag.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
            CmdSave.Enabled = True
            Call ClearForm(Me)
            Call DiagnosisListing
            Call DisableCmdButton(Me)
            Call EntriesEnable(True)
            txtDiagPLNCode.SetFocus
   Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
   End If
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

    
Private Sub SearchDiagnosis()
Const METHOD_NAME = "SearchDiagnosis"
On Error GoTo debugs
Dim strAppend As String
    
  Screen.MousePointer = vbHourglass
  strAppend = ""
  DIAGCriteria = ""
  
  If Len(Trim(txtDiagPLNCode)) Then
    strAppend = strAppend + " and PLNCode = '" & Trim(txtDiagPLNCode) & "'"
    'strAppend = strAppend + " And ProcedureCode = '" & Trim(Mid(txtPRO, 1, 3)) & "'"
    'HLTCode like '%" & Mid(ChkStr(txtHLTcode), 1, 2) & "%'"
  End If
    
  If Len(Trim(txtDiagCode)) Then
    strAppend = strAppend + " and DiagCode = '" & Trim(txtDiagCode) & "'"
  End If
    
  DIAGCriteria = strAppend
  Call DiagnosisListing(strAppend)
  Screen.MousePointer = Default
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub PrintDiagnosis()
    rptDiagnosisList.Show
End Sub

Private Sub SaveDiagnosis()
Const METHOD_NAME = "SaveDiagnosis"
On Error GoTo debugs
   'On Error Resume Next
   Dim RecordSaved
   Dim Diag As New ADODB.Recordset
   Dim SaveDiag As New ADODB.Recordset
   Dim strsql As String
   
    If txtDiagPLNCode = "" Then
        MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtDiagPLNCode.SetFocus
    'ElseIf Len(txtPRO) < 3 Then
    '    MsgBox "Procedure Code must be 3 integer!", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtPRO.SetFocus
    ElseIf txtDiagCode = "" Then
        MsgBox "Please Enter Diagnosis Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtDiagCode.SetFocus
    Else
        strsql = "Select * From DiagNotPayable where PLNCode =  '" & Trim(txtDiagPLNCode) & "'"
        Diag.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If Diag.RecordCount > 0 And EditFlag = False Then
           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           Diag.Close
           txtDiagPLNCode.SetFocus
           txtDiagPLNCode.SelStart = 0
           txtDiagPLNCode.SelLength = Len(txtDiagPLNCode)
           Set Diag = Nothing
           Exit Sub
        Else
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           If RecordSaved = vbYes Then
              If EditFlag = False Then
                 SaveDiag.Open "DiagNotPayable", MediConnCISDB, adOpenKeyset, adLockOptimistic
                 SaveDiag.AddNew
              Else
                 strsql = "Select * From DiagNotPayable Where PLNCode = '" & ChkStr(txtDiagPLNCode) & "'"
                 SaveDiag.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
              End If
              
              With SaveDiag
                !PLNCode = ChkStr(txtDiagPLNCode)
                !DIagCode = ChkStr(txtDiagCode)
                !LastUpdateUser = Logon
                !LastUpdateDate = Date
                
              End With
              SaveDiag.update
            
             Call ClearForm(Me)
             'Call SearchCOP
             Call DiagnosisListing(DIAGCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             Diag.Close
             Set SaveDiag = Nothing
        End If
        EditFlag = False
     End If
   End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub


Private Sub DiagnosisListing(Optional strAppend As String)
Const METHOD_NAME = "DiagnosisListing"
On Error GoTo debugs
 'On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim Diag As ListItem
    Dim sql As String
    Dim TotalDiagCount As Integer
   
    TotalDiagCount = 0
    lstDiagnosisList.ListItems.Clear
    
    sql = " Select PLNCode, DiagCode " & _
      " From DiagNotPayable where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            With rst2
            Set Diag = lstDiagnosisList.ListItems.Add(, , rst2!PLNCode)
                Diag.SubItems(1) = rst2!DIagCode
                'PRO.SubItems(2) = rst2!ReasonRemark
                rst2.MoveNext
            End With
        TotalDiagCount = TotalDiagCount + 1
        txtTotalRecord = TotalDiagCount
        Loop
    End If
    'PROCriteria = ""
    rst2.Close
    Set rst2 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstDiagnosisList_Click()
Dim strsql As String
    
    If lstDiagnosisList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtDiagPLNCode = lstDiagnosisList.SelectedItem.Text
        txtDiagPLNCode.Enabled = False
        txtDiagCode = lstDiagnosisList.SelectedItem.SubItems(1)
        txtDiagCode.Enabled = False
        'txtReaRemark = lstProceduresList.SelectedItem.SubItems(2)
        'txtReaRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstDiagnosisList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstDiagnosisList.SortKey = ColumnHeader.Index - 1
    lstDiagnosisList.Sorted = True
    If lstDiagnosisList.SortOrder = lvwAscending Then
       lstDiagnosisList.SortOrder = lvwDescending
    Else
       lstDiagnosisList.SortOrder = lvwAscending
    End If
    
    txtDiagPLNCode.Enabled = False
End Sub

Private Sub lstDiagnosisList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstDiagnosisList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtDiagPLNCode = lstDiagnosisList.SelectedItem.Text
        txtDiagPLNCode.Enabled = False
        txtDiagCode = lstDiagnosisList.SelectedItem.SubItems(1)
        txtDiagCode.Enabled = False
        'txtReaRemark = lstProceduresList.SelectedItem.SubItems(2)
        'txtReaRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstDiagnosisList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstDiagnosisList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtDiagPLNCode = lstDiagnosisList.SelectedItem.Text
        txtDiagPLNCode.Enabled = False
        txtDiagCode = lstDiagnosisList.SelectedItem.SubItems(1)
        txtDiagCode.Enabled = False
        'txtReaRemark = lstProceduresList.SelectedItem.SubItems(2)
        'txtReaRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabl0ed = False
    End If
End Sub
'######## @@@ END SAVE, EDIT , ADD, DELETE DIAGNOSIS @@@ ########'

'####### @@@ START SAVE, EDIT , ADD, DELETE REPUDIATION @@@ #######'
Private Sub DeleteRepudiation()
Const METHOD_NAME = "DeleteRepudiation"
On Error GoTo debugs
Dim rsDeleteREP As New ADODB.Recordset
Dim rsCheckREP As New ADODB.Recordset
Dim strsql, sqlREP As String
Dim DeleteRecord
    
EditFlag = False
       
If txtREPCode = "" Then
   MsgBox "Please Enter Repudiation Code", vbOKOnly + vbCritical, DialogBoxTitle
ElseIf lstRepudiationList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlREP = "Select RepCode From Repudiation where RepCode = '" & Trim(lstRepudiationList.SelectedItem.Text) & "'"
    
   DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
   If DeleteRecord = vbYes Then
      strsql = "Delete From Repudiation Where RepCode = '" & ChkStr(lstRepudiationList.SelectedItem.Text) & "'"
        rsDeleteREP.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
            CmdSave.Enabled = True
            Call ClearForm(Me)
            Call RepudiationListing
            Call DisableCmdButton(Me)
            Call EntriesEnable(True)
            txtREPCode.SetFocus
   Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
   End If
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

    
Private Sub SearchRepudiation()
Dim strAppend As String
    
  Screen.MousePointer = vbHourglass
  strAppend = ""
  REPCriteria = ""
  
  If Len(Trim(txtREPCode)) Then
    strAppend = strAppend + " and RepCode = '" & Trim(txtREPCode) & "'"
    'strAppend = strAppend + " And ProcedureCode = '" & Trim(Mid(txtPRO, 1, 3)) & "'"
    'HLTCode like '%" & Mid(ChkStr(txtHLTcode), 1, 2) & "%'"
  End If
    
  REPCriteria = strAppend
  Call RepudiationListing(strAppend)
  Screen.MousePointer = Default
  
End Sub

Private Sub PrintRepudiation()
    rptRepudiationList.Show
End Sub

Private Sub SaveRepudiation()
Const METHOD_NAME = "SaveRepudiation"
On Error GoTo debugs
   'On Error Resume Next
   Dim RecordSaved
   Dim REP As New ADODB.Recordset
   Dim SaveREP As New ADODB.Recordset
   Dim strsql As String
   
    If txtREPCode = "" Then
        MsgBox "Please Enter Repudiation Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtREPCode.SetFocus
    'ElseIf Len(txtPRO) < 3 Then
    '    MsgBox "Procedure Code must be 3 integer!", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtPRO.SetFocus
    ElseIf txtREPDesc = "" Then
        MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtREPDesc.SetFocus
    Else
        strsql = "Select * From Repudiation where RepCode =  '" & Trim(txtREPCode) & "'"
        REP.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If REP.RecordCount > 0 And EditFlag = False Then
           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           REP.Close
           txtREPCode.SetFocus
           txtREPCode.SelStart = 0
           txtREPCode.SelLength = Len(txtREPCode)
           Set REP = Nothing
           Exit Sub
        Else
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           If RecordSaved = vbYes Then
              If EditFlag = False Then
                 SaveREP.Open "Repudiation", MediConnCISDB, adOpenKeyset, adLockOptimistic
                 SaveREP.AddNew
              Else
                 strsql = "Select * From Repudiation Where RepCode = '" & ChkStr(txtREPCode) & "'"
                 SaveREP.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
              End If
              
              With SaveREP
                !RepCode = ChkStr(txtREPCode)
                !RepDesc = ChkStr(txtREPDesc)
                !RepLastUpdateUser = Logon
                !RepLastUpdateDate = Date
                
              End With
              SaveREP.update
            
             Call ClearForm(Me)
             'Call SearchCOP
             Call RepudiationListing(REPCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             REP.Close
             Set SaveREP = Nothing
        End If
        EditFlag = False
     End If
   End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub RepudiationListing(Optional strAppend As String)
Const METHOD_NAME = "RepudiationListing"
On Error GoTo debugs
 'On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim REP As ListItem
    Dim sql As String
    Dim TotalREPCount As Integer
   
    TotalREPCount = 0
    lstRepudiationList.ListItems.Clear
    
    sql = " Select RepCode, RepDesc " & _
      " From Repudiation where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            With rst2
            Set REP = lstRepudiationList.ListItems.Add(, , rst2!RepCode)
                REP.SubItems(1) = rst2!RepDesc
                'PRO.SubItems(2) = rst2!ReasonRemark
                rst2.MoveNext
            End With
        TotalREPCount = TotalREPCount + 1
        txtTotalRecord = TotalREPCount
        Loop
    End If
    'PROCriteria = ""
    rst2.Close
    Set rst2 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstRepudiationList_Click()
Dim strsql As String
    
    If lstRepudiationList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtREPCode = lstRepudiationList.SelectedItem.Text
        txtREPCode.Enabled = False
        txtREPDesc = lstRepudiationList.SelectedItem.SubItems(1)
        txtREPDesc.Enabled = False
        'txtReaRemark = lstProceduresList.SelectedItem.SubItems(2)
        'txtReaRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstRepudiationList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstRepudiationList.SortKey = ColumnHeader.Index - 1
    lstRepudiationList.Sorted = True
    If lstRepudiationList.SortOrder = lvwAscending Then
       lstRepudiationList.SortOrder = lvwDescending
    Else
       lstRepudiationList.SortOrder = lvwAscending
    End If
    
    txtREPCode.Enabled = False
End Sub

Private Sub lstRepudiationList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstRepudiationList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtREPCode = lstRepudiationList.SelectedItem.Text
        txtREPCode.Enabled = False
        txtREPDesc = lstRepudiationList.SelectedItem.SubItems(1)
        txtREPDesc.Enabled = False
        'txtReaRemark = lstProceduresList.SelectedItem.SubItems(2)
        'txtReaRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstRepudiationList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstRepudiationList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtREPCode = lstRepudiationList.SelectedItem.Text
        txtREPCode.Enabled = False
        txtREPDesc = lstRepudiationList.SelectedItem.SubItems(1)
        txtREPDesc.Enabled = False
        'txtReaRemark = lstProceduresList.SelectedItem.SubItems(2)
        'txtReaRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabl0ed = False
    End If
End Sub
'####### @@@ END SAVE, EDIT , ADD, DELETE REPUDIATION @@@ #######'

Private Sub txtReason_Keypress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtDrugValue_Keypress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
