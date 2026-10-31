VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaintenanceBenefits 
   Caption         =   "Benefits"
   ClientHeight    =   8592
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8592
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame10 
      Height          =   705
      Left            =   240
      TabIndex        =   35
      Top             =   6840
      Width           =   11655
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   8280
         TabIndex        =   25
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   4200
         TabIndex        =   23
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   3360
         TabIndex        =   22
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   840
         TabIndex        =   19
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   20
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   21
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   9120
         TabIndex        =   26
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   9960
         TabIndex        =   27
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   5040
         TabIndex        =   24
         Top             =   285
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   28
      Top             =   360
      Width           =   11535
      _ExtentX        =   20341
      _ExtentY        =   11240
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Category"
      TabPicture(0)   =   "frmMaintenanceBenefits.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label38"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstCATList"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Benefit Entry"
      TabPicture(1)   =   "frmMaintenanceBenefits.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FGrid"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Benefit Listing"
      TabPicture(2)   =   "frmMaintenanceBenefits.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstBNFList"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame1(4)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
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
         Height          =   1980
         Index           =   4
         Left            =   -74760
         TabIndex        =   48
         Top             =   480
         Width           =   11115
         Begin VB.TextBox TxtBNFPercentage 
            Height          =   315
            Left            =   7200
            MaxLength       =   3
            TabIndex        =   16
            Top             =   1440
            Width           =   1185
         End
         Begin VB.TextBox TxtRMPerAnn 
            Height          =   315
            Left            =   9720
            MaxLength       =   8
            TabIndex        =   15
            Top             =   1080
            Width           =   1185
         End
         Begin VB.ComboBox cboCatCode 
            Height          =   315
            Left            =   7200
            TabIndex        =   13
            Top             =   720
            Width           =   3705
         End
         Begin VB.ComboBox cboBINSCode 
            Height          =   315
            Left            =   7200
            TabIndex        =   12
            Top             =   360
            Width           =   3705
         End
         Begin VB.ComboBox cboBPLNCode 
            Height          =   315
            Left            =   1800
            TabIndex        =   8
            Top             =   720
            Width           =   3615
         End
         Begin VB.ComboBox cboBHLTCode 
            Height          =   315
            Left            =   1800
            TabIndex        =   7
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox txtBNFCalculation 
            Height          =   315
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   10
            Top             =   1440
            Width           =   1275
         End
         Begin VB.TextBox TxtDay 
            Height          =   315
            Left            =   4320
            MaxLength       =   3
            TabIndex        =   11
            Text            =   "  "
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox txtRMPerVisit 
            Height          =   315
            Left            =   7200
            MaxLength       =   8
            TabIndex        =   14
            Top             =   1080
            Width           =   1185
         End
         Begin MSMask.MaskEdBox txtEffdate 
            Height          =   315
            Left            =   1800
            TabIndex        =   9
            Top             =   1080
            Width           =   3615
            _ExtentX        =   6371
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            Caption         =   "BNF Percentage"
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
            Left            =   5760
            TabIndex        =   60
            Top             =   1500
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "No of Ann."
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
            Left            =   8520
            TabIndex        =   59
            Top             =   1100
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Eff Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   255
            TabIndex        =   57
            Top             =   1100
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Health Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label Label21 
            Caption         =   "Plan Claimable"
            Height          =   330
            Left            =   -1845
            TabIndex        =   55
            Top             =   675
            Width           =   1050
         End
         Begin VB.Label Label22 
            Caption         =   "BNF Calculation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   54
            Top             =   1500
            Width           =   1560
         End
         Begin VB.Label Label19 
            Caption         =   "Plan Code    "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   255
            TabIndex        =   53
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Insured Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5760
            TabIndex        =   52
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label11 
            Caption         =   "No of Days        "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3120
            TabIndex        =   51
            Top             =   1500
            Width           =   1035
         End
         Begin VB.Label Label28 
            Caption         =   "BNF Code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5760
            TabIndex        =   50
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "RM Per Visit"
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
            Left            =   5760
            TabIndex        =   49
            Top             =   1095
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   -74760
         TabIndex        =   45
         Top             =   2160
         Width           =   11115
         Begin VB.ComboBox CboGrpCom 
            Height          =   315
            ItemData        =   "frmMaintenanceBenefits.frx":0054
            Left            =   7680
            List            =   "frmMaintenanceBenefits.frx":0056
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   360
            Width           =   3255
         End
         Begin VB.ComboBox CboPayorCode 
            Height          =   315
            ItemData        =   "frmMaintenanceBenefits.frx":0058
            Left            =   1920
            List            =   "frmMaintenanceBenefits.frx":005A
            TabIndex        =   17
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label15 
            Caption         =   "Group Company  "
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
            Left            =   6000
            TabIndex        =   47
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label16 
            Caption         =   "Payor Code    "
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
            TabIndex        =   46
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   -74760
         TabIndex        =   40
         Top             =   420
         Width           =   11055
         Begin VB.ComboBox cboTHLTCode 
            Height          =   315
            Left            =   1800
            TabIndex        =   3
            Top             =   240
            Width           =   3615
         End
         Begin VB.ComboBox cboTPlnCode 
            Height          =   315
            Left            =   7200
            TabIndex        =   5
            Top             =   240
            Width           =   3105
         End
         Begin VB.ComboBox cboTInsCode 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   4
            Top             =   600
            Width           =   3615
         End
         Begin MSMask.MaskEdBox txtTEffdate 
            Height          =   315
            Left            =   7200
            TabIndex        =   6
            Top             =   600
            Width           =   3105
            _ExtentX        =   5482
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label12 
            Caption         =   "Insured Type *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   44
            Top             =   600
            Width           =   1485
         End
         Begin VB.Label Label10 
            Caption         =   "Plan Code     *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5760
            TabIndex        =   43
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Health Type  *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   135
            TabIndex        =   42
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label9 
            Caption         =   "Eff Date       *"
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
            Left            =   5760
            TabIndex        =   41
            Top             =   600
            Width           =   1335
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
         Height          =   1620
         Index           =   5
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   11025
         Begin VB.TextBox TxtCATDescription 
            Height          =   855
            Left            =   1680
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   600
            Width           =   9255
         End
         Begin VB.TextBox TxtCATCode 
            Height          =   285
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   1
            Text            =   "  "
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label37 
            Caption         =   "Description       *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   1635
         End
         Begin VB.Label Label36 
            Caption         =   "Category Code  *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   135
            TabIndex        =   31
            Top             =   240
            Width           =   1560
         End
         Begin VB.Label Label34 
            Caption         =   "Plan Claimable"
            Height          =   330
            Left            =   -1845
            TabIndex        =   30
            Top             =   675
            Width           =   1050
         End
      End
      Begin MSComctlLib.ListView lstCATList 
         Height          =   3885
         Left            =   240
         TabIndex        =   33
         Top             =   2280
         Width           =   11055
         _ExtentX        =   19495
         _ExtentY        =   6858
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
            Text            =   "Category Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Category Description"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid FGrid 
         Height          =   4695
         Left            =   -74760
         TabIndex        =   39
         Top             =   1620
         Width           =   11055
         _ExtentX        =   19495
         _ExtentY        =   8276
         _Version        =   393216
         Rows            =   40
         Cols            =   7
         FixedCols       =   2
         AllowUserResizing=   3
      End
      Begin MSComctlLib.ListView lstBNFList 
         Height          =   3045
         Left            =   -74760
         TabIndex        =   58
         Top             =   3120
         Width           =   11055
         _ExtentX        =   19495
         _ExtentY        =   5376
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "BNF Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "HLT Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "INS Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "PLN Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "BNF Calculation"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "No of Days"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "No of Annum"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "RM Per Visit"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "BNF Percentage"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Eff Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "Payor Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Group Company"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Caption         =   "BENEFIT CATEGORY LIST"
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
         Left            =   240
         TabIndex        =   34
         Top             =   2040
         Width           =   2295
      End
   End
   Begin VB.Label Label17 
      Caption         =   "4 - BNF Limit"
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
      Left            =   9720
      TabIndex        =   64
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "3 - Ascharge"
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
      Left            =   9720
      TabIndex        =   63
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "2 - Higher of Limit or %"
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
      Left            =   7320
      TabIndex        =   62
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "1 - Lower of Limit or %"
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
      Left            =   7320
      TabIndex        =   61
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label Label13 
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
      Height          =   375
      Left            =   4200
      TabIndex        =   36
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label35 
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
      Left            =   120
      TabIndex        =   38
      Top             =   7680
      Width           =   4695
   End
   Begin VB.Label txtTotalRecord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5640
      TabIndex        =   37
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "  Benefits Maintenance"
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
      TabIndex        =   0
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmMaintenanceBenefits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditFlag As Boolean
Public CATCriteria As String
Dim recordInBNF As Integer
Dim FCol, Frow As Integer
Dim strAppend As String
Private m_blnDirection(1 To 13) As Boolean

Private Sub cmdAdd_Click()
    EditFlag = False
    Call ClearForm(Me)
    Call DisableCmdButton(Me)
    Call EntriesEnable(True)
    CmdSave.Enabled = True
    CmdAdd.Enabled = True
   
    Select Case SSTab1.Tab
        Case 0
            TxtCATCode.SetFocus
            TxtCATDescription.Enabled = True
    End Select
End Sub
Private Sub Form_Load()
Const METHOD_NAME = "Form_Load"
On Error GoTo debugs

Call ActionLogUpdate(12, "frmMaintenanceBenefits")


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
     
    Select Case SSTab1.Tab
        Case 0
            If TxtCATCode <> "" Then
                TxtCATCode.Enabled = False
                TxtCATDescription.Enabled = True
                TxtCATDescription.SetFocus
            End If
    End Select

End Sub

Private Sub cmdprint_Click()
    Screen.MousePointer = vbHourglass
    
    
    'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"

    
        Select Case SSTab1.Tab
            Case 0
                Call PrintCAT
            Case 1
            
            Case 2
                If lstBNFList.ListItems.Count <> 0 Then
                    Call ExportListViewtoExcel(lstBNFList)
                Else
                    MsgBox "Report cannot be printed without any record.", vbInformation
                End If
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

   ' MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"


        Select Case SSTab1.Tab
            Case 1
                Call BNFcombolisting
            Case 2
                Call BNFCombolisting2
        End Select

   ' MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' MediConnCISWS.Close
   ' Set MediConnCISWS = Nothing
   ' MediConnCISPayment.Close
   ' Set MediConnCISPayment = Nothing
End Sub

Private Sub cmdSave_Click()
  Screen.MousePointer = vbHourglass
        
      '  MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
      '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
      '  MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        
        Select Case SSTab1.Tab
            Case 0
                Call SaveCAT
            Case 1
                Call SaveBNF
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

Private Sub cmdDelete_Click()
Screen.MousePointer = vbHourglass
    If SSTab1.Tab = 0 Then
       ' MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
       ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
       ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"

            Call DeleteCATEGORY
            CmdAdd.Enabled = True
            
       ' MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
       ' MediConnCISWS.Close
       ' Set MediConnCISWS = Nothing
       ' MediConnCISPayment.Close
       ' Set MediConnCISPayment = Nothing
    End If
Screen.MousePointer = Default
End Sub

Private Sub cmdSearch_Click()
Screen.MousePointer = vbHourglass

   ' MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"

        Select Case SSTab1.Tab
            Case 0
                Call SearchCAT
                Call ClearForm(Me)
                TxtCATCode.Enabled = True
                TxtCATDescription.Enabled = True
                TxtCATCode.SetFocus
            Case 1
                Call searchtBNF
            Case 2
                Call searchBNF
            End Select
            
    '    MediConnCISDB.Close
    '    Set MediConnCISDB = Nothing
    '    medixmasterconn.Close
    '    Set medixmasterconn = Nothing
    '    MediConnCISWS.Close
    '    Set MediConnCISWS = Nothing
    '    MediConnCISPayment.Close
    '    Set MediConnCISPayment = Nothing
        
Screen.MousePointer = Default
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    Call ClearGrid
    
    Select Case SSTab1.Tab
        Case 0
            CmdAdd.Enabled = True
        Case 1
            CmdAdd.Enabled = True
    End Select
        
    txtTotalRecord = 0
    Pervisit = "0"
    Percentage = "0"
    lstCATList.ListItems.Clear
    lstBNFList.ListItems.Clear
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    Call EntriesEnable(True)
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub


'************************Start EntriesEnable*********************************

Private Sub EntriesEnable(status As Boolean)
     
    TxtCATCode.Enabled = status
    CmdSearch.Enabled = status
    
End Sub
'************************End EntriesEnable*********************************
     
    ' '''''''''''''''''' Save, edit , Add, delete Category..... ''''''''''''''
Private Sub DeleteCATEGORY()
Const METHOD_NAME = "DeleteCATEGORY"
On Error GoTo debugs
Dim rsDeleteCAT As New ADODB.Recordset
Dim rsCheckCAT As New ADODB.Recordset
Dim strsql, sqlCAT As String
Dim DeleteRecord
    
EditFlag = False
       
If TxtCATCode = "" Then
   MsgBox "Please Enter Category Code", vbOKOnly + vbCritical, DialogBoxTitle
ElseIf lstCATList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlCAT = "Select BNFCATCode From BNFCategoryNew where BNFCATCode = '" & Trim(lstCATList.SelectedItem.Text) & "'"
    

        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
             strsql = "Delete From BNFCategoryNew Where BNFCATCode = '" & ChkStr(lstCATList.SelectedItem.Text) & "'"
             rsDeleteCAT.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call ClearForm(Me)
             Call CATListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             TxtCATCode.SetFocus
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
    
Private Sub SearchCAT()
Dim strAppend As String
    
  Screen.MousePointer = vbHourglass
  strAppend = ""
  
  If Len(Trim(TxtCATCode)) Then
    strAppend = strAppend + " And BNFCATCode = '" & Trim(TxtCATCode) & "'"
  End If
    
  CATCriteria = strAppend
  Call CATListing(strAppend)
  Screen.MousePointer = Default
  
End Sub

Private Sub PrintCAT()
    RptBNFCat.Show
End Sub

Private Sub SaveCAT()
Const METHOD_NAME = "SaveCAT"
On Error GoTo debugs
   On Error Resume Next
   Dim RecordSaved
   Dim CAT As New ADODB.Recordset
   Dim SaveCAT As New ADODB.Recordset
   Dim strsql As String
   
    If TxtCATCode = "" Then
        MsgBox "Please Enter Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtCATCode.SetFocus
    ElseIf Len(TxtCATCode) < 3 Then
        MsgBox "Category Code must be 3 integer!", vbOKOnly + vbCritical, DialogBoxTitle
        TxtCATCode.SetFocus
    ElseIf TxtCATDescription = "" Then
        MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtCATDescription.SetFocus
    Else
        strsql = "Select * From BNFCategoryNew where BNFCATCode =  '" & Trim(TxtCATCode) & "'"
        CAT.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If CAT.RecordCount > 0 And EditFlag = False Then
           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           CAT.Close
           TxtCATCode.SetFocus
           TxtCATCode.SelStart = 0
           TxtCATCode.SelLength = Len(TxtCATCode)
           Set CAT = Nothing
           Exit Sub
        Else
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           If RecordSaved = vbYes Then
              If EditFlag = False Then
                 SaveCAT.Open "BNFCategoryNew", MediConnCISDB, adOpenKeyset, adLockOptimistic
                 SaveCAT.AddNew
              Else
                 strsql = "Select * From BNFCategoryNew Where BNFCATCode = '" & ChkStr(TxtCATCode) & "'"
                 SaveCAT.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
              End If
              
              With SaveCAT
                !BNFCatCode = ChkStr(TxtCATCode)
                !BNFCATDescription = ChkStr(TxtCATDescription)
                !BNFCATLastUpdateUser = Logon
                !BNFCATLastUpdateDate = Date
                 
              End With
              SaveCAT.update
            
             Call ClearForm(Me)
             Call CATListing(CATCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             CAT.Close
             Set SaveCAT = Nothing
        End If
        EditFlag = False
     End If
   End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub CATListing(Optional strAppend As String)
Const METHOD_NAME = "CATListing"
On Error GoTo debugs
 'On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim CAT As ListItem
    Dim sql As String
    Dim TotalCATCount As Integer
   
    TotalCATCount = 0
    lstCATList.ListItems.Clear
    
    sql = " select BNFCATCode, BNFCATDescription " & _
      " From BNFCategoryNew where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    rst2.Open sql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do While Not rst2.EOF
            With rst2
            Set CAT = lstCATList.ListItems.Add(, , rst2!BNFCatCode)
                CAT.SubItems(1) = rst2!BNFCATDescription
                rst2.MoveNext
            End With
        TotalCATCount = TotalCATCount + 1
        txtTotalRecord = TotalCATCount
        Loop
    End If
    CATCriteria = ""
    rst2.Close
    Set rst2 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub

Private Sub lstBNFlist_click()
    
    If lstBNFList.ListItems.Count > 0 Then
        Call ClearForm(Me)
        'Call BNFCombolisting2
        cboCatCode = lstBNFList.SelectedItem.Text
        cboBHLTCode = lstBNFList.SelectedItem.SubItems(2)
        cboBINSCode = lstBNFList.SelectedItem.SubItems(3)
        cboBPLNCode = lstBNFList.SelectedItem.SubItems(4)
        txtBNFCalculation = lstBNFList.SelectedItem.SubItems(5)
        TxtDay = lstBNFList.SelectedItem.SubItems(6)
        TxtRMPerAnn = lstBNFList.SelectedItem.SubItems(7)
        txtRMPerVisit = lstBNFList.SelectedItem.SubItems(8)
        TxtBNFPercentage = lstBNFList.SelectedItem.SubItems(9)
        If lstBNFList.SelectedItem.SubItems(10) <> "__/__/____" Then
            txtEffdate = lstBNFList.SelectedItem.SubItems(10)
        End If
        CboPayorCode = lstBNFList.SelectedItem.SubItems(11)
        CboGrpCom = lstBNFList.SelectedItem.SubItems(12)
        
    End If
End Sub

Private Sub lstBNFList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstBNFList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstBNFList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstBNFList, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstBNFList, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstBNFList, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(13)
    End Select
End Sub


Private Sub lstBNFList_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstBNFList.ListItems.Count > 0 Then
        Call ClearForm(Me)
        'Call BNFCombolisting2
        cboCatCode = lstBNFList.SelectedItem.Text
        cboBHLTCode = lstBNFList.SelectedItem.SubItems(2)
        cboBINSCode = lstBNFList.SelectedItem.SubItems(3)
        cboBPLNCode = lstBNFList.SelectedItem.SubItems(4)
        txtBNFCalculation = lstBNFList.SelectedItem.SubItems(5)
        TxtDay = lstBNFList.SelectedItem.SubItems(6)
        TxtRMPerAnn = lstBNFList.SelectedItem.SubItems(7)
        txtRMPerVisit = lstBNFList.SelectedItem.SubItems(8)
        TxtBNFPercentage = lstBNFList.SelectedItem.SubItems(9)
        If lstBNFList.SelectedItem.SubItems(10) <> "__/__/____" Then
            txtEffdate = lstBNFList.SelectedItem.SubItems(10)
        End If
        CboPayorCode = lstBNFList.SelectedItem.SubItems(11)
        CboGrpCom = lstBNFList.SelectedItem.SubItems(12)
        
    End If
End Sub

Private Sub lstBNFList_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstBNFList.ListItems.Count > 0 Then
        Call ClearForm(Me)
        'Call BNFCombolisting2
        cboCatCode = lstBNFList.SelectedItem.Text
        cboBHLTCode = lstBNFList.SelectedItem.SubItems(2)
        cboBINSCode = lstBNFList.SelectedItem.SubItems(3)
        cboBPLNCode = lstBNFList.SelectedItem.SubItems(4)
        txtBNFCalculation = lstBNFList.SelectedItem.SubItems(5)
        TxtDay = lstBNFList.SelectedItem.SubItems(6)
        TxtRMPerAnn = lstBNFList.SelectedItem.SubItems(7)
        txtRMPerVisit = lstBNFList.SelectedItem.SubItems(8)
        TxtBNFPercentage = lstBNFList.SelectedItem.SubItems(9)
        If lstBNFList.SelectedItem.SubItems(10) <> "__/__/____" Then
            txtEffdate = lstBNFList.SelectedItem.SubItems(10)
        End If
        CboPayorCode = lstBNFList.SelectedItem.SubItems(11)
        CboGrpCom = lstBNFList.SelectedItem.SubItems(12)
        
    End If
End Sub

Private Sub lstCATList_Click()
Dim strsql As String
    
    If lstCATList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        TxtCATCode = lstCATList.SelectedItem.Text
        TxtCATCode.Enabled = False
        TxtCATDescription = lstCATList.SelectedItem.SubItems(1)
        TxtCATDescription.Enabled = False
            
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstCATList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstCATList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstCATList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    End Select
End Sub


Private Sub FGrid_Click()
Dim temp, tempCol, tempRow, tempTitle As String
Dim Calculation, value1 As Double
Dim value2, value3, value4 As Double
    
    tempCol = FGrid.Col
    tempRow = FGrid.Row
    FGrid.Row = 0
    tempTitle = FGrid.Text
    FGrid.Row = tempRow
    FGrid.Col = 1
    temp = InputBox(tempTitle, FGrid.Text)
    FGrid.Col = tempCol
    If temp <> "" Then
        FGrid = temp
        If tempCol = 2 Then
            FGrid.Col = 2
            If IsNumeric(FGrid.Text) = True Then
                If FGrid.Text >= 1 And FGrid.Text <= 4 Then
                    Calculation = FGrid.Text
                Else
                    FGrid.Text = ""
                End If
            Else
                FGrid.Text = ""
            End If
        End If
        
        If tempCol = 3 Then
            FGrid.Col = 3
            If IsNumeric(FGrid.Text) = True Then
                value1 = FGrid.Text
            Else
                FGrid.Text = ""
            End If
        End If
        
        If tempCol = 4 Then
            FGrid.Col = 4
            If IsNumeric(FGrid.Text) = True Then
                value2 = FGrid.Text
            Else
                FGrid.Text = ""
            End If
        End If
            
        If tempCol = 5 Then
            FGrid.Col = 5
            value3 = FGrid.Text
        End If
        
        If tempCol = 6 Then
            FGrid.Col = 6
            If IsNumeric(FGrid.Text) = True Then
                value4 = FGrid.Text
            Else
                FGrid.Text = ""
            End If
        End If
    End If
End Sub

Private Sub SaveBNF()
Const METHOD_NAME = "SaveBNF"
On Error GoTo debugs
On Error GoTo UpdateErr
Dim BNF As New ADODB.Recordset
Dim SaveBNF As New ADODB.Recordset
Dim RecordSaved
Dim strsql As String
Dim currentRecord As Integer
Dim beginTrans As Boolean
Dim Pervisit As String
Dim Percentage As String
Dim Calculation As String

Calculation = "0"
Pervisit = "0"
Percentage = "0"

If Len(Trim(cboTHLTCode)) = 0 Then
 MsgBox "Please Enter Health Type !", vbOKOnly + vbCritical, DialogBoxTitle
    cboTHLTCode.SetFocus
ElseIf Len(Trim(cboTInsCode)) = 0 Then
    MsgBox "Please Enter Insured Type!", vbOKOnly + vbCritical, DialogBoxTitle
    cboTInsCode.SetFocus
ElseIf Len(Trim(cboTPlnCode)) = 0 Then
    MsgBox "Please Enter Plan Code!", vbOKOnly + vbCritical, DialogBoxTitle
    cboTPlnCode.SetFocus
ElseIf txtTEffdate = "__/__/____" Or IsDate(txtTEffdate) = False Then
    MsgBox "Please Enter (Valid )effectice Date!", vbOKOnly + vbCritical, DialogBoxTitle
    txtTEffdate.SetFocus
Else
  strsql = "Select BNFCode From BNFNew where PLNCode ='" & RemStr(cboTPlnCode) & "' and INSCode ='" & Left(cboTInsCode, 1) & "'" & _
            " AND BNFEffDate ='" & Format(txtTEffdate, "mm/dd/yyyy") & "'"
  BNF.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
  
  If BNF.RecordCount > 0 And EditFlag = False Then
        MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
        BNF.Close
        cboTHLTCode.SetFocus
        cboTHLTCode.SelStart = 0
        cboTHLTCode.SelLength = Len(cboTHLTCode)
        Set BNF = Nothing
        Exit Sub
  Else
      RecordSaved = MsgBox("Do you want to save this record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
             For currentRecord = 1 To recordInBNF - 1 Step 1
               
                With SaveBNF
                FGrid.Row = currentRecord

                FGrid.Col = 2
                Calculation = FGrid.Text
                
                FGrid.Col = 5
                Pervisit = FGrid.Text
                
                FGrid.Col = 6
                Percentage = FGrid.Text

                If Calculation = "1" Then
                    If Pervisit <> "" Or Percentage <> "" Then
                        If EditFlag = False Then
                           SaveBNF.Open " Select * from BNFNew ", MediConnCISDB, adOpenKeyset, adLockOptimistic
                           SaveBNF.AddNew
                        End If
                        !INSCode = Trim(Left(ChkStr(cboTInsCode.Text), 1))
                        !PLNCode = Trim(Left(ChkStr(cboTPlnCode.Text), 4))
                        !HLTCode = (ChkStr(Mid(cboTHLTCode.Text, 1, 1)))
                        !BNFEffDate = txtTEffdate
                        !BNFLastUpdateUser = Logon
                        !BNFLastUpdateDate = Date
                        FGrid.Col = 0
                        !BNFCode = FGrid.Text
                        FGrid.Col = 2
                        If Trim(FGrid.Text) <> "" Then
                            !BNFCalculation = FGrid.Text
                        End If
                        FGrid.Col = 3
                        If Trim(FGrid.Text) <> "" Then
                            !BNFNoOfDays = FGrid.Text
                        End If
                        FGrid.Col = 4
                        If Trim(FGrid.Text) <> "" Then
                            !BNFNoOfAnn = FGrid.Text
                        End If
                        FGrid.Col = 5
                        !BNFRMPerVisit = ChkStr(FGrid.Text)
                        FGrid.Col = 6
                        If Trim(FGrid.Text) <> "" Then
                            !BNFPercentage = ChkStr(FGrid.Text)
                        End If
                        SaveBNF.update
                        SaveBNF.Close
                        Set SaveBNF = Nothing
                    End If
                                
                ElseIf Calculation = "2" Then
                    If Pervisit <> "" Or Percentage <> "" Then
                        If EditFlag = False Then
                           SaveBNF.Open " Select * from BNFNew ", MediConnCISDB, adOpenKeyset, adLockOptimistic
                           SaveBNF.AddNew
                        End If
                        !INSCode = Trim(Left(ChkStr(cboTInsCode.Text), 1))
                        !PLNCode = Trim(Left(ChkStr(cboTPlnCode.Text), 4))
                        !HLTCode = (ChkStr(Mid(cboTHLTCode.Text, 1, 1)))
                        !BNFEffDate = txtTEffdate
                        !BNFLastUpdateUser = Logon
                        !BNFLastUpdateDate = Date
                        FGrid.Col = 0
                        !BNFCode = FGrid.Text
                        FGrid.Col = 2
                        If Trim(FGrid.Text) <> "" Then
                            !BNFCalculation = FGrid.Text
                        End If
                        FGrid.Col = 3
                        If Trim(FGrid.Text) <> "" Then
                            !BNFNoOfDays = FGrid.Text
                        End If
                        FGrid.Col = 4
                        If Trim(FGrid.Text) <> "" Then
                            !BNFNoOfAnn = FGrid.Text
                        End If
                        FGrid.Col = 5
                        !BNFRMPerVisit = ChkStr(FGrid.Text)
                        FGrid.Col = 6
                        If Trim(FGrid.Text) <> "" Then
                            !BNFPercentage = ChkStr(FGrid.Text)
                        End If
                        SaveBNF.update
                        SaveBNF.Close
                        Set SaveBNF = Nothing
                    End If
                    
                ElseIf Calculation = "3" Then
                    If EditFlag = False Then
                       SaveBNF.Open " Select * from BNFNew ", MediConnCISDB, adOpenKeyset, adLockOptimistic
                       SaveBNF.AddNew
                    End If
                    !INSCode = Trim(Left(ChkStr(cboTInsCode.Text), 1))
                    !PLNCode = Trim(Left(ChkStr(cboTPlnCode.Text), 4))
                    !HLTCode = (ChkStr(Mid(cboTHLTCode.Text, 1, 1)))
                    !BNFEffDate = txtTEffdate
                    !BNFLastUpdateUser = Logon
                    !BNFLastUpdateDate = Date
                    FGrid.Col = 0
                    !BNFCode = FGrid.Text
                    FGrid.Col = 2
                    If Trim(FGrid.Text) <> "" Then
                        !BNFCalculation = FGrid.Text
                    End If
                    FGrid.Col = 3
                    If Trim(FGrid.Text) <> "" Then
                        !BNFNoOfDays = FGrid.Text
                    End If
                    FGrid.Col = 4
                    If Trim(FGrid.Text) <> "" Then
                        !BNFNoOfAnn = FGrid.Text
                    End If
                    FGrid.Col = 5
                    !BNFRMPerVisit = ChkStr(FGrid.Text)
                    FGrid.Col = 6
                    If Trim(FGrid.Text) <> "" Then
                        !BNFPercentage = ChkStr(FGrid.Text)
                    End If
                    SaveBNF.update
                    SaveBNF.Close
                    Set SaveBNF = Nothing
                    
                ElseIf Calculation = "4" Then
                    If Pervisit <> "" Then
                        If EditFlag = False Then
                           SaveBNF.Open " Select * from BNFNew ", MediConnCISDB, adOpenKeyset, adLockOptimistic
                           SaveBNF.AddNew
                        End If
                        !INSCode = Trim(Left(ChkStr(cboTInsCode.Text), 1))
                        !PLNCode = Trim(Left(ChkStr(cboTPlnCode.Text), 4))
                        !HLTCode = (ChkStr(Mid(cboTHLTCode.Text, 1, 1)))
                        !BNFEffDate = txtTEffdate
                        !BNFLastUpdateUser = Logon
                        !BNFLastUpdateDate = Date
                        FGrid.Col = 0
                        !BNFCode = FGrid.Text
                        FGrid.Col = 2
                        If Trim(FGrid.Text) <> "" Then
                            !BNFCalculation = FGrid.Text
                        End If
                        FGrid.Col = 3
                        If Trim(FGrid.Text) <> "" Then
                            !BNFNoOfDays = FGrid.Text
                        End If
                        FGrid.Col = 4
                        If Trim(FGrid.Text) <> "" Then
                            !BNFNoOfAnn = FGrid.Text
                        End If
                        FGrid.Col = 5
                        !BNFRMPerVisit = ChkStr(FGrid.Text)
                        FGrid.Col = 6
                        If Trim(FGrid.Text) <> "" Then
                            !BNFPercentage = ChkStr(FGrid.Text)
                        End If
                        SaveBNF.update
                        SaveBNF.Close
                        Set SaveBNF = Nothing
                    End If
                End If
                End With
            Next
                        
            If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
            End If
       End If
       EditFlag = False
   End If
End If
  
Exit Sub
UpdateErr:
    
  MsgBox Err.Description
  
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub BNFcombolisting()
Const METHOD_NAME = "BNFcombolisting"
On Error GoTo debugs

Dim INS As New ADODB.Recordset
Dim PLN As New ADODB.Recordset
Dim subCat As New ADODB.Recordset
Dim cnt As Integer

    FGrid.Row = 0
    FGrid.ColWidth(1) = 3000
    FGrid.ColWidth(2) = 1200
    FGrid.ColWidth(3) = 1500
    FGrid.ColWidth(4) = 1600
    FGrid.ColWidth(5) = 1200
    FGrid.ColWidth(6) = 1200
    FGrid.Col = 0
    FGrid = "Cat Code"
    FGrid.Col = 1
    FGrid = "Cat Name"
    FGrid.Col = 2
    FGrid = "Calculation"
    FGrid.Col = 3
    FGrid = "No of Visit Per day"
    FGrid.Col = 4
    FGrid = "No of Visit Per Annum"
    FGrid.Col = 5
    FGrid = "RM Per Visit"
    FGrid.Col = 6
    FGrid = "Percentage"
    
    INS.Open "select INSCOde , INSDesc from INS", MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic
    PLN.Open "select PLNCode From PLN Order BY PLNCode", MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic 'adOpenKeyset, adLockOptimistic
    subCat.Open "select BNFCATCode, BNFCATDescription from BNFCategoryNew ", MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    cboTHLTCode.Clear
    cboTPlnCode.Clear
    cboTInsCode.Clear
    
    cboTHLTCode.AddItem "S" & " - " & "Sihat"
    cboTHLTCode.AddItem "N" & " - " & "Non-Sihat"

    Do Until INS.EOF
         cboTInsCode.AddItem INS!INSCode & " - " & INS!InsDesc
         INS.MoveNext
    Loop
          
    Do Until PLN.EOF
         cboTPlnCode.AddItem PLN!PLNCode
         PLN.MoveNext
    Loop
    
    cnt = 1
    FGrid.Rows = subCat.RecordCount + 1
    recordInBNF = subCat.RecordCount + 1
    Do Until subCat.EOF
         FGrid.Col = 0
         FGrid.Row = cnt
         FGrid.Text = subCat!BNFCatCode
    
         FGrid.Col = 1
         FGrid.Row = cnt
         FGrid.Text = subCat!BNFCATDescription
         cnt = cnt + 1
         subCat.MoveNext
    Loop
        
    INS.Close
    Set INS = Nothing
    PLN.Close
    Set PLN = Nothing
    subCat.Close
    Set subCat = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub BNFCombolisting2()
Const METHOD_NAME = "BNFCombolisting2"
On Error GoTo debugs
Dim HLT As New ADODB.Recordset
Dim INS As New ADODB.Recordset
Dim PLN As New ADODB.Recordset
Dim subCat As New ADODB.Recordset
Dim PAY As New ADODB.Recordset
Dim GRPCOM As New ADODB.Recordset

Screen.MousePointer = vbHourglass

HLT.Open "select HLTCode, HLTDesc from HLT", MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic
INS.Open "select INSCOde , INSDesc from INS", MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic 'adOpenKeyset, adLockOptimistic
PLN.Open "select PLNCode From PLN", MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic 'adOpenKeyset, adLockOptimistic
subCat.Open "select BNFCATCode, BNFCATDescription from BNFCategoryNew ", MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic 'adOpenKeyset, adLockOptimistic
PAY.Open "select PAYCOde, PAYCompanyName from PAY", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic 'adOpenKeyset, adLockOptimistic
GRPCOM.Open "select Distinct GRPCompany from PLN ", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic 'adOpenKeyset, adLockOptimistic


cboBHLTCode.Clear
cboBPLNCode.Clear
cboCatCode.Clear
cboBINSCode.Clear
CboPayorCode.Clear
CboGrpCom.Clear

Do Until HLT.EOF
     cboBHLTCode.AddItem HLT!HLTCode & " - " & HLT!HLTDesc
     HLT.MoveNext
Loop
      
Do Until INS.EOF
     cboBINSCode.AddItem INS!INSCode & " - " & INS!InsDesc
     INS.MoveNext
Loop
      
Do Until PLN.EOF
     cboBPLNCode.AddItem PLN!PLNCode '& "-" & PLN!PLNDescription
     PLN.MoveNext
Loop

Do Until PAY.EOF
     CboPayorCode.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
     PAY.MoveNext
Loop

Do Until GRPCOM.EOF
     If Len(GRPCOM!GRPCompany) Then
        CboGrpCom.AddItem GRPCOM!GRPCompany
     End If
     GRPCOM.MoveNext
Loop

Do Until subCat.EOF
     cboCatCode.AddItem subCat!BNFCatCode & " -- " & subCat!BNFCATDescription
     subCat.MoveNext
Loop

    HLT.Close
    Set HLT = Nothing
    INS.Close
    Set INS = Nothing
    PLN.Close
    Set PLN = Nothing
    subCat.Close
    Set subCat = Nothing
    PAY.Close
    Set PAY = Nothing
    GRPCOM.Close
    Set GRPCOM = Nothing
    
Screen.MousePointer = Default
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

  '  MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
  '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
        Select Case SSTab1.Tab
            Case 0
                'Call ClearForm(Me)
                'CmdAdd.Enabled = True
                'CmdEdit.Enabled = False
                'CmdDelete.Enabled = False
                'CmdRefresh.Enabled = False
                'CmdSave.Enabled = True
                CmdPrint.Enabled = True
                'CmdSearch.Enabled = True
                'lstCATList.ListItems.Clear
                txtTotalRecord = 0
            Case 1
                Call ClearForm(Me)
                cboTHLTCode.Enabled = True
                cboTInsCode.Enabled = True
                cboTPlnCode.Enabled = True
                txtTEffdate.Enabled = True
                CmdEdit.Enabled = False
                CmdDelete.Enabled = False
                CmdRefresh.Enabled = False
                CmdSearch.Enabled = True
                CmdAdd.Enabled = True
                CmdPrint.Enabled = False
                CmdSave.Enabled = True
                If cboTInsCode.ListCount = 0 Then
                    Call BNFcombolisting
                End If
                txtTotalRecord = 0
            Case 2
            
                CmdAdd.Enabled = False
                CmdEdit.Enabled = False
                CmdDelete.Enabled = False
                CmdPrint.Enabled = True
                CmdRefresh.Enabled = True
                CmdSearch.Enabled = True
                CmdSave.Enabled = False
                If cboBINSCode.ListCount = 0 Then
                    Call BNFCombolisting2
                End If
                lstBNFList.ListItems.Clear
                txtTotalRecord = 0
            
            End Select
       ' MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
       ' MediConnCISWS.Close
       ' Set MediConnCISWS = Nothing
       ' MediConnCISPayment.Close
       ' Set MediConnCISPayment = Nothing
End Sub

Private Sub searchtBNF()
   
    strAppend = ""
    
    If Len(Trim(cboTHLTCode)) = 0 Then
        MsgBox "Please key in Health Type !", vbCritical
    ElseIf Len(Trim(cboTInsCode)) = 0 Then
        MsgBox "Please key in Insured Type !", vbCritical
    ElseIf Len(Trim(cboTPlnCode)) = 0 Then
        MsgBox "Please key in Plan Code!", vbCritical
    ElseIf txtTEffdate = "__/__/____" Then
        MsgBox "Please key in Effective Date!", vbCritical
    ElseIf IsDate(txtTEffdate) = False Then
        MsgBox "Please key in VALID Effective Date!", vbCritical
    Else
        strAppend = strAppend + " And PLNCode = '" & Trim(Mid(cboTPlnCode, 1, 4)) & "'"
        strAppend = strAppend + " And INSCode = '" & Mid(cboTInsCode, 1, 1) & "'"
        strAppend = strAppend + " And HLTCode = '" & Mid(cboTHLTCode, 1, 1) & "'"
        strAppend = strAppend + " And BNFEffDate = '" & Format(txtTEffdate, "mm/dd/yyyy") & "'"
        'TBNFCriteria = strAppend
        Call tBNFlisting(strAppend)
    End If
End Sub

Private Sub searchBNF()
  Dim strAppend As String
   
    strAppend = ""
    
    If Len(Trim(cboCatCode)) Then
        strAppend = strAppend + " And BNFCode = '" & Trim(Mid(cboCatCode, 1, IIf(InStr(1, cboCatCode, "--") = 0, 10, InStr(1, cboCatCode, "--") - 1))) & "'"
    End If
    If Len(Trim(cboBHLTCode)) Then
        strAppend = strAppend + " And HLTCode = '" & Mid(cboBHLTCode, 1, 1) & "'"
    End If
    If Len(Trim(cboBINSCode)) Then
        strAppend = strAppend + " And INSCode = '" & Mid(cboBINSCode, 1, 1) & "'"
    End If
    If Len(Trim(cboBPLNCode)) Then
        strAppend = strAppend + " And PLNCode = '" & Trim(Mid(cboBPLNCode, 1, 4)) & "'"
    End If
    
    If Len(Trim(txtEffdate)) > 0 And IsDate(txtEffdate) = True Then
        strAppend = strAppend + " And BNFEffDate = '" & Format(txtEffdate, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(CboPayorCode)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayorCode, 1, IIf(InStr(1, CboPayorCode, "-") = 0, 4, InStr(1, CboPayorCode, "-") - 1))) & "'"
    End If

    If Len(Trim(CboGrpCom)) Then
        strAppend = strAppend + " AND GRPCompany like '" & RemStr(Mid(CboGrpCom, 1, IIf(InStr(1, CboGrpCom, "-") = 0, 4, InStr(1, CboGrpCom, "-") - 1))) & "%'"
    End If
    
    If Len(Trim(txtBNFCalculation)) Then
        strAppend = strAppend + " And BNFCalculation = '" & Trim(txtBNFCalculation) & "'"
    End If
    
    If Len(Trim(TxtDay)) Then
        strAppend = strAppend + " And BNFNoofDays = '" & Trim(TxtDay) & "'"
    End If
    
    If Len(Trim(txtRMPerVisit)) Then
        strAppend = strAppend + " And BNFRMPerVisit = '" & Trim(txtRMPerVisit) & "'"
    End If
    
    If Len(Trim(TxtRMPerAnn)) Then
        strAppend = strAppend + " And BNFNoofAnn = '" & Trim(TxtRMPerAnn) & "'"
    End If
    
    If Len(Trim(TxtBNFPercentage)) Then
        strAppend = strAppend + " And BNFPercentage = '" & Trim(TxtBNFPercentage) & "'"
    End If
    
    BNFCriteria = strAppend
    Call BNFListing(strAppend)
    
End Sub

Private Sub BNFListing(Optional strAppend As String)
Const METHOD_NAME = "BNFListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim BNF As ListItem
Dim sql As String
Dim TotalBNFCount As Integer
   
    TotalBNFCount = 0
    lstBNFList.ListItems.Clear
    
    sql = " select BNFCode, PLNCode, HLTCode, INSCode, BNFCATDescription, " & _
      " BNFCalculation, BNFNoOfDays, BNFRMPerVisit, BNFEffDate, " & _
      " BNFPercentage, BNFNoOfAnn, PAYCode, GRPCompany " & _
      " From VIEW_CISBNF_PLN_SubCat where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    
    If rst2.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        
        Do Until rst2.EOF
            
            Set BNF = lstBNFList.ListItems.Add(, , rst2!BNFCode)
                
                If Len(Trim(rst2!BNFCATDescription)) Then
                    BNF.SubItems(1) = Trim(rst2!BNFCATDescription)
                End If
                
                If IsNull(rst2!HLTCode) = False Then
                    BNF.SubItems(2) = rst2!HLTCode
                End If
                
                If Len(Trim(rst2!INSCode)) Then
                    BNF.SubItems(3) = rst2!INSCode
                End If
                
                If IsNull(rst2!PLNCode) = False Then
                    BNF.SubItems(4) = rst2!PLNCode
                End If
                                
                If IsNull(rst2!BNFCalculation) = False Then
                    BNF.SubItems(5) = rst2!BNFCalculation
                End If
                
                If IsNull(rst2!BNFNoOfDays) = False Then
                    BNF.SubItems(6) = rst2!BNFNoOfDays
                End If
                
                If IsNull(rst2!BNFNoOfAnn) = False Then
                    BNF.SubItems(7) = rst2!BNFNoOfAnn
                End If
                
                If IsNull(rst2!BNFRMPerVisit) = False Then
                    BNF.SubItems(8) = rst2!BNFRMPerVisit
                End If
                
                If IsNull(rst2!BNFPercentage) = False Then
                    BNF.SubItems(9) = rst2!BNFPercentage
                End If
                
                If IsNull(rst2!BNFEffDate) = False Then
                    BNF.SubItems(10) = rst2!BNFEffDate
                End If
                
                If IsNull(rst2!PAYCode) = False Then
                    BNF.SubItems(11) = rst2!PAYCode
                End If
                
                If IsNull(rst2!GRPCompany) = False Then
                    BNF.SubItems(12) = rst2!GRPCompany
                End If
                
                rst2.MoveNext
        TotalBNFCount = TotalBNFCount + 1
        txtTotalRecord = TotalBNFCount
        Loop
    
    rst2.Close
    Set rst2 = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub tBNFlisting(Optional strAppend As String)
Const METHOD_NAME = "tBNFlisting"
On Error GoTo debugs

Dim BNF As New ADODB.Recordset
Dim strsql As String
Dim counter As Integer
    
    strsql = " Select BNFCode, BNFNoOFDays, BNFCalculation, " & _
             " BNFNoOfAnn, BNFEffDate, BNFPercentage, " & _
             " BNFRMPerVisit from BNFNew where 0 = 0 " & strAppend
    strsql = strsql & " order by BNFCOde "
    BNF.Open strsql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Call ClearGrid
    
    Do While Not BNF.EOF
         With BNF
            For counter = 1 To recordInBNF - 1 Step 1
             FGrid.Row = counter
             FGrid.Col = 0
                If FGrid.Text = !BNFCode Then
                    FGrid.Col = 2
                    If !BNFCalculation <> "" Then
                        FGrid.Text = !BNFCalculation
                    End If
                    FGrid.Col = 3
                    If !BNFNoOfDays <> "" Then
                        FGrid.Text = !BNFNoOfDays
                    End If
                    FGrid.Col = 4
                    If !BNFNoOfAnn <> "" Then
                        FGrid.Text = !BNFNoOfAnn
                    End If
                    FGrid.Col = 5
                    If !BNFRMPerVisit <> "" Then
                        FGrid.Text = !BNFRMPerVisit
                    End If
                    FGrid.Col = 6
                    If !BNFPercentage <> "" Then
                        FGrid.Text = !BNFPercentage
                    End If
                End If
            Next
            .MoveNext
         End With
    Loop
    BNF.Close
    Set BNF = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub ClearGrid()
    
    For FCol = 2 To 6
        FGrid.Col = FCol
        For Frow = 1 To 29
            FGrid.Row = Frow
            FGrid = ""
        Next
    Next
    
End Sub

Private Sub cboTPlnCode_Change()
    If InStr(cboTPlnCode.Text, "null") Then
       cboTPlnCode.Enabled = False
    End If
End Sub

Private Sub cboTPlnCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboTPlnCode.Text, cboTPlnCode.SelStart) + Chr(KeyAscii) + Mid(cboTPlnCode.Text, cboTPlnCode.SelStart + cboTPlnCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboTPlnCode.ListCount - 1
 
        If NewText = LCase(Left(cboTPlnCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboTPlnCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboTPlnCode.Text = ValidValue
                cboTPlnCode.SelStart = 0
                cboTPlnCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboTHLTCode_Change()
    If InStr(cboTHLTCode.Text, "null") Then
       cboTHLTCode.Enabled = False
    End If
End Sub

Private Sub cboTHLTCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboTHLTCode.Text, cboTHLTCode.SelStart) + Chr(KeyAscii) + Mid(cboTHLTCode.Text, cboTHLTCode.SelStart + cboTHLTCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboTHLTCode.ListCount - 1
 
        If NewText = LCase(Left(cboTHLTCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboTHLTCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboTHLTCode.Text = ValidValue
                cboTHLTCode.SelStart = 0
                cboTHLTCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboTInsCode_Change()
    If InStr(cboTInsCode.Text, "null") Then
       cboTInsCode.Enabled = False
    End If
End Sub

Private Sub cboTInsCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboTInsCode.Text, cboTInsCode.SelStart) + Chr(KeyAscii) + Mid(cboTInsCode.Text, cboTInsCode.SelStart + cboTInsCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboTInsCode.ListCount - 1
 
        If NewText = LCase(Left(cboTInsCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboTInsCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboTInsCode.Text = ValidValue
                cboTInsCode.SelStart = 0
                cboTInsCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub TxtCATCode_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub cboBHLTCode_Click()
Const METHOD_NAME = "cboBHLTCode_Click"
On Error GoTo debugs
Dim sql As String
Dim PLN As New ADODB.Recordset
Dim strsql As String
Dim PLNCode As String

If Mid(cboBHLTCode, 1, 1) = "S" Then
    S = True
    If S = True Then
    
    
    'PLN.Open "PLN", medixmasterconn, adOpenKeyset, adLockOptimistic
    'strsql = "SELECT PRMCode, PRMAmount  From PRM "
    'strsql = strsql & "WHERE (INSCode = '" & InsureType & "') AND (HLTCode = '" & HealthValue & "') "
    'strsql = strsql & "AND (PLNCode = '" & PlanValue & "') AND (AGECode = '" & AgeValue & "') "
    'PLNCode = Mid(PLN!PLNCode, 1, 1)
    
    'PLN.Close
    strsql = "Select PLNCode, HLTCode From PLN "
    strsql = strsql & "Where (PLNCode = '" & PLNCode & "') "
    strsql = strsql & " and  (HLTCode = '" & Mid(Trim(cboBHLTCode), 1, 1) & "')"
    PLN.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    
'  strsql = "Select AGECode, HLTCode From AGE "
'        strsql = strsql & "where (AGECode = '" & TxtAgeCode & "')  and (HLTCode = '" & Trim(CblHealth) & "')"
'        AGE.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    Do While Not PLN.EOF
    'txtPLNcode.AddItem PLN!PLNCode
    cboBPLNCode = PLN!PLNCode
    PLN.MoveNext
    Loop
    
    'ElseIf PLN!HLTCode = "N" Then
    'Do Until PLN.EOF
    'txtPLNcode = PLN!PLNCode
    'PLN.MoveNext
    'Loop

    End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cboBHLTCode_KeyPress(KeyAscii As Integer)
'If secondeditflag = False Then

   Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboBHLTCode.Text, cboBHLTCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboBHLTCode.Text, cboBHLTCode.SelStart + cboBHLTCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboBHLTCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboBHLTCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboBHLTCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboBHLTCode.Text = ValidValue
              cboBHLTCode.SelStart = 0
              cboBHLTCode.SelLength = Len(ValidValue)
            End If
            
'End If

End Sub

Private Sub cboBPLNCode_KeyPress(KeyAscii As Integer)
'If secondeditflag = False Then

Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboBPLNCode.Text, cboBPLNCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboBPLNCode.Text, cboBPLNCode.SelStart + cboBPLNCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboBPLNCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboBPLNCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboBPLNCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboBPLNCode.Text = ValidValue
              cboBPLNCode.SelStart = 0
              cboBPLNCode.SelLength = Len(ValidValue)
              End If
            
'End If

End Sub

Private Sub txtEffdate_LostFocus()
    If IsDate(txtEffdate) Then
    Else
        If txtEffdate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtEffdate.SetFocus
        End If
    End If
End Sub

Private Sub cboBINSCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboBINSCode.Text, cboBINSCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboBINSCode.Text, cboBINSCode.SelStart + cboBINSCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboBINSCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboBINSCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboBINSCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboBINSCode.Text = ValidValue
              cboBINSCode.SelStart = 0
              cboBINSCode.SelLength = Len(ValidValue)
              End If
End Sub

Private Sub cbocatcode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboCatCode.Text, cboCatCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboCatCode.Text, cboCatCode.SelStart + cboCatCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboCatCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboCatCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboCatCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboCatCode.Text = ValidValue
              cboCatCode.SelStart = 0
              cboCatCode.SelLength = Len(ValidValue)
              End If
End Sub

Public Sub ExportListViewtoExcel(lvwList As Control)
    Dim vntHeader As Variant
    Dim vntData As Variant
    Dim X As Long
    Dim Y As Long
    Dim intCol As Integer
    Dim lngRow As Long
    
    'Get Counts
    intCol = CInt(lvwList.ColumnHeaders.Count - 1)
    lngRow = CLng(lvwList.ListItems.Count - 1)
    
    ReDim vntHeader(0)
    ReDim vntData(intCol, lngRow)
    
    'Create Header Array


    For X = 0 To intCol
        ReDim Preserve vntHeader(X)
        vntHeader(X) = lvwList.ColumnHeaders(X + 1).Text
    Next
    
    'Create Data Array


    For X = 0 To lngRow
        vntData(0, X) = lvwList.ListItems.Item(X + 1).Text
        


        For Y = 1 To intCol
            vntData(Y, X) = lvwList.ListItems.Item(X + 1).SubItems(Y)
        Next
    Next
    
    'Create Excel Object
    OpenExcel vntData, vntHeader
    
End Sub


Private Sub ExportRecords(vntData As Variant, vntHeader As Variant, ws As Worksheet)
    Dim lngRow As Long
    Dim intCol As Integer
    Dim varData As Variant
    Dim intStart As Integer
    'Select all Cells and and set the number


    '     format to string
        ws.Cells.Select
        ws.Cells.NumberFormat = "@"
        ws.Cells(1, 1).Select
        lngRow = UBound(vntData, 2) + 2
        intCol = UBound(vntData, 1) + 1
        intStart = 2 'Start from line 2
        'Freeze Row 2
        ws.Rows(2).Select
        ws.Activate
        'ActiveWindow..FreezePanes = True
        'Add Headers


        For X = 1 To intCol
            varData = vntHeader(X - 1)
            ws.Cells(1, X) = CStr(varData)
            ws.Cells(1, X).Font.Bold = True
        Next
        
        'Add Data


        For Y = 1 To intCol


            For X = intStart To lngRow
                varData = vntData(Y - 1, X - 2)


                If IsNull(varData) Then 'Make sure no null values, Excel will choke
                    'Add 1 to Move down a column
                    ws.Cells(X + 1, Y) = ""
                Else
                    ws.Cells(X + 1, Y) = CStr(varData) 'Convert to String to preserve formatting
                End If
            Next
        Next
        'Resize Columns to Fit
        ws.Columns.AutoFit
    End Sub


Private Sub OpenExcel(vntData As Variant, vntHeader As Variant)
    On Error GoTo Err_OpenExcel
    Dim objExcel As Excel.Application
    Dim objWrkSht As Worksheet
    Dim X As Integer
    'Create Excel Object
    Set objExcel = CreateObject("Excel.Application")
    'Add the Workbook
    objExcel.Workbooks.Add
    
    '========== Set Template ===============
        'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\DRCRListing\DRCRListing.XLT")
    '========== Set Template ===============
    
    Set objWrkSht = objExcel.ActiveWorkbook.Sheets(1)
    objExcel.Visible = True
    'Fill the Workbook with data
    ExportRecords vntData, vntHeader, objWrkSht
    objExcel.Interactive = True
    ' Clean up:
    Set objExlSht = Nothing
    Set objExcel = Nothing
Err_OpenExcel:


    Select Case Err
        Case 0
        Case 439
        MsgBox "You must have Microsoft Excel installed on your PC.", vbCritical, "Application Not Found"
        Case Else
        MsgBox Err & ": " & error, vbCritical, "OpenExcel Error"
    End Select
End Sub
