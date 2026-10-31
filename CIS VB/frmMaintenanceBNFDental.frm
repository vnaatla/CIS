VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaintenanceBNFDental 
   Caption         =   "Dental Benefits"
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
      TabIndex        =   24
      Top             =   6840
      Width           =   11655
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   5040
         TabIndex        =   20
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   9960
         TabIndex        =   23
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   9120
         TabIndex        =   22
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   17
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   840
         TabIndex        =   15
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   3360
         TabIndex        =   18
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   4200
         TabIndex        =   19
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   8280
         TabIndex        =   21
         Top             =   270
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   25
      Top             =   360
      Width           =   11535
      _ExtentX        =   20341
      _ExtentY        =   11240
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Tooth Category"
      TabPicture(0)   =   "frmMaintenanceBNFDental.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstCATList"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Tooth Reference"
      TabPicture(1)   =   "frmMaintenanceBNFDental.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label38"
      Tab(1).Control(1)=   "lstToothRefList"
      Tab(1).Control(2)=   "Frame1(5)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Benefit Category"
      TabPicture(2)   =   "frmMaintenanceBNFDental.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(1)"
      Tab(2).Control(1)=   "lstBNFCATList"
      Tab(2).Control(2)=   "Label19"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Benefit Entry"
      TabPicture(3)   =   "frmMaintenanceBNFDental.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(1)=   "FGrid"
      Tab(3).ControlCount=   2
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
         Index           =   1
         Left            =   -74760
         TabIndex        =   50
         Top             =   360
         Width           =   11025
         Begin VB.TextBox TxtBNFCATDesc 
            Height          =   855
            Left            =   1680
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   600
            Width           =   9255
         End
         Begin VB.TextBox TxtBNFCATCode 
            Height          =   285
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   7
            Text            =   "  "
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label18 
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
            TabIndex        =   53
            Top             =   600
            Width           =   1635
         End
         Begin VB.Label Label16 
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
            TabIndex        =   52
            Top             =   240
            Width           =   1560
         End
         Begin VB.Label Label15 
            Caption         =   "Plan Claimable"
            Height          =   330
            Left            =   -1845
            TabIndex        =   51
            Top             =   675
            Width           =   1050
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   -74760
         TabIndex        =   45
         Top             =   360
         Width           =   11055
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
            TabIndex        =   11
            Top             =   600
            Width           =   3615
         End
         Begin VB.ComboBox cboTPlnCode 
            Height          =   315
            Left            =   7200
            TabIndex        =   12
            Top             =   240
            Width           =   3105
         End
         Begin VB.ComboBox cboTHLTCode 
            Height          =   315
            Left            =   1800
            TabIndex        =   10
            Top             =   240
            Width           =   3615
         End
         Begin MSMask.MaskEdBox txtTEffdate 
            Height          =   315
            Left            =   7200
            TabIndex        =   13
            Top             =   600
            Width           =   3105
            _ExtentX        =   5482
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
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
            TabIndex        =   49
            Top             =   600
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
            TabIndex        =   48
            Top             =   240
            Width           =   1275
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
            TabIndex        =   47
            Top             =   240
            Width           =   1335
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
            TabIndex        =   46
            Top             =   600
            Width           =   1485
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
         Index           =   0
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   11025
         Begin VB.TextBox TxtCATCode 
            Height          =   285
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   0
            Text            =   "  "
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TxtCATDescription 
            Height          =   855
            Left            =   1680
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Top             =   600
            Width           =   9255
         End
         Begin VB.Label Label7 
            Caption         =   "Plan Claimable"
            Height          =   330
            Left            =   -1845
            TabIndex        =   42
            Top             =   675
            Width           =   1050
         End
         Begin VB.Label Label2 
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
            TabIndex        =   41
            Top             =   240
            Width           =   1560
         End
         Begin VB.Label Label1 
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
            TabIndex        =   40
            Top             =   600
            Width           =   1635
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
         Left            =   -74760
         TabIndex        =   34
         Top             =   360
         Width           =   11025
         Begin VB.ComboBox CboToothRefCat 
            Height          =   315
            Left            =   5640
            TabIndex        =   4
            Top             =   240
            Width           =   5295
         End
         Begin VB.TextBox TxtToothRefDescription 
            Height          =   855
            Left            =   1800
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   600
            Width           =   9135
         End
         Begin VB.TextBox TxtToothRefCode 
            Height          =   285
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   3
            Text            =   "  "
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label14 
            Caption         =   "Tooth Ref. Cat.  *"
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
            Left            =   4080
            TabIndex        =   44
            Top             =   240
            Width           =   1680
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
            TabIndex        =   37
            Top             =   600
            Width           =   1635
         End
         Begin VB.Label Label36 
            Caption         =   "Tooth Ref. Code  *"
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
            TabIndex        =   36
            Top             =   240
            Width           =   1680
         End
         Begin VB.Label Label34 
            Caption         =   "Plan Claimable"
            Height          =   330
            Left            =   -1845
            TabIndex        =   35
            Top             =   675
            Width           =   1050
         End
      End
      Begin MSComctlLib.ListView lstToothRefList 
         Height          =   3885
         Left            =   -74760
         TabIndex        =   6
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tooth Ref. Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tooth Ref. Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tooth Ref. Category"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstCATList 
         Height          =   3885
         Left            =   240
         TabIndex        =   2
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
         TabIndex        =   14
         Top             =   1560
         Width           =   11055
         _ExtentX        =   19495
         _ExtentY        =   8276
         _Version        =   393216
         Rows            =   40
         Cols            =   7
         FixedCols       =   2
         AllowUserResizing=   3
      End
      Begin MSComctlLib.ListView lstBNFCATList 
         Height          =   3885
         Left            =   -74760
         TabIndex        =   9
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
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "BNF CATEGORY LIST"
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
         TabIndex        =   54
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "TOOTH CATEGORY LIST"
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
         Left            =   120
         TabIndex        =   43
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Caption         =   "TOOTH REFERENCE LIST"
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
         Left            =   -74880
         TabIndex        =   38
         Top             =   2040
         Width           =   2295
      End
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
      TabIndex        =   33
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label txtTotalRecord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5640
      TabIndex        =   32
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
      TabIndex        =   31
      Top             =   7680
      Width           =   4695
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
      TabIndex        =   30
      Top             =   7680
      Width           =   1455
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
      TabIndex        =   29
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
      TabIndex        =   28
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "3 - As Charge"
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
      TabIndex        =   27
      Top             =   7560
      Width           =   2175
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
      TabIndex        =   26
      Top             =   7800
      Width           =   2175
   End
End
Attribute VB_Name = "frmMaintenanceBNFDental"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditFlag As Boolean
Public CATCriteria As String
Public BNFCATCriteria As String
Public ToothRefCriteria As String
Public TBNFCriteria As String
Dim recordInBNF As Integer
Dim FCol, Frow As Integer
Dim strAppend As String
Private m_blnDirection(1 To 3) As Boolean

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
        Case 1
            TxtToothRefCode.SetFocus
            TxtToothRefDescription.Enabled = True
        Case 2
            TxtBNFCATCode.SetFocus
            TxtBNFCATDesc.Enabled = True
    End Select
End Sub
Private Sub Form_Load()
Const METHOD_NAME = "Form_Load"
On Error GoTo debugs

Call ActionLogUpdate(12, "frmMaintenanceBNFDental")


debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdEdit_Click()
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
        Case 1
            If TxtToothRefCode <> "" Then
                TxtToothRefCode.Enabled = False
                CboToothRefCat.Enabled = True
                TxtToothRefDescription.Enabled = True
                TxtToothRefDescription.SetFocus
            End If
        Case 2
            If TxtBNFCATCode <> "" Then
                TxtBNFCATCode.Enabled = False
                TxtBNFCATDesc.Enabled = True
                TxtBNFCATDesc.SetFocus
            End If
    End Select

End Sub

Private Sub CmdPrint_Click()
Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call PrintToothCAT
        Case 1
            Call PrintToothRef
        Case 2
            Call PrintBNFCat
        Case 3
            Call PrintBNFEntry
     End Select
    
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
Screen.MousePointer = Default
End Sub

Private Sub CmdRefresh_Click()
    Select Case SSTab1.Tab
        Case 1
            Call CATCombolisting
        Case 3
            Call BNFcombolisting
    End Select
End Sub

Private Sub CmdSave_Click()
  Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call SaveToothCat
        Case 1
            Call SaveToothRef
        Case 2
            Call SaveBNFCategory
        Case 3
            Call SaveBNF
    End Select
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False

    Screen.MousePointer = Default
End Sub

Private Sub cmdDelete_Click()
Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call DeleteCategory
            CmdAdd.Enabled = True
        Case 1
            Call DeleteToothRef
            CmdAdd.Enabled = True
        Case 2
            Call DeleteDentalBNFCategory
            CmdAdd.Enabled = True
        Case 3
            Call DeleteDentalBNFCategory
            CmdAdd.Enabled = True
    End Select
Screen.MousePointer = Default
End Sub

Private Sub CmdSearch_Click()
Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call SearchCAT
            Call ClearForm(Me)
            TxtCATCode.Enabled = True
            TxtCATDescription.Enabled = True
            TxtCATCode.SetFocus
        Case 1
            Call SearchToothRef
            Call ClearForm(Me)
            TxtToothRefCode.Enabled = True
            TxtToothRefDescription.Enabled = True
            TxtToothRefCode.SetFocus
        Case 2
            Call SearchDentalBNFCAT
            Call ClearForm(Me)
            TxtBNFCATCode.Enabled = True
            TxtBNFCATDesc.Enabled = True
            TxtBNFCATCode.SetFocus
        Case 3
            Call searchtBNF
        End Select
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
        Case 2
            CmdAdd.Enabled = True
        Case 3
            CmdAdd.Enabled = True
    End Select
        
    txtTotalRecord = 0
    Pervisit = "0"
    Percentage = "0"
    lstToothRefList.ListItems.Clear
    lstCATList.ListItems.Clear
    lstBNFCATList.ListItems.Clear
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    Call EntriesEnable(True)
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub


'************************Start EntriesEnable*********************************

Private Sub EntriesEnable(status As Boolean)
     
    TxtToothRefCode.Enabled = status
    CboToothRefCat.Enabled = status
    TxtCATCode.Enabled = status
    TxtBNFCATCode.Enabled = status
    CmdSearch.Enabled = status

End Sub
'************************End EntriesEnable*********************************
     
    ' '''''''''''''''''' Save, edit , Add, delete Category..... ''''''''''''''
Private Sub DeleteToothRef()
Const METHOD_NAME = "DeleteToothRef"
On Error GoTo debugs

Dim rsDeleteCAT As New ADODB.Recordset
Dim rsCheckCAT As New ADODB.Recordset
Dim strsql, sqlCAT As String
Dim DeleteRecord
    
EditFlag = False
       
If TxtToothRefCode = "" Then
   MsgBox "Please Enter Category Code", vbOKOnly + vbCritical, DialogBoxTitle
ElseIf lstToothRefList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlCAT = "Select ToothRefCode From ToothRef where ToothRefCode = '" & Trim(lstToothRefList.SelectedItem.Text) & "'"
    

        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
             strsql = "Delete From ToothRef Where ToothRefCode = '" & ChkStr(lstToothRefList.SelectedItem.Text) & "'"
             rsDeleteCAT.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call ClearForm(Me)
             Call ToothRefListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             TxtToothRefCode.SetFocus
             CmdDelete.Enabled = False
             CmdEdit.Enabled = False
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
    
Private Sub SearchToothRef()
Dim strAppend As String
    
  Screen.MousePointer = vbHourglass
  strAppend = ""
  
  If Len(Trim(TxtToothRefCode)) Then
    strAppend = strAppend + " And ToothRefCode = '" & Trim(TxtToothRefCode) & "'"
  End If
  
  If Len(CboToothRefCat) Then
    strAppend = strAppend + " And ToothRefCat = '" & Mid(CboToothRefCat, 1, 1) & "'"
  End If
  
  If Len(Trim(TxtToothRefDescription)) Then
    strAppend = strAppend + " And ToothRefDesc Like '%" & Trim(TxtToothRefDescription) & "%'"
  End If
    
  ToothRefCriteria = strAppend
  Call ToothRefListing(strAppend)
  Screen.MousePointer = Default
  
End Sub

Private Sub PrintToothCAT()
    RptToothCat.Show
End Sub

Private Sub PrintToothRef()
    RptToothRef.Show
End Sub

Private Sub PrintBNFCat()
    RptDentalBNFCat.Show
End Sub

Private Sub PrintBNFEntry()
    RptDentalBNFEntry.Show
End Sub

Private Sub SaveToothRef()
Const METHOD_NAME = "SaveToothRef"
On Error GoTo debugs
Dim RecordSaved
Dim ToothRef As New ADODB.Recordset
Dim SaveToothRef As New ADODB.Recordset
Dim strsql As String
   
    If TxtToothRefCode = "" Then
        MsgBox "Please Enter Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtToothRefCode.SetFocus
    ElseIf CboToothRefCat = "" Then
        MsgBox "Please Enter Category", vbOKOnly + vbCritical, DialogBoxTitle
        CboToothRefCat.SetFocus
    ElseIf TxtToothRefDescription = "" Then
        MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtToothRefDescription.SetFocus
    Else
        strsql = "Select * From ToothRef where ToothRefCode =  '" & Trim(TxtToothRefCode) & "'"
        ToothRef.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If ToothRef.RecordCount > 0 And EditFlag = False Then
           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           ToothRef.Close
           TxtToothRefCode.SetFocus
           TxtToothRefCode.SelStart = 0
           TxtToothRefCode.SelLength = Len(TxtToothRefCode)
           Set ToothRef = Nothing
           Exit Sub
        Else
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           If RecordSaved = vbYes Then
              If EditFlag = False Then
                 SaveToothRef.Open "ToothRef", MediConnCISDB, adOpenKeyset, adLockOptimistic
                 SaveToothRef.AddNew
              Else
                 strsql = "Select * From ToothRef Where ToothRefCode = '" & ChkStr(TxtToothRefCode) & "'"
                 SaveToothRef.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
              End If
              
              With SaveToothRef
                !ToothRefCode = ChkStr(TxtToothRefCode)
                !ToothRefDesc = ChkStr(TxtToothRefDescription)
                !ToothRefCat = Mid(CboToothRefCat, 1, 1)
                !ToothLastUpdateUser = Logon
                !ToothLastUpdateDate = Date
                 
              End With
              SaveToothRef.update
            
             Call ClearForm(Me)
             Call ToothRefListing(ToothRefCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             SaveToothRef.Close
             Set SaveToothRef = Nothing
        End If
        EditFlag = False
     End If
   End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub ToothRefListing(Optional strAppend As String)
Const METHOD_NAME = "ToothRefListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim ToothRef As ListItem
Dim sql As String
Dim TotalCATCount As Integer
   
    TotalCATCount = 0
    lstToothRefList.ListItems.Clear
    
    sql = " Select * From VIEW_ToothRef_Category where 0 = 0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            With rst2
            Set ToothRef = lstToothRefList.ListItems.Add(, , rst2!ToothRefCode)
                ToothRef.SubItems(1) = Trim(rst2!ToothRefDesc)
                ToothRef.SubItems(2) = Trim(rst2!ToothRefCatDesc)
                rst2.MoveNext
            End With
        TotalCATCount = TotalCATCount + 1
        txtTotalRecord = TotalCATCount
        Loop
    End If
    ToothRefCriteria = ""
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

Private Sub lstBNFCATList_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If lstBNFCATList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        TxtBNFCATCode = lstBNFCATList.SelectedItem.Text
        TxtBNFCATCode.Enabled = False
        TxtBNFCATDesc = lstBNFCATList.SelectedItem.SubItems(1)
        TxtBNFCATDesc.Enabled = False
            
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstBNFCATList_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstBNFCATList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        TxtBNFCATCode = lstBNFCATList.SelectedItem.Text
        TxtBNFCATCode.Enabled = False
        TxtBNFCATDesc = lstBNFCATList.SelectedItem.SubItems(1)
        TxtBNFCATDesc.Enabled = False
            
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstCATList_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstCATList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        TxtCATCode = lstCATList.SelectedItem.Text
        TxtCATCode.Enabled = False
        TxtCATDescription = lstCATList.SelectedItem.SubItems(1)
        TxtCATDescription.Enabled = False
            
        CmdDelete.Enabled = True
    
    End If
End Sub

Private Sub lstCATList_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstCATList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        TxtCATCode = lstCATList.SelectedItem.Text
        TxtCATCode.Enabled = False
        TxtCATDescription = lstCATList.SelectedItem.SubItems(1)
        TxtCATDescription.Enabled = False
            
        CmdDelete.Enabled = True
    
    End If
End Sub

Private Sub lstToothRefList_Click()
    
    If lstToothRefList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        TxtToothRefCode = lstToothRefList.SelectedItem.Text
        TxtToothRefCode.Enabled = False
        TxtToothRefDescription = lstToothRefList.SelectedItem.SubItems(1)
        TxtToothRefDescription.Enabled = False
        CboToothRefCat = lstToothRefList.SelectedItem.SubItems(2)
        CboToothRefCat.Enabled = False
            
        CmdDelete.Enabled = True
    
    End If
    
End Sub

Private Sub lstToothRefList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstToothRefList, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
    Case 2
        SortListView lstToothRefList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstToothRefList, ColumnHeader.Index, ldtString, m_blnDirection(3)
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
  strsql = "Select BNFCode From DentalBNF where PLNCode ='" & RemStr(cboTPlnCode) & "' and INSCode ='" & Left(cboTInsCode, 1) & "'" & _
            " AND BNFEffDate ='" & Format(txtTEffdate, "mm/dd/yyyy") & "'"
  BNF.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
  
  'If BNF.RecordCount > 0 And EditFlag = False Then
  If BNF.RecordCount > 0 Then
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
                        'If EditFlag = False Then
                           SaveBNF.Open " Select * from DentalBNF ", MediConnCISDB, adOpenKeyset, adLockOptimistic
                           SaveBNF.AddNew
                        'End If
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
                        'If EditFlag = False Then
                           SaveBNF.Open " Select * from DentalBNF ", MediConnCISDB, adOpenKeyset, adLockOptimistic
                           SaveBNF.AddNew
                        'End If
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
                    'If EditFlag = False Then
                       SaveBNF.Open " Select * from DentalBNF ", MediConnCISDB, adOpenKeyset, adLockOptimistic
                       SaveBNF.AddNew
                    'End If
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
                        'If EditFlag = False Then
                           SaveBNF.Open " Select * from DentalBNF ", MediConnCISDB, adOpenKeyset, adLockOptimistic
                           SaveBNF.AddNew
                        'End If
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
                        
            'If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            'Else
                'MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
            'End If
       End If
       'EditFlag = False
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
Dim Cnt As Integer

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
    subCat.Open "select DentalBNFCATCode, DentalBNFCATDesc from DentalBNFCategory ", MediConnCISDB, adOpenKeyset, adLockOptimistic
    
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
    
    Cnt = 1
    FGrid.Rows = subCat.RecordCount + 1
    recordInBNF = subCat.RecordCount + 1
    Do Until subCat.EOF
         FGrid.Col = 0
         FGrid.Row = Cnt
         FGrid.Text = subCat!DentalBNFCATCode
    
         FGrid.Col = 1
         FGrid.Row = Cnt
         FGrid.Text = subCat!DentalBNFCATDesc
         Cnt = Cnt + 1
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

Private Sub CATCombolisting()
Const METHOD_NAME = "CATCombolisting"
On Error GoTo debugs

Dim CAT As New ADODB.Recordset


Screen.MousePointer = vbHourglass

CAT.Open "Select ToothRefCatCode, ToothRefCatDesc from ToothRefCategory", MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic

CboToothRefCat.Clear

Do Until CAT.EOF
     CboToothRefCat.AddItem CAT!ToothRefCatCode & " - " & Trim(CAT!ToothRefCatDesc)
     CAT.MoveNext
Loop
      

CAT.Close
Set CAT = Nothing
    
Screen.MousePointer = Default

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub lstToothRefList_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If lstToothRefList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        TxtToothRefCode = lstToothRefList.SelectedItem.Text
        TxtToothRefCode.Enabled = False
        TxtToothRefDescription = lstToothRefList.SelectedItem.SubItems(1)
        TxtToothRefDescription.Enabled = False
        CboToothRefCat = lstToothRefList.SelectedItem.SubItems(2)
        CboToothRefCat.Enabled = False
            
        CmdDelete.Enabled = True
    
    End If
End Sub

Private Sub lstToothRefList_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If lstToothRefList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        TxtToothRefCode = lstToothRefList.SelectedItem.Text
        TxtToothRefCode.Enabled = False
        TxtToothRefDescription = lstToothRefList.SelectedItem.SubItems(1)
        TxtToothRefDescription.Enabled = False
        CboToothRefCat = lstToothRefList.SelectedItem.SubItems(2)
        CboToothRefCat.Enabled = False
            
        CmdDelete.Enabled = True
    
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
    Case 0
        txtTotalRecord = 0
        CmdPrint.Enabled = True
        CmdRefresh.Enabled = False
    Case 1
        Call CATCombolisting
        txtTotalRecord = 0
        CmdPrint.Enabled = True
        CmdRefresh.Enabled = True
    Case 2
        txtTotalRecord = 0
        CmdPrint.Enabled = True
        CmdRefresh.Enabled = False
    Case 3
        Call ClearForm(Me)
        cboTHLTCode.Enabled = True
        cboTInsCode.Enabled = True
        cboTPlnCode.Enabled = True
        txtTEffdate.Enabled = True
        CmdEdit.Enabled = False
        CmdDelete.Enabled = False
        CmdRefresh.Enabled = True
        CmdSearch.Enabled = True
        CmdAdd.Enabled = True
        CmdPrint.Enabled = True
        CmdSave.Enabled = True
        'If cboTInsCode.ListCount = 0 Then
            Call BNFcombolisting
        'End If
        txtTotalRecord = 0
    
    End Select

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
        TBNFCriteria = strAppend
        Call tBNFlisting(strAppend)
    End If
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
    
    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If rst2.RecordCount = 0 Then
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
             " BNFRMPerVisit from DentalBNF where 0 = 0 " & strAppend
    strsql = strsql & " order by BNFCode "
    BNF.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    Call ClearGrid
    If BNF.RecordCount Then
    Do Until BNF.EOF
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
    End If
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
        For Frow = 1 To 12
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

Private Sub TxtToothRefCode_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub CboToothRefCat_Change()
    If InStr(CboToothRefCat.Text, "null") Then
       CboToothRefCat.Enabled = False
    End If
End Sub

Private Sub CboToothRefCat_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboToothRefCat.Text, CboToothRefCat.SelStart) + Chr(KeyAscii) + Mid(CboToothRefCat.Text, CboToothRefCat.SelStart + CboToothRefCat.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboToothRefCat.ListCount - 1
 
        If NewText = LCase(Left(CboToothRefCat.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboToothRefCat.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboToothRefCat.Text = ValidValue
                CboToothRefCat.SelStart = 0
                CboToothRefCat.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub SearchCAT()
Dim strAppend As String
    
  Screen.MousePointer = vbHourglass
  strAppend = ""
  
  If Len(Trim(TxtCATCode)) Then
    strAppend = strAppend + " And ToothRefCatCode = '" & Trim(TxtCATCode) & "'"
  End If
  
  If Len(Trim(TxtCATDescription)) Then
    strAppend = strAppend + " And ToothRefCatDesc like '%" & Trim(TxtCATDescription) & "%'"
  End If
    
  CATCriteria = strAppend
  Call ToothCATListing(strAppend)
  Screen.MousePointer = Default
  
End Sub

Private Sub SaveToothCat()
Const METHOD_NAME = "SaveToothCat"
On Error GoTo debugs

Dim RecordSaved
Dim CAT As New ADODB.Recordset
Dim SaveToothCat As New ADODB.Recordset
Dim strsql As String
   
    If TxtCATCode = "" Then
        MsgBox "Please Enter Category Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtCATCode.SetFocus
    ElseIf TxtCATDescription = "" Then
        MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtCATDescription.SetFocus
    Else
        strsql = "Select * From ToothRefCategory where ToothRefCatCode =  '" & Trim(TxtCATCode) & "'"
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
                 SaveToothCat.Open "ToothRefCategory", MediConnCISDB, adOpenKeyset, adLockOptimistic
                 SaveToothCat.AddNew
              Else
                 strsql = "Select * From ToothRefCategory Where ToothRefCatCode = '" & ChkStr(TxtCATCode) & "'"
                 SaveToothCat.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
              End If
              
              With SaveToothCat
                !ToothRefCatCode = ChkStr(TxtCATCode)
                !ToothRefCatDesc = ChkStr(TxtCATDescription)
                !ToothRefCatLastUpdateUser = Logon
                !ToothRefCatLastUpdateDate = Date
                 
              End With
              SaveToothCat.update
            
             Call ClearForm(Me)
             Call ToothCATListing(CATCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             SaveToothCat.Close
             Set SaveToothCat = Nothing
        End If
        EditFlag = False
     End If
   End If
   
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub ToothCATListing(Optional strAppend As String)
Const METHOD_NAME = "ToothCATListing"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim CAT As ListItem
Dim sql As String
Dim TotalCATCount As Integer
   
    TotalCATCount = 0
    lstCATList.ListItems.Clear
    
    sql = " select ToothRefCatCode, ToothRefCatDesc " & _
          " From ToothRefCategory where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            With rst2
            Set CAT = lstCATList.ListItems.Add(, , rst2!ToothRefCatCode)
                CAT.SubItems(1) = Trim(rst2!ToothRefCatDesc)
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

Private Sub lstCATList_Click()
    
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
        SortListView lstCATList, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
    Case 2
        SortListView lstCATList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    End Select
End Sub

Private Sub DeleteCategory()
Const METHOD_NAME = "DeleteCategory"
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
   sqlCAT = "Select ToothRefCatCode From ToothRefCategory where ToothRefCatCode = '" & Trim(lstCATList.SelectedItem.Text) & "'"
    

        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
             strsql = "Delete From ToothRefCategory Where ToothRefCatCode = '" & ChkStr(lstCATList.SelectedItem.Text) & "'"
             rsDeleteCAT.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call ClearForm(Me)
             Call ToothCATListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             TxtCATCode.SetFocus
             CmdDelete.Enabled = False
             CmdEdit.Enabled = False
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

Private Sub lstBNFCATList_Click()
    
    If lstBNFCATList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        TxtBNFCATCode = lstBNFCATList.SelectedItem.Text
        TxtBNFCATCode.Enabled = False
        TxtBNFCATDesc = lstBNFCATList.SelectedItem.SubItems(1)
        TxtBNFCATDesc.Enabled = False
            
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstBNFCATList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstBNFCATList, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
    Case 2
        SortListView lstBNFCATList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    End Select
End Sub

Private Sub DeleteDentalBNFCategory()
Const METHOD_NAME = "DeleteDentalBNFCategory"
On Error GoTo debugs

Dim rsDeleteCAT As New ADODB.Recordset
Dim rsCheckCAT As New ADODB.Recordset
Dim strsql, sqlCAT As String
Dim DeleteRecord
    
EditFlag = False
       
If TxtBNFCATCode = "" Then
   MsgBox "Please Enter Category Code", vbOKOnly + vbCritical, DialogBoxTitle
ElseIf lstBNFCATList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlCAT = "Select DentalBNFCATCode From DentalBNFCategory where DentalBNFCATCode = '" & Trim(lstBNFCATList.SelectedItem.Text) & "'"
    

        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
             strsql = "Delete From DentalBNFCategory Where DentalBNFCATCode = '" & ChkStr(lstBNFCATList.SelectedItem.Text) & "'"
             rsDeleteCAT.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call ClearForm(Me)
             Call DentalBNFCATListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             TxtBNFCATCode.SetFocus
             CmdDelete.Enabled = False
             CmdEdit.Enabled = False
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

Private Sub DentalBNFCATListing(Optional strAppend As String)
Const METHOD_NAME = "DentalBNFCATListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim CAT As ListItem
Dim sql As String
Dim TotalCATCount As Integer
   
    TotalCATCount = 0
    lstBNFCATList.ListItems.Clear
    
    sql = " select DentalBNFCATCode, DentalBNFCATDesc " & _
          " From DentalBNFCategory where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            With rst2
            Set CAT = lstBNFCATList.ListItems.Add(, , rst2!DentalBNFCATCode)
                CAT.SubItems(1) = Trim(rst2!DentalBNFCATDesc)
                rst2.MoveNext
            End With
        TotalCATCount = TotalCATCount + 1
        txtTotalRecord = TotalCATCount
        Loop
    End If
    BNFCATCriteria = ""
    rst2.Close
    Set rst2 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SearchDentalBNFCAT()
Dim strAppend As String
    
  Screen.MousePointer = vbHourglass
  strAppend = ""
  
  If Len(Trim(TxtBNFCATCode)) Then
    strAppend = strAppend + " And DentalBNFCATCode = '" & Trim(TxtBNFCATCode) & "'"
  End If
  
  If Len(Trim(TxtBNFCATDesc)) Then
    strAppend = strAppend + " And DentalBNFCATDesc like '%" & Trim(TxtBNFCATDesc) & "%'"
  End If
    
  BNFCATCriteria = strAppend
  Call DentalBNFCATListing(strAppend)
  Screen.MousePointer = Default
  
End Sub

Private Sub SaveBNFCategory()
Const METHOD_NAME = "SaveBNFCategory"
On Error GoTo debugs

Dim RecordSaved
Dim CAT As New ADODB.Recordset
Dim SaveToothCat As New ADODB.Recordset
Dim strsql As String
   
    If TxtBNFCATCode = "" Then
        MsgBox "Please Enter Category Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtBNFCATCode.SetFocus
    ElseIf TxtBNFCATDesc = "" Then
        MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtBNFCATDesc.SetFocus
    Else
        strsql = "Select * From DentalBNFCategory where DentalBNFCATCode =  '" & Trim(TxtBNFCATCode) & "'"
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
                 SaveToothCat.Open "DentalBNFCategory", MediConnCISDB, adOpenKeyset, adLockOptimistic
                 SaveToothCat.AddNew
              Else
                 strsql = "Select * From DentalBNFCategory Where DentalBNFCATCode = '" & ChkStr(TxtBNFCATCode) & "'"
                 SaveToothCat.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
              End If
              
              With SaveToothCat
                !DentalBNFCATCode = ChkStr(TxtBNFCATCode)
                !DentalBNFCATDesc = ChkStr(TxtBNFCATDesc)
                !DentalBNFCATLastUpdateUser = Logon
                !DentalBNFCATLastUpdateDate = Date
                 
              End With
              SaveToothCat.update
            
             Call ClearForm(Me)
             Call DentalBNFCATListing(CATCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             SaveToothCat.Close
             Set SaveToothCat = Nothing
        End If
        EditFlag = False
     End If
   End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub
