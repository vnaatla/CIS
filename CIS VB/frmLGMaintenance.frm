VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLGMaintenance 
   Caption         =   "frmLGMaintenance"
   ClientHeight    =   8604
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   12732
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8604
   ScaleWidth      =   12732
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      Height          =   615
      Left            =   120
      TabIndex        =   57
      Top             =   7920
      Width           =   12555
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
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
         Left            =   8880
         TabIndex        =   64
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Left            =   5520
         TabIndex        =   63
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
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
         Left            =   6360
         TabIndex        =   62
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
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
         Left            =   10560
         TabIndex        =   61
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdCancel 
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
         Height          =   315
         Left            =   11400
         TabIndex        =   60
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
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
         Left            =   9720
         TabIndex        =   59
         Top             =   180
         Width           =   840
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
         Height          =   315
         Left            =   8040
         TabIndex        =   58
         Top             =   180
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7275
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12525
      _ExtentX        =   22098
      _ExtentY        =   12827
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "GL Template Binding"
      TabPicture(0)   =   "frmLGMaintenance.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstLGTemplate"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "GL Type Binding"
      TabPicture(1)   =   "frmLGMaintenance.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label34"
      Tab(1).Control(1)=   "lstGLType"
      Tab(1).Control(2)=   "Frame5"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Template Name Binding"
      TabPicture(2)   =   "frmLGMaintenance.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label11"
      Tab(2).Control(1)=   "lstGLTemplate"
      Tab(2).Control(2)=   "Frame1"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Decline Reason "
      TabPicture(3)   =   "frmLGMaintenance.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label26"
      Tab(3).Control(1)=   "lstDeclineReason"
      Tab(3).Control(2)=   "Frame3"
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame1 
         Height          =   2640
         Left            =   -74880
         TabIndex        =   74
         Top             =   480
         Width           =   12135
         Begin VB.TextBox txttemplateremarks1 
            Height          =   735
            Left            =   2400
            TabIndex        =   94
            Top             =   1680
            Width           =   8175
         End
         Begin VB.TextBox txtTemplateID 
            Height          =   285
            Left            =   2400
            MaxLength       =   15
            TabIndex        =   77
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtTemplate 
            Height          =   285
            Left            =   2400
            MaxLength       =   50
            TabIndex        =   76
            Top             =   600
            Width           =   5775
         End
         Begin VB.TextBox txtTemplateDesc 
            Height          =   525
            Left            =   2400
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   75
            Top             =   960
            Width           =   6495
         End
         Begin VB.Label Label23 
            Caption         =   "Template Remarks *"
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
            Left            =   240
            TabIndex        =   93
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label10 
            Caption         =   "GL Template ID"
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
            Left            =   255
            TabIndex        =   80
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label Label9 
            Caption         =   "GL Template Name *"
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
            TabIndex        =   79
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label7 
            Caption         =   "Template Decription *"
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
            Left            =   240
            TabIndex        =   78
            Top             =   960
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   16
         Top             =   420
         Width           =   11055
         Begin VB.ComboBox cboClnPrePost 
            Height          =   315
            Left            =   9795
            Sorted          =   -1  'True
            TabIndex        =   33
            Top             =   1790
            Width           =   1095
         End
         Begin VB.ComboBox cboClnConversion 
            Height          =   315
            Left            =   9795
            Sorted          =   -1  'True
            TabIndex        =   32
            Top             =   1520
            Width           =   1095
         End
         Begin VB.ComboBox cboClnTax 
            Height          =   315
            Left            =   7440
            Sorted          =   -1  'True
            TabIndex        =   31
            Top             =   1775
            Width           =   1335
         End
         Begin VB.TextBox txtClnHomepage 
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
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   29
            Top             =   1265
            Width           =   3435
         End
         Begin VB.TextBox txtClnEmail 
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
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   28
            Top             =   1010
            Width           =   3435
         End
         Begin VB.TextBox txtClnFax 
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
            Left            =   9600
            MaxLength       =   15
            TabIndex        =   27
            Top             =   755
            Width           =   1290
         End
         Begin VB.TextBox txtClnPhone 
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
            Left            =   7440
            MaxLength       =   15
            TabIndex        =   26
            Top             =   755
            Width           =   1275
         End
         Begin VB.TextBox txtClnCEO 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   25
            Top             =   2030
            Width           =   3975
         End
         Begin VB.TextBox txtClnAdmContact 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   24
            Top             =   1770
            Width           =   3975
         End
         Begin VB.TextBox txtClnContact 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   23
            Top             =   1520
            Width           =   3975
         End
         Begin VB.TextBox txtClnAdd3 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   22
            Top             =   1265
            Width           =   3975
         End
         Begin VB.TextBox txtClnAdd2 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   21
            Top             =   1010
            Width           =   3975
         End
         Begin VB.TextBox txtClnAdd1 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   20
            Top             =   755
            Width           =   3975
         End
         Begin VB.TextBox txtClnName 
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
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   19
            Top             =   500
            Width           =   3975
         End
         Begin VB.TextBox txtClnGrpName 
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
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   18
            Top             =   500
            Width           =   3465
         End
         Begin VB.TextBox txtClnGrpCode 
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
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   17
            Top             =   255
            Width           =   930
         End
         Begin MSMask.MaskEdBox txtClnEffDate 
            Height          =   285
            Left            =   7440
            TabIndex        =   30
            Top             =   1520
            Width           =   1335
            _ExtentX        =   2350
            _ExtentY        =   487
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label20 
            Caption         =   "Adm. Contact Person"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   48
            Top             =   1815
            Width           =   1695
         End
         Begin VB.Label Label19 
            Caption         =   "Initial Group Name  *"
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
            TabIndex        =   47
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "CEO Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   46
            Top             =   2085
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Clinic Group Code   *"
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
            Index           =   2
            Left            =   120
            TabIndex        =   45
            Top             =   255
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Clinic Group Address   *"
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
            TabIndex        =   44
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Cln Grp. Name *"
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
            Index           =   1
            Left            =   6100
            TabIndex        =   43
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6100
            TabIndex        =   42
            Top             =   1040
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "Homepage"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6100
            TabIndex        =   41
            Top             =   1295
            Width           =   870
         End
         Begin VB.Label Label5 
            Caption         =   "Contact Person"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   40
            Top             =   1575
            Width           =   1695
         End
         Begin VB.Label Label17 
            Caption         =   "Fax No"
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
            Left            =   8880
            TabIndex        =   39
            Top             =   800
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Tel No"
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
            Left            =   6100
            TabIndex        =   38
            Top             =   800
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Eff. Date  *"
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
            Index           =   3
            Left            =   6100
            TabIndex        =   37
            Top             =   1550
            Width           =   1050
         End
         Begin VB.Label Label2 
            Caption         =   "Conversion"
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
            Index           =   2
            Left            =   8880
            TabIndex        =   36
            Top             =   1550
            Width           =   930
         End
         Begin VB.Label Label2 
            Caption         =   "Pre&&Post"
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
            Index           =   5
            Left            =   8880
            TabIndex        =   35
            Top             =   1810
            Width           =   810
         End
         Begin VB.Label Label2 
            Caption         =   "R&&B Tax"
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
            Index           =   4
            Left            =   6100
            TabIndex        =   34
            Top             =   1790
            Width           =   810
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1560
         Left            =   -74760
         TabIndex        =   11
         Top             =   420
         Width           =   12135
         Begin VB.TextBox txtGLDesc 
            Height          =   525
            Left            =   2400
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   72
            Top             =   960
            Width           =   5775
         End
         Begin VB.TextBox txtGLType 
            Height          =   285
            Left            =   2400
            MaxLength       =   50
            TabIndex        =   13
            Top             =   600
            Width           =   5775
         End
         Begin VB.TextBox txtGLID 
            Height          =   285
            Left            =   2400
            MaxLength       =   15
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "GL DESCRIPTION *"
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
            Left            =   240
            TabIndex        =   73
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label14 
            Caption         =   "GL TYPE*"
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
            TabIndex        =   15
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label29 
            Caption         =   "GL ID"
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
            Left            =   255
            TabIndex        =   14
            Top             =   240
            Width           =   1800
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3315
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   12075
         Begin VB.ComboBox CmbdependentTemplateName 
            Height          =   315
            Left            =   5160
            TabIndex        =   92
            Top             =   1920
            Width           =   3375
         End
         Begin VB.TextBox txttemplateremarks 
            Height          =   615
            Left            =   1800
            TabIndex        =   90
            Top             =   2400
            Width           =   9735
         End
         Begin VB.ComboBox CmbTemplateName 
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
            Left            =   5160
            Sorted          =   -1  'True
            TabIndex        =   69
            Top             =   1560
            Width           =   3375
         End
         Begin VB.ComboBox cmbLGFormat 
            Height          =   315
            Left            =   9840
            Sorted          =   -1  'True
            TabIndex        =   67
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox cmbPaycode 
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
            TabIndex        =   65
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox cboVIP 
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
            ItemData        =   "frmLGMaintenance.frx":0070
            Left            =   1500
            List            =   "frmLGMaintenance.frx":007A
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   1515
            Width           =   1695
         End
         Begin VB.ComboBox cmbGrpCompany 
            Height          =   315
            Left            =   1560
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   360
            Width           =   4455
         End
         Begin VB.ComboBox cmbGLType 
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
            Left            =   4200
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox cmbPlnCode 
            Height          =   315
            Left            =   6960
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lbldependenttemplate 
            Caption         =   "Dependent Template"
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
            Left            =   3480
            TabIndex        =   91
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label lbltempremarks 
            Caption         =   "Template Remarks"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   89
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "LG Template Name"
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
            Left            =   3480
            TabIndex        =   70
            Top             =   1605
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "LG Format *"
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
            Left            =   8760
            TabIndex        =   68
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblGLType 
            Caption         =   "Pay Code*"
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
            Left            =   240
            TabIndex        =   66
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblGrpCompany 
            Caption         =   "Group Company*"
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
            Index           =   11
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Plan Code*"
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
            Index           =   3
            Left            =   6120
            TabIndex        =   9
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblGLType 
            Caption         =   "GLType *"
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
            Index           =   13
            Left            =   3480
            TabIndex        =   8
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "VIP Indicator"
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
            Index           =   14
            Left            =   240
            TabIndex        =   7
            Top             =   1560
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1560
         Left            =   -74760
         TabIndex        =   1
         Top             =   420
         Width           =   12135
         Begin VB.TextBox txtDecId 
            Height          =   285
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   85
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtDeclineReason 
            Height          =   285
            Left            =   2280
            MaxLength       =   250
            TabIndex        =   84
            Top             =   600
            Width           =   5775
         End
         Begin VB.TextBox txtDecRemarks 
            Height          =   525
            Left            =   2280
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   83
            Top             =   960
            Width           =   6495
         End
         Begin VB.Label Label22 
            Caption         =   "Decline ID"
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
            Left            =   135
            TabIndex        =   88
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label Label21 
            Caption         =   "Decline Reason *"
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
            TabIndex        =   87
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label13 
            Caption         =   "Decline Remarks"
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
            Left            =   120
            TabIndex        =   86
            Top             =   960
            Width           =   1815
         End
      End
      Begin MSComctlLib.ListView lstGrpClinicList 
         Height          =   2835
         Left            =   -74760
         TabIndex        =   49
         Top             =   3180
         Width           =   11055
         _ExtentX        =   19495
         _ExtentY        =   4995
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
            Text            =   "Clinic Group Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clinic Group Name"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstGLType 
         Height          =   4680
         Left            =   -74760
         TabIndex        =   50
         Top             =   2400
         Width           =   12120
         _ExtentX        =   21378
         _ExtentY        =   8255
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
            Text            =   "GL ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "GL TYPE"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "GL Description"
            Object.Width           =   4233
         EndProperty
      End
      Begin MSComctlLib.ListView lstLGTemplate 
         Height          =   2820
         Left            =   240
         TabIndex        =   51
         Top             =   4200
         Width           =   12075
         _ExtentX        =   21294
         _ExtentY        =   4974
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FlatScrollBar   =   -1  'True
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Group Company"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Pay Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Plan Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "GL Type"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "GL Template"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "LG Format"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "VIP Indicator"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView lstDeclineReason 
         Height          =   4200
         Left            =   -74760
         TabIndex        =   52
         Top             =   2640
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   7408
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
            Text            =   "ID"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Decline Reason"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Decline Remarks"
            Object.Width           =   4233
         EndProperty
      End
      Begin MSComctlLib.ListView lstGLTemplate 
         Height          =   3600
         Left            =   -74880
         TabIndex        =   81
         Top             =   3540
         Width           =   12120
         _ExtentX        =   21378
         _ExtentY        =   6350
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
            Text            =   "ID"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Template Name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Template Description"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "TmpRemarks"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label11 
         Caption         =   "GL TEMPLATE  LIST"
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
         TabIndex        =   82
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "GL TEMPLATE BINDING LIST"
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
         TabIndex        =   71
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "CLINIC GROUP LIST"
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
         TabIndex        =   55
         Top             =   2940
         Width           =   1275
      End
      Begin VB.Label Label34 
         Caption         =   "GL TYPE LIST"
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
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Decline Reasons"
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
         TabIndex        =   53
         Top             =   2280
         Width           =   2295
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000015&
      Caption         =   "LG Maintenanace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Index           =   2
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmLGMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SqlQry As String
Dim MaintRS As ADODB.Recordset
Dim LsRecords As ListItem
Dim Rcnt As Long
Dim EdtFlag As Boolean
Dim SearchGrpComp As String

Private Sub cmbGrpCompany_LostFocus()
    Call BindPlanCode
    Call BindPayCode
End Sub



Private Sub cmdAdd_Click()
    Call CmdClear_Click
    EdtFlag = False
    cmdSave.Enabled = True
End Sub
Private Sub cmdCancel_Click()
      Unload Me
End Sub

Private Sub CmdClear_Click()
Select Case SSTab1.Tab
        Case 0 'GL Template
                cmbGrpCompany.Clear
                cmbPaycode.Clear
                cmbGLType.Clear
                cmbPlnCode.Clear
                cmbLGFormat.Clear
                cboVIP.Clear
                CmbTemplateName.Clear
                CmbdependentTemplateName.Clear
                txttemplateremarks = ""
                Call BindGLType
                Call BindGroupCompany
                Call BindGLTemplateName
                Call BindGLTemplateName_Dep
                Call BindCLNLGType
                Call BindVIPInd
        Case 1 'GL Type
                txtGLID = ""
                txtGLType = ""
                txtGLDesc = ""
        Case 2 'Template Name
                txtTemplateID = ""
                txtTemplate = ""
                txtTemplateDesc = ""
                txttemplateremarks1 = ""
        Case 3 'Decline Reason
                txtDecId = ""
                txtDeclineReason = ""
                txtDecRemarks = ""
    End Select
End Sub

Private Sub cmdEdit_Click()
    EdtFlag = True
    cmdSave.Enabled = True
End Sub
Private Sub cmdSave_Click()
Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0 'GL Template
            Call SaveGLTemplate
            SearchGrpComp = cmbGrpCompany
        Case 1 'GL Type
            Call SaveGLType
        Case 2 'Template Name
            Call SaveTemplate
        Case 3 'Decline Reason
            Call SaveDecline
    End Select
    Call CmdClear_Click
    Call cmdSearch_Click
Screen.MousePointer = vbDefault
End Sub
Public Sub SaveGLTemplate()
Const METHOD_NAME = "SaveGLTemplate"
On Error GoTo debugs
 Dim GrpCompnay() As String
 Dim Msg As String
 Dim GrpComp As String
 Dim tmpdepind As String
 Msg = ""
 Set MaintRS = New ADODB.Recordset
 SqlQry = ""
 
 Dim Pos As Integer
 Pos = InStr(cmbGrpCompany, "-")
 
 If Len(CmbdependentTemplateName) = 0 Then
    tmpdepind = "N"
 Else
    tmpdepind = "Y"
 End If
 
 If Pos = 0 Then
     GrpComp = cmbGrpCompany
 Else
     GrpCompnay() = Split(cmbGrpCompany, "-")
     GrpComp = GrpCompnay(1)
 End If
    
    If EdtFlag Then
        'Update Tbl_OPGLMaster
        SqlQry = "Update Tbl_OPGLMaster Set GRPCompnay='" + Trim(GrpComp) + "', GLType='" + Trim(Mid(cmbGLType, 5)) + "',GLTemplate='" + Trim(CmbTemplateName) + "',Paycode='" + Trim(cmbPaycode) + "',LGFormat='" + Trim(cmbLGFormat) + "',VIPInd='" + Trim(cboVIP) + "',DepInd='" + Trim(tmpdepind) + "',DepTemplate='" + Trim(CmbdependentTemplateName) + "',TmpRemarks='" + Trim(txttemplateremarks) + "'  From Tbl_OPGLMaster where PLNCode ='" + Trim(cmbPlnCode) + "'"
        MaintRS.Open SqlQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
        Msg = "Record Updated Successfully.."
    Else
        'Insert Tbl_OPGLMaster
        SqlQry = "insert into Tbl_OPGLMaster (GRPCompnay , PLNCode , GLType, GLTemplate , Paycode , LGFormat,VIPInd,DepInd,DepTemplate,TmpRemarks) Values  ('" + Trim(GrpComp) + "','" + Trim(Mid(cmbPlnCode, 1, 4)) + "','" + Trim(Mid(cmbGLType, 5)) + "','" + Trim(CmbTemplateName) + "','" + Trim(cmbPaycode) + "','" + Trim(cmbLGFormat) + "','" + Trim(Mid(cboVIP, 1, 1)) + "','" + Trim(tmpdepind) + "','" + Trim(CmbdependentTemplateName) + "','" + Trim(txttemplateremarks) + "')"
        MaintRS.Open SqlQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
        Msg = "New Record Inserted Successfully.."
    End If
        'MaintRS.Close
        Set MaintRS = Nothing
        MsgBox (Msg)
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Public Sub SaveGLType()
Const METHOD_NAME = "SaveGLType"
On Error GoTo debugs
Dim Msg As String
Msg = ""
    Set MaintRS = New ADODB.Recordset
    SqlQry = ""
    If EdtFlag Then
        SqlQry = "Update  Tbl_GLType set GLType='" + Trim(txtGLType) + "',GLDesc='" + Trim(txtGLDesc) + "' where GLID='" + Trim(txtGLID) + "'"
        MaintRS.Open SqlQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
        Msg = "Record Updated Successfully.."
    Else
        SqlQry = "Insert into Tbl_GLType (GLType,GLDesc) values ('" + Trim(txtGLType) + "','" + Trim(txtGLDesc) + "')"
        MaintRS.Open SqlQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
        Msg = "New Record Inserted Successfully.."
    End If
        'MaintRS.Close
        Set MaintRS = Nothing
        MsgBox (Msg)
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub SaveTemplate()
Const METHOD_NAME = "SaveTemplate"
On Error GoTo debugs

Dim Msg As String
Msg = ""
    Set MaintRS = New ADODB.Recordset
    SqlQry = ""
    'Select GLTemplateID , GLTemplateName , GLTemplateDesc  from Tbl_GLTemplates
    If EdtFlag Then
        SqlQry = "Update  Tbl_GLTemplates set GLTemplateName='" + Trim(txtTemplate) + "',GLTemplateDesc='" + Trim(txtTemplateDesc) + "', TmpRemarks='" + Trim(txttemplateremarks1) + "'  where GLTemplateID='" + Trim(txtTemplateID) + "'"
        MaintRS.Open SqlQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
        Msg = "Record Updated Successfully.."
    Else
        SqlQry = "Insert into Tbl_GLTemplates (GLTemplateName,GLTemplateDesc,TmpRemarks ) values ('" + Trim(txtTemplate) + "','" + Trim(txtTemplateDesc) + "','" + Trim(txttemplateremarks1) + "')"
        MaintRS.Open SqlQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
        Msg = "New Record Inserted Successfully.."
    End If
        'MaintRS.Close
        Set MaintRS = Nothing
        MsgBox (Msg)
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub SaveDecline()
Const METHOD_NAME = "SaveDecline"
On Error GoTo debugs
Dim Msg As String
Msg = ""
    Set MaintRS = New ADODB.Recordset
    SqlQry = ""
    'Select ID, DeclineReason , Remarks  from Tbl_DeclineReason
    If EdtFlag Then
        SqlQry = "Update  Tbl_DeclineReason set DeclineReason='" + Trim(txtDeclineReason) + "',Remarks='" + Trim(txtDecRemarks) + "' where ID='" + Trim(txtDecId) + "'"
        MaintRS.Open SqlQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
        Msg = "Record Updated Successfully.."
    Else
        SqlQry = "Insert into Tbl_DeclineReason (DeclineReason,Remarks) values ('" + Trim(txtDeclineReason) + "','" + Trim(txtDecRemarks) + "')"
        MaintRS.Open SqlQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
        Msg = "New Record Inserted Successfully.."
    End If
        'MaintRS.Close
        Set MaintRS = Nothing
        MsgBox (Msg)
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub cmdSearch_Click()
    cmdEdit.Enabled = False
 Select Case SSTab1.Tab
            Case 0 'GL Template
                Call SearchGLTemplate
            Case 1 'GL Type
                Call SearchGLType
            Case 2 'Template Name
                Call SearchTemplate
            Case 3 'Decline Reason
               Call SearchDecline
            End Select
End Sub
Private Sub Form_Load()
    Call ActionLogUpdate(12, "frmLGMaintenance")
    EdtFlag = False
    Call BindGLType
    Call BindGroupCompany
    Call BindGLTemplateName
    Call BindGLTemplateName_Dep
    Call BindCLNLGType
    Call BindVIPInd
End Sub
Public Sub BindVIPInd()
    cboVIP.Clear
    cboVIP.AddItem ("N-No")
    cboVIP.AddItem ("Y-Yes")
End Sub

Public Sub SearchGLTemplate()
Const METHOD_NAME = "SearchGLTemplate"
On Error GoTo debugs
    Dim Subqry As String
    Dim GrpCompnay() As String
    Dim GrpComp As String
    Dim Pos As Integer
    Pos = InStr(cmbGrpCompany, "-")
    If Pos = 0 Then
        GrpComp = cmbGrpCompany
    Else
        GrpCompnay() = Split(cmbGrpCompany, "-")
        GrpComp = GrpCompnay(1)
    End If
    
    If Len(Trim(cmbGrpCompany)) = 0 Then
        If Len(Trim(SearchGrpComp)) Then
             'GrpCompnay() = Split(SearchGrpComp, "-")
             Subqry = ""
             Subqry = Subqry + " and GRPCompnay like '%" + Trim(GrpComp) + "%'"
        Else
            MsgBox ("Plese Select Group Company to Search")
            Exit Sub
        End If
    Else
        GrpCompnay() = Split(cmbGrpCompany, "-")
        Subqry = ""
        Subqry = Subqry + " and GRPCompnay like '%" + Trim(GrpCompnay(1)) + "%'"
    End If
    
    If Len(Trim(cmbPaycode)) Then
        Subqry = Subqry + " and Paycode like '" + cmbPaycode + "'"
    End If
    
    If Len(Trim(cmbPlnCode)) Then
        Subqry = Subqry + " and PLNCode like '" + cmbPlnCode + "'"
    End If
    Rcnt = 0
    
    Set MaintRS = New ADODB.Recordset
    lstLGTemplate.ListItems.Clear
    SqlQry = ""
    SqlQry = "Select GRPCompnay , PLNCode , GLType, GLTemplate , Paycode , LGFormat,VIPInd from Tbl_OPGLMaster where 1=1 " + Subqry
    MaintRS.Open SqlQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do Until MaintRS.EOF
        Rcnt = Rcnt + 1
        Set LsRecords = lstLGTemplate.ListItems.Add(, , Rcnt)
        LsRecords.SubItems(1) = IIf(Len(MaintRS!GrpCompnay), MaintRS!GrpCompnay, "")
        LsRecords.SubItems(2) = IIf(Len(MaintRS!PAYCode), MaintRS!PAYCode, "")
        LsRecords.SubItems(3) = IIf(Len(MaintRS!PLNCode), MaintRS!PLNCode, "")
        LsRecords.SubItems(4) = IIf(Len(MaintRS!GLType), MaintRS!GLType, "")
        LsRecords.SubItems(5) = IIf(Len(MaintRS!GLTemplate), MaintRS!GLTemplate, "")
        LsRecords.SubItems(6) = IIf(Len(MaintRS!LGFormat), MaintRS!LGFormat, "")
        LsRecords.SubItems(7) = IIf(Len(MaintRS!VipInd), MaintRS!VipInd, "")
    MaintRS.MoveNext
    Loop
    MaintRS.Close
    Set MaintRS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub SearchGLType()
Const METHOD_NAME = "SearchGLType"
On Error GoTo debugs
    Dim Subqry As String
    If Len(Trim(txtGLType)) Then
        Subqry = Subqry + " and GLType like '%" + txtGLType + "%'"
    End If
    Rcnt = 0
    Set MaintRS = New ADODB.Recordset
    SqlQry = ""
    SqlQry = "Select GLID , GLType , GLDesc  from Tbl_GLType  where 1=1 " + Subqry
    MaintRS.Open SqlQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
    lstGLType.ListItems.Clear
    Do Until MaintRS.EOF
        Rcnt = Rcnt + 1
        Set LsRecords = lstGLType.ListItems.Add(, , MaintRS!GLID)
        LsRecords.SubItems(1) = IIf(Len(MaintRS!GLType), MaintRS!GLType, "")
        LsRecords.SubItems(2) = IIf(Len(MaintRS!GLDesc), MaintRS!GLDesc, "")
    MaintRS.MoveNext
    Loop
    MaintRS.Close
    Set MaintRS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub SearchTemplate()
Const METHOD_NAME = "SearchTemplate"
On Error GoTo debugs
    Dim Subqry As String
    If Len(Trim(txtTemplate)) Then
        Subqry = Subqry + " and GLTemplateName like '%" + txtTemplate + "%'"
    End If
    Rcnt = 0
    Set MaintRS = New ADODB.Recordset
    SqlQry = ""
    SqlQry = "Select GLTemplateID , GLTemplateName , GLTemplateDesc,TmpRemarks  from Tbl_GLTemplates   where 1=1 " + Subqry
    MaintRS.Open SqlQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
    lstGLTemplate.ListItems.Clear
    Do Until MaintRS.EOF
        Rcnt = Rcnt + 1
        Set LsRecords = lstGLTemplate.ListItems.Add(, , MaintRS!GLTemplateID)
        LsRecords.SubItems(1) = IIf(Len(MaintRS!GLTemplateName), MaintRS!GLTemplateName, "")
        LsRecords.SubItems(2) = IIf(Len(MaintRS!GLTemplateDesc), MaintRS!GLTemplateDesc, "")
        LsRecords.SubItems(3) = IIf(Len(MaintRS!TmpRemarks), MaintRS!TmpRemarks, "")
    MaintRS.MoveNext
    Loop
    MaintRS.Close
    Set MaintRS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub


Public Sub SearchDecline()
Const METHOD_NAME = "SearchDecline"
On Error GoTo debugs
    Dim Subqry As String
    If Len(Trim(txtDeclineReason)) Then
        Subqry = Subqry + " and DeclineReason like '%" + txtDeclineReason + "%'"
    End If
    Rcnt = 0
    Set MaintRS = New ADODB.Recordset
    SqlQry = ""
    SqlQry = "Select ID, DeclineReason , Remarks  from Tbl_DeclineReason   where 1=1 " + Subqry
    MaintRS.Open SqlQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
    lstDeclineReason.ListItems.Clear
    Do Until MaintRS.EOF
        Rcnt = Rcnt + 1
        Set LsRecords = lstDeclineReason.ListItems.Add(, , MaintRS!ID)
        LsRecords.SubItems(1) = IIf(Len(MaintRS!DeclineReason), MaintRS!DeclineReason, "")
        LsRecords.SubItems(2) = IIf(Len(MaintRS!Remarks), MaintRS!Remarks, "")
    MaintRS.MoveNext
    Loop
    MaintRS.Close
    Set MaintRS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Public Sub BindPlanCode()
Const METHOD_NAME = "BindPlanCode"
On Error GoTo debugs

Set MaintRS = New ADODB.Recordset
SqlQry = ""
Dim GrpCode As String
GrpCode = Mid(cmbGrpCompany, 1, 4)
SqlQry = "select PLNCode, PLNDesc   from PLN where GRPCompany='" + GrpCode + "'  "
MaintRS.Open SqlQry, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
cmbPlnCode.Clear
Do Until MaintRS.EOF
    cmbPlnCode.AddItem MaintRS!PLNCode & " - " & MaintRS!PlnDesc
MaintRS.MoveNext
Loop
MaintRS.Close
Set MaintRS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub BindPayCode()
Const METHOD_NAME = "BindPayCode"
On Error GoTo debugs
Set MaintRS = New ADODB.Recordset
SqlQry = ""
Dim GrpCode As String
GrpCode = Mid(cmbGrpCompany, 1, 4)
SqlQry = "select  Distinct Paycode  from PLN where GRPCompany='" + GrpCode + "'  "
MaintRS.Open SqlQry, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
cmbPaycode.Clear
Do Until MaintRS.EOF
    cmbPaycode.AddItem MaintRS!PAYCode
MaintRS.MoveNext
Loop
MaintRS.Close
Set MaintRS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub


Public Sub BindGroupCompany()
Const METHOD_NAME = "BindGroupCompany"
On Error GoTo debugs
Set MaintRS = New ADODB.Recordset
SqlQry = ""
'SqlQry = "Select Distinct GRPCoCode , GRPCoName , GRPCoPaycode  from GRPCompany Union " & _
'         " Select Distinct  CoGRPCode as GRPCoCode , CoGRPName as GRPCoName, CoGRPPaycode as GRPCoPaycode    from HISDB..CompanyGRP "

SqlQry = "Select Distinct GRPCoCode , GRPCoName , GRPCoPaycode  from GRPCompany"

MaintRS.Open SqlQry, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
cmbGrpCompany.Clear
Do Until MaintRS.EOF
    cmbGrpCompany.AddItem MaintRS!GRPCoCode & " - " & MaintRS!GrpCoName
MaintRS.MoveNext
Loop
MaintRS.Close
Set MaintRS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub BindGLType()
Const METHOD_NAME = "BindGLType"
On Error GoTo debugs

Set MaintRS = New ADODB.Recordset
SqlQry = ""
cmbGLType.Clear
SqlQry = "Select GLID,GLType,GLDesc from Tbl_GLType"
MaintRS.Open SqlQry, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
Do Until MaintRS.EOF
    cmbGLType.AddItem MaintRS!GLID & " - " & MaintRS!GLType
MaintRS.MoveNext
Loop
MaintRS.Close
Set MaintRS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub BindCLNLGType()
Const METHOD_NAME = "BindCLNLGType"
On Error GoTo debugs

Set MaintRS = New ADODB.Recordset
SqlQry = ""
cmbLGFormat.Clear
SqlQry = "select CLNGLFormat  from HISDB..CLN  Group by CLNGLFormat"
MaintRS.Open SqlQry, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
Do Until MaintRS.EOF
   If Len(MaintRS!CLNGLFormat) Then
        cmbLGFormat.AddItem MaintRS!CLNGLFormat
   End If
    
MaintRS.MoveNext
Loop
MaintRS.Close
Set MaintRS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Public Sub BindGLTemplateName()
Const METHOD_NAME = "BindGLTemplateName"
On Error GoTo debugs
Set MaintRS = New ADODB.Recordset
SqlQry = ""
CmbTemplateName.Clear
SqlQry = "Select GLTemplateID,GLTemplateDesc,GLTemplateName,TmpRemarks from Tbl_GLTemplates"
MaintRS.Open SqlQry, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
Do Until MaintRS.EOF
    CmbTemplateName.AddItem MaintRS!GLTemplateName
MaintRS.MoveNext
Loop

MaintRS.Close
Set MaintRS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub CmbTemplateName_Click()
'Const METHOD_NAME = "CmbTemplateName_Click"
'On Error GoTo debugs
Set templateremark = New ADODB.Recordset
SqlQry = ""
txttemplateremarks = ""
SqlQry = "Select GLTemplateID,GLTemplateDesc,GLTemplateName,TmpRemarks from Tbl_GLTemplates where GLTemplateName='" + Trim(CmbTemplateName) + "'"
templateremark.Open SqlQry, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
txttemplateremarks = IIf(Len(templateremark!TmpRemarks), templateremark!TmpRemarks, "")
templateremark.Close
Set templateremark = Nothing
'If Err.Description <> "" Then
 '   Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
'End If
End Sub
Public Sub BindGLTemplateName_Dep()
Const METHOD_NAME = "BindGLTemplateName_Dep"
On Error GoTo debugs
Set MaintRS = New ADODB.Recordset
SqlQry = ""
CmbdependentTemplateName.Clear
SqlQry = "Select GLTemplateID,GLTemplateDesc,GLTemplateName from Tbl_GLTemplates"
MaintRS.Open SqlQry, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
Do Until MaintRS.EOF
    CmbdependentTemplateName.AddItem MaintRS!GLTemplateName
MaintRS.MoveNext
Loop
MaintRS.Close
Set MaintRS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstDeclineReason_Click()
Dim lstRec As ListItem
If lstDeclineReason.ListItems.Count > 0 Then
    Set lstRec = lstDeclineReason.SelectedItem
    txtDecId = lstDeclineReason.SelectedItem
    txtDecId.Enabled = False
    txtDeclineReason = lstRec.SubItems(1)
    txtDecRemarks = lstRec.SubItems(2)
    cmdEdit.Enabled = True
End If
End Sub

Private Sub lstGLTemplate_Click()
Dim lstRec As ListItem
If lstGLTemplate.ListItems.Count > 0 Then
   Set lstRec = lstGLTemplate.SelectedItem
    txtTemplateID = lstGLTemplate.SelectedItem
    txtTemplateID.Enabled = False
    txtTemplate = lstRec.SubItems(1)
    txtTemplateDesc = lstRec.SubItems(2)
    cmdEdit.Enabled = True
End If
End Sub

Private Sub lstGLType_Click()
Dim lstRec As ListItem
If lstGLType.ListItems.Count > 0 Then
   Set lstRec = lstGLType.SelectedItem
    txtGLID = lstGLType.SelectedItem
    txtGLID.Enabled = False
    txtGLType = lstRec.SubItems(1)
    txtGLDesc = lstRec.SubItems(2)
    cmdEdit.Enabled = True
End If
End Sub

Private Sub lstLGTemplate_Click()
Dim lstRec As ListItem
If lstLGTemplate.ListItems.Count > 0 Then
   Set lstRec = lstLGTemplate.SelectedItem
    cmbGrpCompany = lstRec.SubItems(1)
    cmbPaycode = lstRec.SubItems(2)
    cmbGLType = lstRec.SubItems(4)
    cmbPlnCode = lstRec.SubItems(3)
    cmbLGFormat = lstRec.SubItems(6)
    cboVIP = lstRec.SubItems(7)
    CmbTemplateName = lstRec.SubItems(5)
    cmdEdit.Enabled = True
End If
End Sub

