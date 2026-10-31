VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaintenanceEntity 
   Caption         =   "Entity Maintenance"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6195
      Left            =   120
      TabIndex        =   72
      Top             =   600
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   10927
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Clinic Group "
      TabPicture(0)   =   "frmMaintenanceEntity.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label46"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstGrpClinicList"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Doctor"
      TabPicture(1)   =   "frmMaintenanceEntity.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "lstDoctorList"
      Tab(1).Control(2)=   "Label34"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Clinic Info"
      TabPicture(2)   =   "frmMaintenanceEntity.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "lstClinicInfoList"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Group Company Clinic"
      TabPicture(3)   =   "frmMaintenanceEntity.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(1)=   "lstGrpComCLN"
      Tab(3).Control(2)=   "Label26"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Clinic Access"
      TabPicture(4)   =   "frmMaintenanceEntity.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label27"
      Tab(4).Control(1)=   "Frame1"
      Tab(4).Control(2)=   "lstClinicAccess"
      Tab(4).ControlCount=   3
      Begin MSComctlLib.ListView lstClinicAccess 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   157
         Top             =   2520
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clinic Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clinic Name"
            Object.Width           =   5469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Effective From"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Effective To"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Created By"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Created Date"
            Object.Width           =   3881
         EndProperty
      End
      Begin VB.Frame Frame1 
         Height          =   1575
         Left            =   -74760
         TabIndex        =   153
         Top             =   480
         Width           =   11055
         Begin VB.CommandButton cmdCLNLookup 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   320
            Left            =   3600
            TabIndex        =   160
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtCAClnCode 
            Enabled         =   0   'False
            Height          =   320
            Left            =   1320
            TabIndex        =   159
            Top             =   240
            Width           =   2175
         End
         Begin MSMask.MaskEdBox txtEffDateFrom 
            Height          =   285
            Left            =   1320
            TabIndex        =   162
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtEffDateTo 
            Height          =   285
            Left            =   1320
            TabIndex        =   163
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblClinicName 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4080
            TabIndex        =   161
            Top             =   270
            Width           =   45
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Effective To:*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   156
            Top             =   990
            Width           =   1020
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Effective From:*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   155
            Top             =   645
            UseMnemonic     =   0   'False
            Width           =   1200
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Clinic Code:*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   154
            Top             =   255
            Width           =   930
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1080
         Left            =   -74760
         TabIndex        =   129
         Top             =   720
         Width           =   11055
         Begin VB.ComboBox CboGrpComName 
            Height          =   315
            Left            =   1920
            Sorted          =   -1  'True
            TabIndex        =   67
            Top             =   240
            Width           =   7815
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   9840
            TabIndex        =   68
            Top             =   240
            Width           =   255
         End
         Begin VB.ComboBox CboCLNName 
            Height          =   315
            Left            =   1920
            Sorted          =   -1  'True
            TabIndex        =   69
            Top             =   525
            Width           =   7815
         End
         Begin VB.Label Label13 
            Caption         =   "Clinic Name *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   131
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label25 
            Caption         =   "Group Company *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   255
            TabIndex        =   130
            Top             =   240
            Width           =   1800
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4275
         Left            =   -74880
         TabIndex        =   100
         Top             =   660
         Width           =   11235
         Begin VB.TextBox txtPymtInfoClnName 
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
            Left            =   1680
            MaxLength       =   200
            TabIndex        =   33
            Top             =   3070
            Width           =   3615
         End
         Begin VB.TextBox txtInfoClnName 
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
            Left            =   4020
            MaxLength       =   50
            TabIndex        =   2
            Top             =   240
            Width           =   7095
         End
         Begin VB.TextBox txtInfoClnMIX 
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
            Left            =   6780
            MaxLength       =   50
            TabIndex        =   17
            Top             =   500
            Width           =   4335
         End
         Begin VB.ComboBox cboInfoClnType 
            Height          =   315
            Left            =   6780
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   735
            Width           =   1695
         End
         Begin MSMask.MaskEdBox txtInfoClnReceivedDate 
            Height          =   285
            Left            =   6780
            TabIndex        =   20
            Top             =   1020
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboInfoClnPanel 
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
            Left            =   6780
            Sorted          =   -1  'True
            TabIndex        =   22
            Top             =   1260
            Width           =   1695
         End
         Begin VB.TextBox txtInfoCLNCode 
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
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   1
            Top             =   240
            Width           =   930
         End
         Begin VB.ComboBox CboInfoClnGrp 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   500
            Width           =   3615
         End
         Begin VB.TextBox txtInfoClnAdd1 
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
            Left            =   1680
            MaxLength       =   200
            TabIndex        =   4
            Top             =   780
            Width           =   3615
         End
         Begin VB.TextBox txtInfoClnAdd2 
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
            Left            =   1680
            MaxLength       =   200
            TabIndex        =   5
            Top             =   1040
            Width           =   3615
         End
         Begin VB.TextBox txtInfoClnAdd3 
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
            Left            =   1680
            MaxLength       =   200
            TabIndex        =   6
            Top             =   1280
            Width           =   3615
         End
         Begin VB.TextBox txtInfoClnPostCode 
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
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   7
            Top             =   1530
            Width           =   1455
         End
         Begin VB.ComboBox cboInfoClnState 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   1780
            Width           =   1455
         End
         Begin VB.TextBox txtInfoClnContact 
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
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   11
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtInfoClnDocName 
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
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   13
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox txtInfoClnAccPerson 
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
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   15
            Top             =   2520
            Width           =   1455
         End
         Begin VB.ComboBox cboInfoClnCity 
            Height          =   315
            Left            =   4080
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   1530
            Width           =   1215
         End
         Begin VB.TextBox txtInfoClnTelNo 
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
            MaxLength       =   12
            TabIndex        =   12
            Top             =   1820
            Width           =   1215
         End
         Begin VB.TextBox txtInfoClnFaxNo 
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
            MaxLength       =   12
            TabIndex        =   14
            Text            =   "  "
            Top             =   2070
            Width           =   1215
         End
         Begin VB.TextBox txtInfoClnNoDoc 
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
            MaxLength       =   3
            TabIndex        =   16
            Top             =   2320
            Width           =   1215
         End
         Begin VB.ComboBox cboInfoCln24hrs 
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
            Left            =   6780
            Sorted          =   -1  'True
            TabIndex        =   24
            Top             =   1520
            Width           =   1695
         End
         Begin VB.TextBox txtOperatinghr 
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
            Left            =   6780
            TabIndex        =   26
            Top             =   1800
            Width           =   1695
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
            Left            =   6780
            Sorted          =   -1  'True
            TabIndex        =   28
            Top             =   2090
            Width           =   1695
         End
         Begin VB.TextBox txtPymtInfoClnAdd1 
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
            Left            =   1680
            MaxLength       =   200
            TabIndex        =   37
            Top             =   3330
            Width           =   3615
         End
         Begin VB.TextBox txtPymtInfoClnAdd2 
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
            Left            =   1680
            MaxLength       =   200
            TabIndex        =   38
            Top             =   3590
            Width           =   3615
         End
         Begin VB.TextBox txtPymtInfoClnPostCode 
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
            Left            =   7920
            MaxLength       =   5
            TabIndex        =   34
            Top             =   4200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.ComboBox cboPymtInfoClnCity 
            Height          =   315
            Left            =   7560
            Sorted          =   -1  'True
            TabIndex        =   35
            Top             =   4200
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox cboPymtInfoClnState 
            Height          =   315
            Left            =   7440
            Sorted          =   -1  'True
            TabIndex        =   36
            Top             =   4200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtPymtInfoClnAdd3 
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
            Left            =   1680
            MaxLength       =   200
            TabIndex        =   39
            Top             =   3840
            Width           =   3615
         End
         Begin VB.ComboBox cboInfoClnBNFType 
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
            Left            =   9420
            Sorted          =   -1  'True
            TabIndex        =   19
            Top             =   735
            Width           =   1695
         End
         Begin VB.ComboBox cboInfoClnXray 
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
            Left            =   9420
            Sorted          =   -1  'True
            TabIndex        =   21
            Top             =   1020
            Width           =   1695
         End
         Begin VB.TextBox txtInfoClnFormema 
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
            Left            =   9420
            MaxLength       =   15
            TabIndex        =   23
            Top             =   1310
            Width           =   1695
         End
         Begin VB.TextBox txtInfoClnBranch 
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
            Left            =   9420
            MaxLength       =   50
            TabIndex        =   25
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox txtInfoClnBankAcc 
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
            Left            =   9420
            MaxLength       =   25
            TabIndex        =   27
            Top             =   1820
            Width           =   1695
         End
         Begin VB.TextBox txtFFSPerc 
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
            Left            =   9420
            MaxLength       =   50
            TabIndex        =   149
            Top             =   2070
            Width           =   1695
         End
         Begin VB.TextBox txtInfoClnBank 
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
            Left            =   6780
            MaxLength       =   50
            TabIndex        =   30
            Top             =   2340
            Width           =   4335
         End
         Begin VB.TextBox txtInfoClnRemark 
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
            Left            =   6780
            MaxLength       =   1500
            TabIndex        =   31
            Top             =   2600
            Width           =   4335
         End
         Begin VB.TextBox txtInfoClnHome 
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
            Left            =   6780
            MaxLength       =   50
            TabIndex        =   32
            Top             =   2850
            Width           =   4335
         End
         Begin VB.ComboBox cboOnlineStatus 
            Height          =   315
            Left            =   6780
            Sorted          =   -1  'True
            TabIndex        =   145
            Top             =   3100
            Width           =   1815
         End
         Begin VB.ComboBox cboGLFormat 
            Height          =   315
            Left            =   9480
            Sorted          =   -1  'True
            TabIndex        =   147
            Top             =   3110
            Width           =   1680
         End
         Begin VB.TextBox TxtLocation 
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
            Left            =   6780
            MaxLength       =   100
            TabIndex        =   29
            Top             =   3390
            Width           =   1695
         End
         Begin VB.TextBox txtRegNo 
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
            MaxLength       =   50
            TabIndex        =   151
            Top             =   3370
            Width           =   1640
         End
         Begin VB.TextBox txtInfoClnEmail 
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
            Left            =   6780
            MaxLength       =   1000
            TabIndex        =   10
            Top             =   3630
            Width           =   4335
         End
         Begin VB.Label Label71 
            Caption         =   "Reg. No"
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
            Index           =   2
            Left            =   8640
            TabIndex        =   152
            Top             =   3390
            Width           =   1215
         End
         Begin VB.Label Label71 
            Caption         =   "FFS %"
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
            Index           =   1
            Left            =   8520
            TabIndex        =   150
            Top             =   2070
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "GL Format"
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
            Index           =   1
            Left            =   8640
            TabIndex        =   148
            Top             =   3150
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Online Status"
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
            Index           =   0
            Left            =   5400
            TabIndex        =   146
            Top             =   3150
            Width           =   975
         End
         Begin VB.Label Label63 
            Caption         =   "State"
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
            Index           =   1
            Left            =   4200
            TabIndex        =   142
            Top             =   3840
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label60 
            Caption         =   "City"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   4320
            TabIndex        =   141
            Top             =   3840
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label Label61 
            Caption         =   "Post Code"
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
            Index           =   1
            Left            =   4080
            TabIndex        =   140
            Top             =   3720
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Label Label2 
            Caption         =   "Clinic Address"
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
            Left            =   240
            TabIndex        =   139
            Top             =   3360
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Payee"
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
            Left            =   240
            TabIndex        =   138
            Top             =   3060
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Payment Details"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   137
            Top             =   2820
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Operating Hours *"
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
            Left            =   5400
            TabIndex        =   136
            Top             =   1800
            Width           =   1695
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
            Left            =   5400
            TabIndex        =   135
            Top             =   2100
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Location *"
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
            Left            =   5400
            TabIndex        =   134
            Top             =   3400
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "24 Hours *"
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
            Index           =   14
            Left            =   5400
            TabIndex        =   127
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label71 
            Caption         =   "Branch"
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
            Index           =   0
            Left            =   8520
            TabIndex        =   126
            Top             =   1590
            Width           =   1215
         End
         Begin VB.Label Label70 
            Caption         =   "Bank A/C No"
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
            Left            =   8520
            TabIndex        =   125
            Top             =   1860
            Width           =   975
         End
         Begin VB.Label Label69 
            Caption         =   "Name Of Bank"
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
            Left            =   5400
            TabIndex        =   124
            Top             =   2360
            Width           =   1215
         End
         Begin VB.Label Label68 
            Caption         =   "Remarks"
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
            Left            =   5400
            TabIndex        =   123
            Top             =   2620
            Width           =   855
         End
         Begin VB.Label Label66 
            Caption         =   "Received Date *"
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
            Left            =   5400
            TabIndex        =   122
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label65 
            Caption         =   "Clinic Initial *"
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
            Left            =   5400
            TabIndex        =   121
            Top             =   555
            Width           =   1215
         End
         Begin VB.Label Label64 
            Caption         =   "For. Ref."
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
            Left            =   8520
            TabIndex        =   120
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label62 
            Caption         =   "X-Ray"
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
            Left            =   8520
            TabIndex        =   119
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Panel *"
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
            Index           =   13
            Left            =   5400
            TabIndex        =   118
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label59 
            Caption         =   "No Of Doc"
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
            Left            =   3240
            TabIndex        =   117
            Top             =   2320
            Width           =   975
         End
         Begin VB.Label Label48 
            Caption         =   "Doctor Name"
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
            Left            =   240
            TabIndex        =   116
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label49 
            Caption         =   "Contact Person"
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
            Left            =   240
            TabIndex        =   115
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label Label63 
            Caption         =   "State *"
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
            Index           =   0
            Left            =   240
            TabIndex        =   114
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Clinic Type*"
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
            Left            =   5400
            TabIndex        =   113
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label36 
            Caption         =   "Bnf. Type*"
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
            Left            =   8520
            TabIndex        =   112
            Top             =   795
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Clinic Grp. Code"
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
            Index           =   11
            Left            =   240
            TabIndex        =   111
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Clinic Code        "
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
            Index           =   6
            Left            =   240
            TabIndex        =   110
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Clinic Address       *"
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
            Index           =   6
            Left            =   240
            TabIndex        =   109
            Top             =   800
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Clinic Name *"
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
            Index           =   6
            Left            =   2760
            TabIndex        =   108
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Fax No.      "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3240
            TabIndex        =   107
            Top             =   2080
            Width           =   1005
         End
         Begin VB.Label Label22 
            Caption         =   "Tel No. "
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
            Left            =   3240
            TabIndex        =   106
            Top             =   1850
            Width           =   1230
         End
         Begin VB.Label Label23 
            Caption         =   "Email"
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
            Left            =   5400
            TabIndex        =   105
            Top             =   3650
            Width           =   615
         End
         Begin VB.Label Label24 
            Caption         =   "Homepage"
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
            Left            =   5400
            TabIndex        =   104
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label61 
            Caption         =   "Post Code *"
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
            Index           =   0
            Left            =   240
            TabIndex        =   103
            Top             =   1560
            Width           =   1230
         End
         Begin VB.Label Label60 
            Caption         =   "City *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   3240
            TabIndex        =   102
            Top             =   1560
            Width           =   525
         End
         Begin VB.Label Label67 
            Caption         =   "Accounts Person"
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
            Left            =   240
            TabIndex        =   101
            Top             =   2520
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1080
         Left            =   -74760
         TabIndex        =   95
         Top             =   720
         Width           =   11055
         Begin VB.TextBox txtDocCode 
            Height          =   285
            Left            =   1920
            MaxLength       =   15
            TabIndex        =   65
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtDocName 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   66
            Top             =   600
            Width           =   5775
         End
         Begin VB.Label Label29 
            Caption         =   "Doctor Code*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   255
            TabIndex        =   97
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label Label14 
            Caption         =   "Doctor Name *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   96
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2415
         Left            =   240
         TabIndex        =   77
         Top             =   720
         Width           =   11055
         Begin VB.TextBox txtClnGrpCode 
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
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   0
            Top             =   255
            Width           =   930
         End
         Begin VB.TextBox txtClnGrpName 
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
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   47
            Top             =   500
            Width           =   3465
         End
         Begin VB.TextBox txtClnName 
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
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   40
            Top             =   500
            Width           =   3975
         End
         Begin VB.TextBox txtClnAdd1 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   41
            Top             =   755
            Width           =   3975
         End
         Begin VB.TextBox txtClnAdd2 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   42
            Top             =   1010
            Width           =   3975
         End
         Begin VB.TextBox txtClnAdd3 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   43
            Top             =   1265
            Width           =   3975
         End
         Begin VB.TextBox txtClnContact 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   44
            Top             =   1520
            Width           =   3975
         End
         Begin VB.TextBox txtClnAdmContact 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   45
            Top             =   1770
            Width           =   3975
         End
         Begin VB.TextBox txtClnCEO 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   46
            Top             =   2030
            Width           =   3975
         End
         Begin VB.TextBox txtClnPhone 
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
            Left            =   7440
            MaxLength       =   15
            TabIndex        =   48
            Top             =   755
            Width           =   1275
         End
         Begin VB.TextBox txtClnFax 
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
            Left            =   9600
            MaxLength       =   15
            TabIndex        =   49
            Top             =   755
            Width           =   1290
         End
         Begin VB.TextBox txtClnEmail 
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
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   50
            Top             =   1010
            Width           =   3435
         End
         Begin VB.TextBox txtClnHomepage 
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
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   51
            Top             =   1265
            Width           =   3435
         End
         Begin MSMask.MaskEdBox txtClnEffDate 
            Height          =   285
            Left            =   7440
            TabIndex        =   52
            Top             =   1520
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboClnTax 
            Height          =   315
            Left            =   7440
            Sorted          =   -1  'True
            TabIndex        =   54
            Top             =   1775
            Width           =   1335
         End
         Begin VB.ComboBox cboClnConversion 
            Height          =   315
            Left            =   9795
            Sorted          =   -1  'True
            TabIndex        =   53
            Top             =   1520
            Width           =   1095
         End
         Begin VB.ComboBox cboClnPrePost 
            Height          =   315
            Left            =   9795
            Sorted          =   -1  'True
            TabIndex        =   55
            Top             =   1790
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "R&&B Tax"
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
            Left            =   6100
            TabIndex        =   92
            Top             =   1790
            Width           =   810
         End
         Begin VB.Label Label2 
            Caption         =   "Pre&&Post"
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
            Index           =   5
            Left            =   8880
            TabIndex        =   91
            Top             =   1810
            Width           =   810
         End
         Begin VB.Label Label2 
            Caption         =   "Conversion"
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
            Left            =   8880
            TabIndex        =   90
            Top             =   1550
            Width           =   930
         End
         Begin VB.Label Label2 
            Caption         =   "Eff. Date  *"
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
            Left            =   6100
            TabIndex        =   89
            Top             =   1550
            Width           =   1050
         End
         Begin VB.Label Label16 
            Caption         =   "Tel No"
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
            Left            =   6100
            TabIndex        =   88
            Top             =   800
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Fax No"
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
            Left            =   8880
            TabIndex        =   87
            Top             =   800
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Contact Person"
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
            Left            =   120
            TabIndex        =   86
            Top             =   1575
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "Homepage"
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
            Left            =   6100
            TabIndex        =   85
            Top             =   1295
            Width           =   870
         End
         Begin VB.Label Label15 
            Caption         =   "Email"
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
            Left            =   6100
            TabIndex        =   84
            Top             =   1040
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Cln Grp. Name *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   6100
            TabIndex        =   83
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Clinic Group Address   *"
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
            Left            =   120
            TabIndex        =   82
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Clinic Group Code   *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   81
            Top             =   255
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "CEO Name"
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
            Left            =   120
            TabIndex        =   80
            Top             =   2085
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "Initial Group Name  *"
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
            TabIndex        =   79
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label20 
            Caption         =   "Adm. Contact Person"
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
            Left            =   120
            TabIndex        =   78
            Top             =   1815
            Width           =   1695
         End
      End
      Begin MSComctlLib.ListView lstGrpClinicList 
         Height          =   2600
         Left            =   240
         TabIndex        =   93
         Top             =   3480
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   4577
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
            Size            =   8.25
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
      Begin MSComctlLib.ListView lstDoctorList 
         Height          =   3900
         Left            =   -74760
         TabIndex        =   98
         Top             =   2160
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   6879
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Doctor Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Doctor Name"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstClinicInfoList 
         Height          =   1150
         Left            =   -74880
         TabIndex        =   128
         Top             =   4980
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   38
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clinic Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clinic Name"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clinic Group Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Clinic Initial"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Clinic Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Benefit Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Address 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Address 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Address 3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Post Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "City"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "State"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Received Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Panel"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "X-Ray"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Formema Ref"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "No Of Doctor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Tel. No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Fax No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Email"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Homepage"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Bank Account"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Bank"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Branch"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Account Person"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "24 Hours"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "Remarks"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "FFS Status"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "Operating Hour"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "Pay. Clinic Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "Pay. Add1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "Pay. Add2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "Pay. Add3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "Pay. Post Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   "Pay. City"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "Pay. State"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "Online Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "Registration No"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstGrpComCLN 
         Height          =   3900
         Left            =   -74760
         TabIndex        =   132
         Top             =   2160
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   6879
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Group Company Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Group Company Name"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clinic Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Clinic Name"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "PostCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "State"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "City"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Location"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "24 Hour"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Phone No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Fax No"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "CLINIC ACCESS FOR BACKDATE ENTRY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74760
         TabIndex        =   158
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "GROUP COMPANY CLINIC LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74760
         TabIndex        =   133
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "DOCTOR LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74760
         TabIndex        =   99
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "CLINIC GROUP LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   240
         TabIndex        =   94
         Top             =   3240
         Width           =   1275
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      TabIndex        =   70
      Top             =   0
      Width           =   11805
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Entity Maintenance - Clinic"
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
         Left            =   120
         TabIndex        =   71
         Top             =   120
         Width           =   3525
      End
   End
   Begin VB.Frame Frame9 
      Height          =   615
      Left            =   120
      TabIndex        =   76
      Top             =   6800
      Width           =   11595
      Begin VB.CommandButton cmdBack 
         Cancel          =   -1  'True
         Caption         =   "&Back"
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
         Left            =   6000
         TabIndex        =   144
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
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
         Left            =   5160
         TabIndex        =   143
         Top             =   180
         Width           =   840
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
         Height          =   315
         Left            =   4320
         TabIndex        =   61
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
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
         Left            =   8880
         TabIndex        =   62
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
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
         Left            =   3480
         TabIndex        =   60
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdCancel 
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
         Height          =   315
         Left            =   10560
         TabIndex        =   64
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
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
         Height          =   315
         Left            =   9720
         TabIndex        =   63
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
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
         Height          =   315
         Left            =   1800
         TabIndex        =   58
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
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
         Height          =   315
         Left            =   960
         TabIndex        =   57
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Left            =   120
         TabIndex        =   56
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
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
         Left            =   2640
         TabIndex        =   59
         Top             =   180
         Width           =   840
      End
   End
   Begin VB.Label Label33 
      Caption         =   "Bold - Searchable        * - Mandatory Fields"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   75
      Top             =   7440
      Width           =   4095
   End
   Begin VB.Label Label6 
      Caption         =   "TOTAL RECORDS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5160
      TabIndex        =   74
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label txtTotalRecord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   73
      Top             =   7440
      Width           =   1815
   End
End
Attribute VB_Name = "frmMaintenanceEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Public GRPClnCriteria As String
Public GRPCoCriteria As String
Public DOCCriteria As String
Public CLNCriteria As String
Public GRPCOMCLNCriteria As String
Private m_blnDirection(1 To 6) As Boolean

Private Sub cboClnConversion_Change()
    If InStr(cboClnConversion.Text, "null") Then
       cboClnConversion.Enabled = False
    End If
End Sub

Private Sub cboClnConversion_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboClnConversion.Text, cboClnConversion.SelStart) + Chr(KeyAscii) + Mid(cboClnConversion.Text, cboClnConversion.SelStart + cboClnConversion.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboClnConversion.ListCount - 1
 
        If NewText = LCase(Left(cboClnConversion.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboClnConversion.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboClnConversion.Text = ValidValue
               cboClnConversion.SelStart = 0
               cboClnConversion.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboClnPrePost_Change()
    If InStr(cboClnPrePost.Text, "null") Then
       cboClnPrePost.Enabled = False
    End If
End Sub

Private Sub cboClnPrePost_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboClnPrePost.Text, cboClnPrePost.SelStart) + Chr(KeyAscii) + Mid(cboClnPrePost.Text, cboClnPrePost.SelStart + cboClnPrePost.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboClnPrePost.ListCount - 1
 
        If NewText = LCase(Left(cboClnPrePost.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboClnPrePost.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboClnPrePost.Text = ValidValue
               cboClnPrePost.SelStart = 0
               cboClnPrePost.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboClnTax_Change()
    If InStr(cboClnTax.Text, "null") Then
       cboClnTax.Enabled = False
    End If
End Sub

Private Sub cboClnTax_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboClnTax.Text, cboClnTax.SelStart) + Chr(KeyAscii) + Mid(cboClnTax.Text, cboClnTax.SelStart + cboClnTax.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboClnTax.ListCount - 1
 
        If NewText = LCase(Left(cboClnTax.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboClnTax.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboClnTax.Text = ValidValue
               cboClnTax.SelStart = 0
               cboClnTax.SelLength = Len(ValidValue)
             End If
        End If
End Sub

'Private Sub cboCoConversion_Change()
    'If InStr(cboCoConversion.Text, "null") Then
       'cboCoConversion.Enabled = False
    'End If
'End Sub

'Private Sub cboCoConversion_KeyPress(KeyAscii As Integer)
'Dim NewText As String
'Dim ValidCount As Integer
'Dim ValidValue As String
'Dim i As Integer
    'Clipboard.Clear
    'If KeyAscii >= 32 And KeyAscii <> 127 Then
    'NewText = LCase(Left(cboCoConversion.Text, cboCoConversion.SelStart) + Chr(KeyAscii) + Mid(cboCoConversion.Text, cboCoConversion.SelStart + cboCoConversion.SelLength + 1))

    'ValidCount = 0
    'ValidValue = ""

    'For i = 0 To cboCoConversion.ListCount - 1
 
        'If NewText = LCase(Left(cboCoConversion.List(i), Len(NewText))) Then
                'ValidCount = ValidCount + 1
                'ValidValue = cboCoConversion.List(i)
        'End If
        'Next
        'If ValidCount <= 1 Then KeyAscii = 0
             'If ValidCount = 1 Then
               'cboCoConversion.Text = ValidValue
               'cboCoConversion.SelStart = 0
               'cboCoConversion.SelLength = Len(ValidValue)
             'End If
        'End If
'End Sub

'Private Sub CboCoGrpPayor_Change()
    'If InStr(CboCoGrpPayor.Text, "null") Then
       'CboCoGrpPayor.Enabled = False
    'End If
'End Sub

'Private Sub CboCoGrpPayor_KeyPress(KeyAscii As Integer)
'Dim NewText As String
'Dim ValidCount As Integer
'Dim ValidValue As String
'Dim i As Integer
    'Clipboard.Clear
    'If KeyAscii >= 32 And KeyAscii <> 127 Then
    'NewText = LCase(Left(CboCoGrpPayor.Text, CboCoGrpPayor.SelStart) + Chr(KeyAscii) + Mid(CboCoGrpPayor.Text, CboCoGrpPayor.SelStart + CboCoGrpPayor.SelLength + 1))

    'ValidCount = 0
    'ValidValue = ""

    'For i = 0 To CboCoGrpPayor.ListCount - 1
 
        'If NewText = LCase(Left(CboCoGrpPayor.List(i), Len(NewText))) Then
                'ValidCount = ValidCount + 1
                'ValidValue = CboCoGrpPayor.List(i)
        'End If
        'Next
        'If ValidCount <= 1 Then KeyAscii = 0
             'If ValidCount = 1 Then
               'CboCoGrpPayor.Text = ValidValue
               'CboCoGrpPayor.SelStart = 0
               'CboCoGrpPayor.SelLength = Len(ValidValue)
             'End If
        'End If
'End Sub

'Private Sub cboCoPrePost_Change()
    'If InStr(cboCoPrePost.Text, "null") Then
       'cboCoPrePost.Enabled = False
    'End If
'End Sub

'Private Sub cboCoPrePost_KeyPress(KeyAscii As Integer)
'Dim NewText As String
'Dim ValidCount As Integer
'Dim ValidValue As String
'Dim i As Integer
    'Clipboard.Clear
    'If KeyAscii >= 32 And KeyAscii <> 127 Then
    'NewText = LCase(Left(cboCoPrePost.Text, cboCoPrePost.SelStart) + Chr(KeyAscii) + Mid(cboCoPrePost.Text, cboCoPrePost.SelStart + cboCoPrePost.SelLength + 1))

    'ValidCount = 0
    'ValidValue = ""

    'For i = 0 To cboCoPrePost.ListCount - 1
 
        'If NewText = LCase(Left(cboCoPrePost.List(i), Len(NewText))) Then
                'ValidCount = ValidCount + 1
                'ValidValue = cboCoPrePost.List(i)
        'End If
        'Next
        'If ValidCount <= 1 Then KeyAscii = 0
             'If ValidCount = 1 Then
               'cboCoPrePost.Text = ValidValue
               'cboCoPrePost.SelStart = 0
               'cboCoPrePost.SelLength = Len(ValidValue)
             'End If
        'End If
'End Sub

'Private Sub cboCoTax_Change()
    'If InStr(cboCoTax.Text, "null") Then
       'cboCoTax.Enabled = False
    'End If
'End Sub

'Private Sub cboCoTax_KeyPress(KeyAscii As Integer)
'Dim NewText As String
'Dim ValidCount As Integer
'Dim ValidValue As String
'Dim i As Integer
    'Clipboard.Clear
    'If KeyAscii >= 32 And KeyAscii <> 127 Then
    'NewText = LCase(Left(cboCoTax.Text, cboCoTax.SelStart) + Chr(KeyAscii) + Mid(cboCoTax.Text, cboCoTax.SelStart + cboCoTax.SelLength + 1))

    'ValidCount = 0
    'ValidValue = ""

    'For i = 0 To cboCoTax.ListCount - 1
 
        'If NewText = LCase(Left(cboCoTax.List(i), Len(NewText))) Then
                'ValidCount = ValidCount + 1
                'ValidValue = cboCoTax.List(i)
        'End If
        'Next
        'If ValidCount <= 1 Then KeyAscii = 0
             'If ValidCount = 1 Then
               'cboCoTax.Text = ValidValue
               'cboCoTax.SelStart = 0
               'cboCoTax.SelLength = Len(ValidValue)
             'End If
        'End If
'End Sub

Private Sub cboInfoClnGrp_Change()
    If InStr(CboInfoClnGrp.Text, "null") Then
       CboInfoClnGrp.Enabled = False
    End If
End Sub

Private Sub cboInfoClnGrp_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboInfoClnGrp.Text, CboInfoClnGrp.SelStart) + Chr(KeyAscii) + Mid(CboInfoClnGrp.Text, CboInfoClnGrp.SelStart + CboInfoClnGrp.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboInfoClnGrp.ListCount - 1
 
        If NewText = LCase(Left(CboInfoClnGrp.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboInfoClnGrp.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboInfoClnGrp.Text = ValidValue
               CboInfoClnGrp.SelStart = 0
               CboInfoClnGrp.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboInfoClnCity_Change()
    If InStr(cboInfoClnCity.Text, "null") Then
       cboInfoClnCity.Enabled = False
    End If
End Sub

Private Sub cboInfoClnCity_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboInfoClnCity.Text, cboInfoClnCity.SelStart) + Chr(KeyAscii) + Mid(cboInfoClnCity.Text, cboInfoClnCity.SelStart + cboInfoClnCity.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboInfoClnCity.ListCount - 1
 
        If NewText = LCase(Left(cboInfoClnCity.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboInfoClnCity.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboInfoClnCity.Text = ValidValue
               cboInfoClnCity.SelStart = 0
               cboInfoClnCity.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboInfoClnState_Change()
    If InStr(cboInfoClnState.Text, "null") Then
       cboInfoClnState.Enabled = False
    End If
End Sub

Private Sub cboInfoClnState_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboInfoClnState.Text, cboInfoClnState.SelStart) + Chr(KeyAscii) + Mid(cboInfoClnState.Text, cboInfoClnState.SelStart + cboInfoClnState.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboInfoClnState.ListCount - 1
 
        If NewText = LCase(Left(cboInfoClnState.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboInfoClnState.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboInfoClnState.Text = ValidValue
               cboInfoClnState.SelStart = 0
               cboInfoClnState.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboInfoClnType_Change()
    If InStr(cboInfoClnType.Text, "null") Then
       cboInfoClnType.Enabled = False
    End If
End Sub

Private Sub cboInfoClnType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboInfoClnType.Text, cboInfoClnType.SelStart) + Chr(KeyAscii) + Mid(cboInfoClnType.Text, cboInfoClnType.SelStart + cboClnConversion.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboInfoClnType.ListCount - 1
 
        If NewText = LCase(Left(cboInfoClnType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboInfoClnType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboInfoClnType.Text = ValidValue
               cboInfoClnType.SelStart = 0
               cboInfoClnType.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboInfoClnBNFType_Change()
    If InStr(cboInfoClnBNFType.Text, "null") Then
       cboInfoClnBNFType.Enabled = False
    End If
End Sub

Private Sub cboInfoClnBNFType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboInfoClnBNFType.Text, cboInfoClnBNFType.SelStart) + Chr(KeyAscii) + Mid(cboInfoClnBNFType.Text, cboInfoClnBNFType.SelStart + cboInfoClnBNFType.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboInfoClnBNFType.ListCount - 1
 
        If NewText = LCase(Left(cboInfoClnBNFType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboInfoClnBNFType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboInfoClnBNFType.Text = ValidValue
               cboInfoClnBNFType.SelStart = 0
               cboInfoClnBNFType.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboInfoClnXray_Change()
    If InStr(cboInfoClnXray.Text, "null") Then
       cboInfoClnXray.Enabled = False
    End If
End Sub

Private Sub cboInfoClnXray_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboInfoClnXray.Text, cboInfoClnXray.SelStart) + Chr(KeyAscii) + Mid(cboInfoClnXray.Text, cboInfoClnXray.SelStart + cboInfoClnXray.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboInfoClnXray.ListCount - 1
 
        If NewText = LCase(Left(cboInfoClnXray.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboInfoClnXray.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboInfoClnXray.Text = ValidValue
               cboInfoClnXray.SelStart = 0
               cboInfoClnXray.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboInfoClnPanel_Change()
    If InStr(cboInfoClnPanel.Text, "null") Then
       cboInfoClnPanel.Enabled = False
    End If
End Sub

Private Sub cboInfoClnPanel_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboInfoClnPanel.Text, cboInfoClnPanel.SelStart) + Chr(KeyAscii) + Mid(cboInfoClnPanel.Text, cboInfoClnPanel.SelStart + cboInfoClnPanel.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboInfoClnPanel.ListCount - 1
 
        If NewText = LCase(Left(cboInfoClnPanel.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboInfoClnPanel.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboInfoClnPanel.Text = ValidValue
               cboInfoClnPanel.SelStart = 0
               cboInfoClnPanel.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboInfoCln24hrs_Change()
    If InStr(cboInfoCln24hrs.Text, "null") Then
       cboInfoCln24hrs.Enabled = False
    End If
End Sub

Private Sub cboInfoCln24hrs_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboInfoCln24hrs.Text, cboInfoCln24hrs.SelStart) + Chr(KeyAscii) + Mid(cboInfoCln24hrs.Text, cboInfoCln24hrs.SelStart + cboInfoCln24hrs.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboInfoCln24hrs.ListCount - 1
 
        If NewText = LCase(Left(cboInfoCln24hrs.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboInfoCln24hrs.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboInfoCln24hrs.Text = ValidValue
               cboInfoCln24hrs.SelStart = 0
               cboInfoCln24hrs.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboInfoFFS_Change()
    If InStr(cboInfoFFS.Text, "null") Then
       cboInfoFFS.Enabled = False
    End If
End Sub

Private Sub cboInfoFFS_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboInfoFFS.Text, cboInfoFFS.SelStart) + Chr(KeyAscii) + Mid(cboInfoFFS.Text, cboInfoFFS.SelStart + cboInfoFFS.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboInfoFFS.ListCount - 1
 
        If NewText = LCase(Left(cboInfoFFS.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboInfoFFS.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboInfoFFS.Text = ValidValue
               cboInfoFFS.SelStart = 0
               cboInfoFFS.SelLength = Len(ValidValue)
             End If
        End If
End Sub


Private Sub cboOnlineStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboOnlineStatus.Text, cboOnlineStatus.SelStart) + Chr(KeyAscii) + Mid(cboOnlineStatus.Text, cboOnlineStatus.SelStart + cboOnlineStatus.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboOnlineStatus.ListCount - 1
 
        If NewText = LCase(Left(cboOnlineStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboOnlineStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboOnlineStatus.Text = ValidValue
               cboOnlineStatus.SelStart = 0
               cboOnlineStatus.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboGLFormat_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboGLFormat.Text, cboGLFormat.SelStart) + Chr(KeyAscii) + Mid(cboGLFormat.Text, cboGLFormat.SelStart + cboGLFormat.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboGLFormat.ListCount - 1
 
        If NewText = LCase(Left(cboGLFormat.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboGLFormat.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboGLFormat.Text = ValidValue
               cboGLFormat.SelStart = 0
               cboGLFormat.SelLength = Len(ValidValue)
             End If
        End If
End Sub


Private Sub cboPymtInfoClnCity_Change()
 If InStr(cboPymtInfoClnCity.Text, "null") Then
       cboPymtInfoClnCity.Enabled = False
    End If
End Sub

Private Sub cboPymtInfoClnCity_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPymtInfoClnCity.Text, cboPymtInfoClnCity.SelStart) + Chr(KeyAscii) + Mid(cboPymtInfoClnCity.Text, cboPymtInfoClnCity.SelStart + cboPymtInfoClnCity.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPymtInfoClnCity.ListCount - 1
 
        If NewText = LCase(Left(cboPymtInfoClnCity.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPymtInfoClnCity.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboPymtInfoClnCity.Text = ValidValue
               cboPymtInfoClnCity.SelStart = 0
               cboPymtInfoClnCity.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Call GRPCOMCLNComboListing
    Else
        CboGrpComName.Clear
        CboCLNName.Clear
    End If
End Sub

'Private Sub Check2_Click()
    'If Check2.Value = 1 Then
        'Call GRPCOMCLNComboListing
    'Else
        'CboCLNName.Clear
    'End If
'End Sub

'#################### Start Add, Edit, Delete, Search, Save, Exit ####################'
Private Sub cmdAdd_Click()
    EditFlag = False
    Call DisableCmdButton(Me)
    Call EntriesEnable(True)
    cmdSave.Enabled = True
    cmdAdd.Enabled = True
   
    Select Case SSTab1.Tab
        'Case 0
            'Call ClearForm(Me)
            'txtCompanyName.SetFocus
            'txtGrpCoCode.Enabled = False
        Case 0
            Call ClearForm(Me)
            txtClnName.SetFocus
            txtClnGrpCode.Enabled = False
        Case 1
            Call ClearForm(Me)
            txtDocName.SetFocus
            txtDocCode.Enabled = False
        Case 2
            Call ClearForm(Me)
            txtInfoClnName.SetFocus
            txtInfoCLNCode.Enabled = True
        Case 3
            CboGrpComName.Text = ""
            CboCLNName.Text = ""
            lstGrpComCLN.ListItems.Clear
            CboGrpComName.Enabled = True
            CboGrpComName.SetFocus
        Case 4
            txtCAClnCode.Enabled = True
            cmdCLNLookup.Enabled = True
            txtEffDateFrom.Enabled = True
            txtEffDateTo.Enabled = True
            txtCAClnCode.SetFocus
    End Select

End Sub

Private Sub cmdBack_Click()
    lstClinicInfoList.Height = 1740
    lstClinicInfoList.Top = 4560
    Frame4.Visible = True
End Sub

Private Sub cmdCLNLookup_Click()
    frmSearchCLN.Show vbModal
End Sub

Private Sub CmdView_Click()
    lstClinicInfoList.Height = 5820
    lstClinicInfoList.Width = 11235
    lstClinicInfoList.Top = 360
    Frame4.Visible = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    
    Select Case SSTab1.Tab
        'Case 0
            'Call ClearForm(Me)
        Case 0
            Call ClearForm(Me)
        Case 1
            Call ClearForm(Me)
        Case 2
            Call ClearForm(Me)
        Case 3
            CboGrpComName.Text = ""
            CboCLNName.Text = ""
            lstGrpComCLN.ListItems.Clear
            CboGrpComName.Enabled = True
            CboGrpComName.SetFocus
        Case 4
            lblClinicName.Caption = ""
            Call ClearForm(Me)
    End Select
    
    cmdAdd.Enabled = True
    txtTotalRecord = 0
    lstGrpClinicList.ListItems.Clear
    'lstGrpCoList.ListItems.Clear
    lstDoctorList.ListItems.Clear
    lstClinicInfoList.ListItems.Clear
    lstGrpComCLN.ListItems.Clear
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdSave.Enabled = False
    Call EntriesEnable(True)

End Sub

Private Sub cmdDelete_Click()
Screen.MousePointer = vbHourglass
        
    'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        
        Select Case SSTab1.Tab
        'Case 0
           'Call DeleteGrpCompany
           'cmdAdd.Enabled = True
        Case 0
            Call DeleteClinicGrp
        Case 1
            Call DeleteDoctor
        Case 2
            Call DeleteClinic
        Case 3
            Call DeleteGrpComClinic
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

Private Sub cmdEdit_Click()
    EditFlag = True
    Call EntriesEnable(True)
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
     
    'If txtGrpCoCode <> "" Then
       'txtGrpCoCode.Enabled = False
       'txtCompanyName.Enabled = False
       'txtCoAdd1.SetFocus
    'End If

    If txtClnGrpCode <> "" Then
       txtClnGrpCode.Enabled = False
       txtClnName.Enabled = False
       txtClnGrpName.Enabled = True
       txtClnGrpName.SetFocus
    End If
    
    If txtDocCode <> "" Then
       txtDocCode.Enabled = False
       txtDocName.Enabled = False
    End If

    If txtInfoCLNCode <> "" Then
       txtInfoCLNCode.Enabled = False
       txtInfoClnName.Enabled = False
    End If
    
    If CboGrpComName <> "" Then
       CboGrpComName.Enabled = True
       CboCLNName.Enabled = True
    End If
    
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
    
    'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        Select Case SSTab1.Tab
            'Case 0
                'Call PrintGRPCO
            Case 0
                Call PrintGRPCLN
            Case 1
                Call PrintDoctor
            Case 2
                Call ExportToExcelClinic(lstClinicInfoList)
            Case 3
                Call PrintGrpComCln
        End Select
        cmdAdd.Enabled = True
        cmdEdit.Enabled = False
    
   ' MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing
    'medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' MediConnCISWS.Close
   ' Set MediConnCISWS = Nothing
   ' MediConnCISPayment.Close
  '  Set MediConnCISPayment = Nothing
    
    Screen.MousePointer = Default
End Sub

Private Sub CmdRefresh_Click()
    Screen.MousePointer = vbHourglass
        
   ' MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        Call EntriesEnable(True)
        Select Case SSTab1.Tab
        'Case 0
            'Call ClearForm(Me)
            'Call GRPCompanyComboListing
        Case 0
            Call ClearForm(Me)
            Call GRPClinicComboListing
        Case 1
            Call ClearForm(Me)
            Call DoctorListing
        Case 2
            Call ClearForm(Me)
            Call CLNComboListing
        Case 3
            CboGrpComName.Text = ""
            CboCLNName.Text = ""
            lstGrpComCLN.ListItems.Clear
            Call GRPCOMCLNComboListing
        Case 4
             Call ClearForm(Me)
            Call ClinicAccessListing
        End Select
        txtTotalRecord = 0
    
   ' MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' MediConnCISWS.Close
  '  Set MediConnCISWS = Nothing
   ' MediConnCISPayment.Close
   ' Set MediConnCISPayment = Nothing
    Screen.MousePointer = Default
End Sub

Private Sub cmdSave_Click()
Screen.MousePointer = vbHourglass
        
  '  MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
  '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
        Select Case SSTab1.Tab
            'Case 0
                'Call SaveGrpCompany
                'Call ClearForm(Me)
                'txtGrpCoCode.Enabled = True
                'txtGrpCoName.Enabled = True
                'txtGrpCoCode.SetFocus
            Case 0
                Call SaveClinicGrp
                'Call ClearForm(Me)
            Case 1
                Call SaveDoctor
                'Call ClearForm(Me)
                'txtDocCode.Enabled = True
                'txtDocName.Enabled = True
                'txtDocCode.SetFocus
            Case 2
                Call SaveClinic
            Case 3
                Call SaveGrpComClinic
            Case 4
                Call SaveClnAccess
            End Select
            
    '    MediConnCISDB.Close
     '   Set MediConnCISDB = Nothing
    '    medixmasterconn.Close
     '   Set medixmasterconn = Nothing
      '  MediConnCISWS.Close
    '    Set MediConnCISWS = Nothing
    '    MediConnCISPayment.Close
    '    Set MediConnCISPayment = Nothing
    '    medixmasterconnMaint.Close
    '    Set medixmasterconnMaint = Nothing
Screen.MousePointer = Default

End Sub

Private Sub cmdSearch_Click()
Screen.MousePointer = vbHourglass

 '   MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
'    MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
'    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
 '   MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"

        Select Case SSTab1.Tab
            'Case 0
                'Call SearchGrpCompany
                'Call ClearForm(Me)
                'txtGrpCoCode.Enabled = True
                'txtCompanyName.Enabled = True
                'txtGrpCoCode.SetFocus
            Case 0
                Call SearchClinicGrp
                Call ClearForm(Me)
                Call EntriesEnable(False)
                txtClnGrpCode.Enabled = True
                txtClnGrpName.Enabled = True
                txtClnGrpCode.SetFocus
            Case 1
                Call SearchDoctor
                Call ClearForm(Me)
                txtDocCode.Enabled = True
                txtDocName.Enabled = True
                txtDocCode.SetFocus
            Case 2
                Call SearchCLN
                Call ClearForm(Me)
                txtInfoCLNCode.Enabled = True
                txtInfoClnName.Enabled = True
                txtInfoCLNCode.SetFocus
            Case 3
                'CboGrpComName.Text = ""
                'CboCLNName.Text = ""
                lstGrpComCLN.ListItems.Clear
                Call SearchGrpComCLN
                CboGrpComName.Enabled = True
                CboCLNName.Enabled = True
                CboGrpComName.SetFocus
            Case 4
                lstClinicAccess.ListItems.Clear
                Call SearchClinicAccess
            End Select
        
        'MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
       ' medixmasterconn.Close
      '  Set medixmasterconn = Nothing
      '  MediConnCISWS.Close
      '  Set MediConnCISWS = Nothing
       ' MediConnCISPayment.Close
       ' Set MediConnCISPayment = Nothing
Screen.MousePointer = Default

End Sub
'#################### End Add, Edit, Delete, Search, Save, Exit ####################'



Private Sub dtp_Change()

End Sub

'Private Sub Form_Activate()
    'SSTab1.Tab = 0
'End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "frmMaintenanceEntity")
   ' MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

        Call GRPClinicComboListing
        Call CLNComboListing
        'dtpEffDateFrom.Value = Now
        'dtpEffDateTo.Value = Now
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
    'MediConnCISPayment.Close
    'Set MediConnCISPayment = Nothing
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'--- Enable PrintScreen feature
'--- IMPORTANT - Don't forget to set .KeyPreview = True for this form at design time!!
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
        
End Sub


'**********************************Start Group Company**************************************
'Private Sub SaveGrpCompany()
'Dim RecordSaved
'Dim CompanyGrp As New ADODB.Recordset
'Dim SaveCompanyGrp As New ADODB.Recordset
'Dim strsql As String
'Dim COMSeq As String
        
    'If txtCompanyName = "" Then
        'MsgBox "Please Enter Company Name", vbOKOnly + vbCritical, DialogBoxTitle
        'txtCompanyName.SetFocus
    'ElseIf CboCoGrpPayor = "" Then
        'MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        'CboCoGrpPayor.SetFocus
    'ElseIf txtCoAdd1 = "" Then
        'MsgBox "Please Enter Group Company Address", vbOKOnly + vbCritical, DialogBoxTitle
        'txtCoAdd1.SetFocus
    'ElseIf txtCoEffDate = "__/__/____" Then
        'MsgBox "Please Enter Effective Date", vbOKOnly + vbCritical, DialogBoxTitle
        'txtCoEffDate.SetFocus
    'Else
        'strsql = "Select * From GRPCompany Where GRPCoName = '" & Trim(txtCompanyName) & "'"
        'CompanyGrp.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
      'If CompanyGrp.RecordCount > 0 And EditFlag = False Then
      'MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
      'CompanyGrp.Close
      'txtCompanyName.SetFocus
      'txtCompanyName.SelStart = 0
      'txtCompanyName.SelLength = Len(txtCompanyName)
      'Set CompanyGrp = Nothing
      'Exit Sub
      'Else
      'CompanyGrp.Close
      'Set CompanyGrp = Nothing
      'RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        'If RecordSaved = vbYes Then
            'If EditFlag = False Then
                'CompanyGrp.Open "GRPCompany", MediConnCISDB, adOpenKeyset, adLockOptimistic
                'CompanyGrp.AddNew
                'COMSeq = GenerateGRPCoSeq(Mid(txtCompanyName, 1, 1))
            'Else
                'strsql = "Select * From GRPCompany Where GRPCoCode = '" & Trim(txtGrpCoCode) & "'"
                'CompanyGrp.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                'COMSeq = txtGrpCoCode
            'End If
            
                'With CompanyGrp
                    '!GRPCoCode = ChkStr(COMSeq)
                    '!GrpCoName = ChkStr(txtCompanyName)
                    '!GRPCoPAYCode = Mid(CboCoGrpPayor, 1, 2)
                    '!GRPCoAdd1 = IIf(Len(txtCoAdd1), ChkStr(txtCoAdd1), Null)
                    '!GRPCoAdd2 = IIf(Len(txtCoAdd2), ChkStr(txtCoAdd2), Null)
                    '!GRPCoAdd3 = IIf(Len(txtCoAdd3), ChkStr(txtCoAdd3), Null)
                    '!GRPCoAdmContact = IIf(Len(txtCoAdmContact), ChkStr(txtCoAdmContact), Null)
                    '!GRPCoContact = IIf(Len(txtCoContact), ChkStr(txtCoContact), Null)
                    '!GRPCoCEO = IIf(Len(txtCoCEO), ChkStr(txtCoCEO), Null)
                    '!GRPCoPhone = IIf(Len(txtCoPhone), ChkStr(txtCoPhone), Null)
                    '!GRPCoFax = IIf(Len(txtCoFax), ChkStr(txtCoFax), Null)
                    '!GRPCoEmail = IIf(Len(txtCoEmail), ChkStr(txtCoEmail), Null)
                    '!GRPCoHomepage = IIf(Len(txtCoHomepage), ChkStr(txtCoHomepage), Null)
                    '!GRPCoEffDate = IIf(IsDate(txtCoEffDate), txtCoEffDate, Null)
                    '!GRPCoTax = Mid(cboCoTax, 1, 1)
                    '!GRPCoConversion = Mid(cboCoConversion, 1, 1)
                    '!GRPCoPrePost = Mid(cboCoConversion, 1, 1)
                    '!GRPCoLastUpdateUser = Logon
                    '!GRPCoLastUpdateDate = Date
                'End With
                'CompanyGrp.Update
                'Call ClearForm(Me)
                'Call GRPCompanyListing
                'If EditFlag = False Then
                    'MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                'Else
                    'MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                'End If
                
            'CompanyGrp.Close
            'Set SaveCompanyGrp = Nothing
        'End If
        'EditFlag = False
    'End If
'End If
'End Sub

'Private Sub DeleteGrpCompany()
'Dim DeleteGrpCo As New ADODB.Recordset
'Dim rsChkCo As New ADODB.Recordset
'Dim strsql, sqlCompany As String
'Dim DeleteRecord

'EditFlag = False
'If txtGrpCoCode = "" Then
   'MsgBox "Please Enter Company Group Code", vbOKOnly + vbCritical, DialogBoxTitle
   'txtGrpCoCode.SetFocus
'ElseIf lstGrpCoList.ListItems.Count = 0 Then
   'MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
'Else
   'sqlCompany = "Select GRPCoCode from GRPCompany where GRPCoCode = '" & RemStr(txtGrpCoCode) & "'"
   
   'rsChkCo.MaxRecords = 1
   'rsChkCo.Open sqlCompany, MediConnCISDB, adOpenKeyset, adLockOptimistic
   
   'If rsChkCo.RecordCount Then
        'MsgBox "This record cannot be deleted because it is used in Company Table!", vbOKOnly + vbCritical, "Error Message"
   'Else
   'DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     'If DeleteRecord = vbYes And rsChkCo.RecordCount = 0 Then
           'strsql = "Delete From GRPCompany Where GRPCoCode = '" & ChkStr(lstGrpCoList.SelectedItem.Text) & "'"
           'DeleteGrpCo.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
           'MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           'cmdSave.Enabled = True
           'Call ClearForm(Me)
           'Call GRPCompanyListing
           'Call DisableCmdButton(Me)
           'Call EntriesEnable(True)
           'txtGrpCoCode.SetFocus
         'End If
     'End If
'End If
'Exit Sub
'End Sub

'Private Sub GRPCompanyListing(Optional strAppend As String)
'Dim rst2 As New ADODB.Recordset
'Dim CompanyGrpMaster As ListItem
'Dim sql As String
'Dim TotalGRPCOMCount As Integer
    
    'sql = "Select * from GRPCompany where 0=0 "
   
    'If Len(Trim(strAppend)) Then
        'sql = sql + strAppend
    'End If
              
    'TotalGRPCOMCount = 0
    'lstGrpCoList.ListItems.Clear
    'rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    'If rst2.RecordCount = 0 Then
        'MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    'Else
        'Do Until rst2.EOF
            'Set CompanyGrpMaster = lstGrpCoList.ListItems.Add(, , rst2!GRPCoCode)
                'CompanyGrpMaster.SubItems(1) = rst2!GrpCoName
                'CompanyGrpMaster.SubItems(2) = rst2!GRPCoPAYCode
                'rst2.MoveNext
        'TotalGRPCOMCount = TotalGRPCOMCount + 1
        'txtTotalRecord = TotalGRPCOMCount
        'Loop
    'rst2.Close
    'End If
'End Sub

'Private Sub GRPCompanyComboListing()
'Dim PAY As New ADODB.Recordset
        
    'PAY.Open "PAY", MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    'CboCoGrpPayor.Clear

    'Do Until PAY.EOF
        'CboCoGrpPayor.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
        'PAY.MoveNext
    'Loop
            
    'PAY.Close
    
    'cboCoConversion.Clear
    'cboCoConversion.AddItem "N" & " - " & "No"
    'cboCoConversion.AddItem "Y" & " - " & "Yes"
    
    'cboCoTax.Clear
    'cboCoTax.AddItem "N" & " - " & "No"
    'cboCoTax.AddItem "Y" & " - " & "Yes"
    
    'cboCoPrePost.Clear
    'cboCoPrePost.AddItem "N" & " - " & "No"
    'cboCoPrePost.AddItem "Y" & " - " & "Yes"
    
'End Sub

Private Sub lstClinicInfoList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim StrSql As String
    
    If lstClinicInfoList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            
        StrSql = "Select * from CLN Where CLNCode = '" & lstClinicInfoList.SelectedItem.Text & "'"
        rst3.Open StrSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        
        txtInfoCLNCode = rst3!CLNCode
        
        If Trim(rst3!CLNName) <> " " Then
            txtInfoClnName = rst3!CLNName
        End If
        
        If Trim(rst3!ClinicGrpCode) <> " " Then
            CboInfoClnGrp = rst3!ClinicGrpCode
        End If
        
        If Trim(rst3!CLNAdd1) <> " " Then
            txtInfoClnAdd1 = rst3!CLNAdd1
        End If
        
        If Trim(rst3!CLNAdd2) <> " " Then
            txtInfoClnAdd2 = rst3!CLNAdd2
        End If
        
        If Trim(rst3!CLNAdd3) <> " " Then
            txtInfoClnAdd3 = rst3!CLNAdd3
        End If
        
        If Trim(rst3!CLNPhone) <> " " Then
            txtInfoClnTelNo = rst3!CLNPhone
        End If
        
        If Trim(rst3!CLNFax) <> " " Then
            txtInfoClnFaxNo = rst3!CLNFax
        End If
        
        If Trim(rst3!CLNEMAIL) <> " " Then
            txtInfoClnEmail = rst3!CLNEMAIL
        End If
        
        If Trim(rst3!CLNHomepage) <> " " Then
            txtInfoClnHome = rst3!CLNHomepage
        End If
        
        If Trim(rst3!CLNRepresentative) <> " " Then
            txtInfoClnContact = rst3!CLNRepresentative
        End If
        
        If Trim(rst3!CLNDocName) <> " " Then
            txtInfoClnDocName = rst3!CLNDocName
        End If
        
        If Trim(rst3!CLNBNFType) <> " " Then
            cboInfoClnBNFType = rst3!CLNBNFType
        End If
        
        If Trim(rst3!CLNForRef) <> " " Then
            txtInfoClnFormema = rst3!CLNForRef
        End If
        
        If Trim(rst3!CLNNoofDoc) <> " " Then
            txtInfoClnNoDoc = rst3!CLNNoofDoc
        End If
               
        If Trim(rst3!CLNType) <> " " Then
            cboInfoClnType = rst3!CLNType
        End If
        
        If Trim(rst3!CLNXRay) <> " " Then
            cboInfoClnXray = rst3!CLNXRay
        End If
        
        If Trim(rst3!CLNPanel) <> " " Then
            cboInfoClnPanel = rst3!CLNPanel
        End If
        
        If Trim(rst3!CLNState) <> " " Then
            cboInfoClnState = rst3!CLNState
        End If
        
        If Trim(rst3!CLNCity) <> " " Then
            cboInfoClnCity = rst3!CLNCity
        End If
        
        If Trim(rst3!CLNPostCode) <> " " Then
            txtInfoClnPostCode = rst3!CLNPostCode
        End If
               
        If Trim(rst3!CLNinitial) <> " " Then
            txtInfoClnMIX = rst3!CLNinitial
        End If
                
        If rst3!CLNReceivedDate <> "__/__/____" Then
            txtInfoClnReceivedDate = rst3!CLNReceivedDate
        End If
                
        If Trim(rst3!CLNBankAcc) <> " " Then
            txtInfoClnBankAcc = rst3!CLNBankAcc
        End If
        
        If Trim(rst3!CLNBankName) <> " " Then
            txtInfoClnBank = rst3!CLNBankName
        End If
        
        If Trim(rst3!CLNBankBrn) <> " " Then
            txtInfoClnBranch = rst3!CLNBankBrn
        End If
        
        If Trim(rst3!CLNAccPerson) <> " " Then
            txtInfoClnAccPerson = rst3!CLNAccPerson
        End If
        
        If Trim(rst3!CLN24Hrs) <> " " Then
            cboInfoCln24hrs = rst3!CLN24Hrs
        End If

        If Trim(rst3!CLNRemark) <> " " Then
            txtInfoClnRemark = rst3!CLNRemark
        End If
        
        If Trim(rst3!CLNLocation) <> " " Then
            TxtLocation = rst3!CLNLocation
        End If
        
        If Trim(rst3!CLNFFSStatus) <> " " Then
            cboInfoFFS = rst3!CLNFFSStatus
        End If
        
        If Trim(rst3!CLNMCOProcfeePerc) <> " " Then
            txtFFSPerc = rst3!CLNMCOProcfeePerc
        End If
        
        If Len(rst3!CLNOperatingHr) Then
            txtOperatinghr = rst3!CLNOperatingHr
        End If
                
        If Len(rst3!CLNPName) Then
            txtPymtInfoClnName = rst3!CLNPName
        End If
        
        If Len(rst3!CLNPAdd1) Then
            txtPymtInfoClnAdd1 = rst3!CLNPAdd1
        End If
        
        If Len(rst3!CLNPAdd2) Then
            txtPymtInfoClnAdd2 = rst3!CLNPAdd2
        End If
        
        If Len(rst3!CLNPAdd3) Then
            txtPymtInfoClnAdd3 = rst3!CLNPAdd3
        End If
        
        If Len(rst3!CLNPPostCode) Then
            txtPymtInfoClnPostCode = rst3!CLNPPostCode
        End If
        
        If Len(rst3!CLNPcity) Then
            cboPymtInfoClnCity = rst3!CLNPcity
        End If
        
        If Len(rst3!CLNPState) Then
            cboPymtInfoClnState = rst3!CLNPState
        End If
        
        If Len(rst3!CLNOnlineStatus) Then
            cboOnlineStatus = rst3!CLNOnlineStatus
        End If
        
        If Len(rst3!CLNGLFormat) Then
            cboGLFormat = rst3!CLNGLFormat
        End If
        
        If Len(rst3!CLNRegNo) Then
            txtRegNo = rst3!CLNRegNo
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing


   End If
End Sub

Private Sub lstClinicInfoList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim StrSql As String
    
    If lstClinicInfoList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

        StrSql = "Select * from CLN Where CLNCode = '" & lstClinicInfoList.SelectedItem.Text & "'"
        rst3.Open StrSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        
        txtInfoCLNCode = rst3!CLNCode
        
        If Trim(rst3!CLNName) <> " " Then
            txtInfoClnName = rst3!CLNName
        End If
        
        If Trim(rst3!ClinicGrpCode) <> " " Then
            CboInfoClnGrp = rst3!ClinicGrpCode
        End If
        
        If Trim(rst3!CLNAdd1) <> " " Then
            txtInfoClnAdd1 = rst3!CLNAdd1
        End If
        
        If Trim(rst3!CLNAdd2) <> " " Then
            txtInfoClnAdd2 = rst3!CLNAdd2
        End If
        
        If Trim(rst3!CLNAdd3) <> " " Then
            txtInfoClnAdd3 = rst3!CLNAdd3
        End If
        
        If Trim(rst3!CLNPhone) <> " " Then
            txtInfoClnTelNo = rst3!CLNPhone
        End If
        
        If Trim(rst3!CLNFax) <> " " Then
            txtInfoClnFaxNo = rst3!CLNFax
        End If
        
        If Trim(rst3!CLNEMAIL) <> " " Then
            txtInfoClnEmail = rst3!CLNEMAIL
        End If
        
        If Trim(rst3!CLNHomepage) <> " " Then
            txtInfoClnHome = rst3!CLNHomepage
        End If
        
        If Trim(rst3!CLNRepresentative) <> " " Then
            txtInfoClnContact = rst3!CLNRepresentative
        End If
        
        If Trim(rst3!CLNDocName) <> " " Then
            txtInfoClnDocName = rst3!CLNDocName
        End If
        
        If Trim(rst3!CLNBNFType) <> " " Then
            cboInfoClnBNFType = rst3!CLNBNFType
        End If
        
        If Trim(rst3!CLNForRef) <> " " Then
            txtInfoClnFormema = rst3!CLNForRef
        End If
        
        If Trim(rst3!CLNNoofDoc) <> " " Then
            txtInfoClnNoDoc = rst3!CLNNoofDoc
        End If
               
        If Trim(rst3!CLNType) <> " " Then
            cboInfoClnType = rst3!CLNType
        End If
        
        If Trim(rst3!CLNXRay) <> " " Then
            cboInfoClnXray = rst3!CLNXRay
        End If
        
        If Trim(rst3!CLNPanel) <> " " Then
            cboInfoClnPanel = rst3!CLNPanel
        End If
        
        If Trim(rst3!CLNState) <> " " Then
            cboInfoClnState = rst3!CLNState
        End If
        
        If Trim(rst3!CLNCity) <> " " Then
            cboInfoClnCity = rst3!CLNCity
        End If
        
        If Trim(rst3!CLNPostCode) <> " " Then
            txtInfoClnPostCode = rst3!CLNPostCode
        End If
               
        If Trim(rst3!CLNinitial) <> " " Then
            txtInfoClnMIX = rst3!CLNinitial
        End If
                
        If rst3!CLNReceivedDate <> "__/__/____" Then
            txtInfoClnReceivedDate = rst3!CLNReceivedDate
        End If
                
        If Trim(rst3!CLNBankAcc) <> " " Then
            txtInfoClnBankAcc = rst3!CLNBankAcc
        End If
        
        If Trim(rst3!CLNBankName) <> " " Then
            txtInfoClnBank = rst3!CLNBankName
        End If
        
        If Trim(rst3!CLNBankBrn) <> " " Then
            txtInfoClnBranch = rst3!CLNBankBrn
        End If
        
        If Trim(rst3!CLNAccPerson) <> " " Then
            txtInfoClnAccPerson = rst3!CLNAccPerson
        End If
        
        If Trim(rst3!CLN24Hrs) <> " " Then
            cboInfoCln24hrs = rst3!CLN24Hrs
        End If

        If Trim(rst3!CLNRemark) <> " " Then
            txtInfoClnRemark = rst3!CLNRemark
        End If
        
        If Trim(rst3!CLNLocation) <> " " Then
            TxtLocation = rst3!CLNLocation
        End If
        
        If Trim(rst3!CLNFFSStatus) <> " " Then
            cboInfoFFS = rst3!CLNFFSStatus
        End If
        
        If Trim(rst3!CLNMCOProcfeePerc) <> " " Then
            txtFFSPerc = rst3!CLNMCOProcfeePerc
        End If
        
        If Len(rst3!CLNOperatingHr) Then
            txtOperatinghr = rst3!CLNOperatingHr
        End If
                
        If Len(rst3!CLNPName) Then
            txtPymtInfoClnName = rst3!CLNPName
        End If
        
        If Len(rst3!CLNPAdd1) Then
            txtPymtInfoClnAdd1 = rst3!CLNPAdd1
        End If
        
        If Len(rst3!CLNPAdd2) Then
            txtPymtInfoClnAdd2 = rst3!CLNPAdd2
        End If
        
        If Len(rst3!CLNPAdd3) Then
            txtPymtInfoClnAdd3 = rst3!CLNPAdd3
        End If
        
        If Len(rst3!CLNPPostCode) Then
            txtPymtInfoClnPostCode = rst3!CLNPPostCode
        End If
        
        If Len(rst3!CLNPcity) Then
            cboPymtInfoClnCity = rst3!CLNPcity
        End If
        
        If Len(rst3!CLNPState) Then
            cboPymtInfoClnState = rst3!CLNPState
        End If
        
        If Len(rst3!CLNOnlineStatus) Then
            cboOnlineStatus = rst3!CLNOnlineStatus
        End If
        
        If Len(rst3!CLNGLFormat) Then
            cboGLFormat = rst3!CLNGLFormat
        End If
        
        If Len(rst3!CLNRegNo) Then
            txtRegNo = rst3!CLNRegNo
        End If


        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
     '   medixmasterconn.Close
     '   Set medixmasterconn = Nothing


   End If
End Sub

Private Sub lstGrpClinicList_KeyDown(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "lstGrpClinicList_KeyDown"
On Error GoTo debugs
Dim rst3 As New ADODB.Recordset
Dim StrSql As String
        
    If lstGrpClinicList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        
        StrSql = "Select * from GRPClinic Where GRPClnCode = '" & lstGrpClinicList.SelectedItem.Text & "'"
        rst3.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        
        txtClnGrpCode = rst3!GRPClnCode
        txtClnGrpName = rst3!GRPClnName
'        If Len(rst3!GRPClnPaycode) Then
'            cboClnGrpPayor = rst3!GRPClnPaycode
'        End If
        txtClnName = rst3!GRPClnInitial
        If Len(rst3!GRPClnAdd1) Then
            txtClnAdd1 = rst3!GRPClnAdd1
        End If
        If Len(rst3!GRPClnAdd2) Then
            txtClnAdd2 = rst3!GRPClnAdd2
        End If
        If Len(rst3!GRPClnAdd3) Then
            txtClnAdd3 = rst3!GRPClnAdd3
        End If
        If Len(rst3!GRPClnAdmContact) Then
            txtClnAdmContact = rst3!GRPClnAdmContact
        End If
        If Len(rst3!GRPClnContact) Then
            txtClnContact = rst3!GRPClnContact
        End If
        If Len(rst3!GRPClnCEO) Then
            txtClnCEO = rst3!GRPClnCEO
        End If
        If Len(rst3!GRPClnPhone) Then
            txtClnPhone = rst3!GRPClnPhone
        End If
        If Len(rst3!GRPClnFax) Then
            txtClnFax = rst3!GRPClnFax
        End If
        If Len(rst3!GRPClnEmail) Then
            txtClnEmail = rst3!GRPClnEmail
        End If
        If Len(rst3!GRPClnHomepage) Then
            txtClnHomepage = rst3!GRPClnHomepage
        End If
        If Len(rst3!GRPClnEffDate) Then
            txtClnEffDate = rst3!GRPClnEffDate
        End If
        If Len(rst3!GRPClnTax) Then
            cboClnTax = rst3!GRPClnTax
        End If
        If Len(rst3!GRPClnConversion) Then
            cboClnConversion = rst3!GRPClnConversion
        End If
        If Len(rst3!GRPClnPrePost) Then
            cboClnPrePost = rst3!GRPClnPrePost
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' MediConnCISDB.Close
      '  Set MediConnCISDB = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstGrpClinicList_KeyUp(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "lstGrpClinicList_KeyUp"
On Error GoTo debugs
Dim rst3 As New ADODB.Recordset
Dim StrSql As String
        
    If lstGrpClinicList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        
       ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        
        StrSql = "Select * from GRPClinic Where GRPClnCode = '" & lstGrpClinicList.SelectedItem.Text & "'"
        rst3.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        
        txtClnGrpCode = rst3!GRPClnCode
        txtClnGrpName = rst3!GRPClnName
        txtClnName = rst3!GRPClnInitial
        If Len(rst3!GRPClnAdd1) Then
            txtClnAdd1 = rst3!GRPClnAdd1
        End If
        If Len(rst3!GRPClnAdd2) Then
            txtClnAdd2 = rst3!GRPClnAdd2
        End If
        If Len(rst3!GRPClnAdd3) Then
            txtClnAdd3 = rst3!GRPClnAdd3
        End If
        If Len(rst3!GRPClnAdmContact) Then
            txtClnAdmContact = rst3!GRPClnAdmContact
        End If
        If Len(rst3!GRPClnContact) Then
            txtClnContact = rst3!GRPClnContact
        End If
        If Len(rst3!GRPClnCEO) Then
            txtClnCEO = rst3!GRPClnCEO
        End If
        If Len(rst3!GRPClnPhone) Then
            txtClnPhone = rst3!GRPClnPhone
        End If
        If Len(rst3!GRPClnFax) Then
            txtClnFax = rst3!GRPClnFax
        End If
        If Len(rst3!GRPClnEmail) Then
            txtClnEmail = rst3!GRPClnEmail
        End If
        If Len(rst3!GRPClnHomepage) Then
            txtClnHomepage = rst3!GRPClnHomepage
        End If
        If Len(rst3!GRPClnEffDate) Then
            txtClnEffDate = rst3!GRPClnEffDate
        End If
        If Len(rst3!GRPClnTax) Then
            cboClnTax = rst3!GRPClnTax
        End If
        If Len(rst3!GRPClnConversion) Then
            cboClnConversion = rst3!GRPClnConversion
        End If
        If Len(rst3!GRPClnPrePost) Then
            cboClnPrePost = rst3!GRPClnPrePost
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
    
     '   MediConnCISDB.Close
      '  Set MediConnCISDB = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

'Private Sub lstGrpCoList_Click()
'Dim rst3 As New ADODB.Recordset
'Dim strsql As String
        
    'If lstGrpCoList.ListItems.Count <> 0 Then
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
    
        'strsql = "Select * from GRPCompany Where GRPCoCode = '" & lstGrpCoList.SelectedItem.Text & "'"
        'rst3.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        'txtGrpCoCode = rst3!GRPCoCode
        'txtCompanyName = rst3!GrpCoName
        
        'If Len(rst3!GRPCoPAYCode) Then
            'CboCoGrpPayor = rst3!GRPCoPAYCode
        'End If

        'If Len(rst3!GRPCoAdd1) Then
            'txtCoAdd1 = rst3!GRPCoAdd1
        'End If
        'If Len(rst3!GRPCoAdd2) Then
            'txtCoAdd2 = rst3!GRPCoAdd2
        'End If
        'If Len(rst3!GRPCoAdd3) Then
            'txtCoAdd3 = rst3!GRPCoAdd3
        'End If
        'If Len(rst3!GRPCoAdmContact) Then
            'txtCoAdmContact = rst3!GRPCoAdmContact
        'End If
        'If Len(rst3!GRPCoContact) Then
            'txtCoContact = rst3!GRPCoContact
        'End If
        'If Len(rst3!GRPCoCEO) Then
            'txtCoCEO = rst3!GRPCoCEO
        'End If
        'If Len(rst3!GRPCoPhone) Then
            'txtCoPhone = rst3!GRPCoPhone
        'End If
        'If Len(rst3!GRPCoFax) Then
            'txtCoFax = rst3!GRPCoFax
        'End If
        'If Len(rst3!GRPCoEmail) Then
            'txtCoEmail = rst3!GRPCoEmail
        'End If
        'If Len(rst3!GRPCoHomepage) Then
            'txtCoHomepage = rst3!GRPCoHomepage
        'End If
        'If Len(rst3!GRPCoEffDate) Then
            'txtCoEffDate = rst3!GRPCoEffDate
        'End If
        'If Len(rst3!GRPCoTax) Then
            'cboCoTax = rst3!GRPCoTax
        'End If
        'If Len(rst3!GRPCoConversion) Then
            'cboCoConversion = rst3!GRPCoConversion
        'End If
        'If Len(rst3!GRPCoPrePost) Then
            'cboCoPrePost = rst3!GRPCoPrePost
        'End If
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
'End Sub

'Private Sub lstGrpCoList_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim rst3 As New ADODB.Recordset
'Dim strsql As String
        
    'If lstGrpCoList.ListItems.Count <> 0 Then
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
    
        'strsql = "Select * from GRPCompany Where GRPCoCode = '" & lstGrpCoList.SelectedItem.Text & "'"
        'rst3.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        
        'txtGrpCoCode = rst3!GRPCoCode
        'txtCompanyName = rst3!GrpCoName
        
        'If Len(rst3!GRPCoPAYCode) Then
            'CboCoGrpPayor = rst3!GRPCoPAYCode
        'End If

        'If Len(rst3!GRPCoAdd1) Then
            'txtCoAdd1 = rst3!GRPCoAdd1
        'End If
        'If Len(rst3!GRPCoAdd2) Then
            'txtCoAdd2 = rst3!GRPCoAdd2
        'End If
        'If Len(rst3!GRPCoAdd3) Then
            'txtCoAdd3 = rst3!GRPCoAdd3
        'End If
        'If Len(rst3!GRPCoAdmContact) Then
            'txtCoAdmContact = rst3!GRPCoAdmContact
        'End If
        'If Len(rst3!GRPCoContact) Then
            'txtCoContact = rst3!GRPCoContact
        'End If
        'If Len(rst3!GRPCoCEO) Then
            'txtCoCEO = rst3!GRPCoCEO
        'End If
        'If Len(rst3!GRPCoPhone) Then
            'txtCoPhone = rst3!GRPCoPhone
        'End If
        'If Len(rst3!GRPCoFax) Then
            'txtCoFax = rst3!GRPCoFax
        'End If
        'If Len(rst3!GRPCoEmail) Then
            'txtCoEmail = rst3!GRPCoEmail
        'End If
        'If Len(rst3!GRPCoHomepage) Then
            'txtCoHomepage = rst3!GRPCoHomepage
        'End If
        'If Len(rst3!GRPCoEffDate) Then
            'txtCoEffDate = rst3!GRPCoEffDate
        'End If
        'If Len(rst3!GRPCoTax) Then
            'cboCoTax = rst3!GRPCoTax
        'End If
        'If Len(rst3!GRPCoConversion) Then
            'cboCoConversion = rst3!GRPCoConversion
        'End If
        'If Len(rst3!GRPCoPrePost) Then
            'cboCoPrePost = rst3!GRPCoPrePost
        'End If
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
'End Sub

'Private Sub lstGrpCoList_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim rst3 As New ADODB.Recordset
'Dim strsql As String
        
    'If lstGrpCoList.ListItems.Count <> 0 Then
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
    
        'strsql = "Select * from GRPCompany Where GRPCoCode = '" & lstGrpCoList.SelectedItem.Text & "'"
        'rst3.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        'txtGrpCoCode = rst3!GRPCoCode
        'txtCompanyName = rst3!GrpCoName
        
                
        'If Len(rst3!GRPCoPAYCode) Then
            'CboCoGrpPayor = rst3!GRPCoPAYCode
        'End If
        
        'If Len(rst3!GRPCoAdd1) Then
            'txtCoAdd1 = rst3!GRPCoAdd1
        'End If
        'If Len(rst3!GRPCoAdd2) Then
            'txtCoAdd2 = rst3!GRPCoAdd2
        'End If
        'If Len(rst3!GRPCoAdd3) Then
            'txtCoAdd3 = rst3!GRPCoAdd3
        'End If
        'If Len(rst3!GRPCoAdmContact) Then
            'txtCoAdmContact = rst3!GRPCoAdmContact
        'End If
        'If Len(rst3!GRPCoContact) Then
            'txtCoContact = rst3!GRPCoContact
        'End If
        'If Len(rst3!GRPCoCEO) Then
            'txtCoCEO = rst3!GRPCoCEO
        'End If
        'If Len(rst3!GRPCoPhone) Then
            'txtCoPhone = rst3!GRPCoPhone
        'End If
        'If Len(rst3!GRPCoFax) Then
            'txtCoFax = rst3!GRPCoFax
        'End If
        'If Len(rst3!GRPCoEmail) Then
            'txtCoEmail = rst3!GRPCoEmail
        'End If
        'If Len(rst3!GRPCoHomepage) Then
            'txtCoHomepage = rst3!GRPCoHomepage
        'End If
        'If Len(rst3!GRPCoEffDate) Then
            'txtCoEffDate = rst3!GRPCoEffDate
        'End If
        'If Len(rst3!GRPCoTax) Then
            'cboCoTax = rst3!GRPCoTax
        'End If
        'If Len(rst3!GRPCoConversion) Then
            'cboCoConversion = rst3!GRPCoConversion
        'End If
        'If Len(rst3!GRPCoPrePost) Then
            'cboCoPrePost = rst3!GRPCoPrePost
        'End If
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
'End Sub

'Private Sub SearchGrpCompany()
'Dim strAppend As String
   
    'Screen.MousePointer = vbHourglass
    'strAppend = ""
    
    'If Len(Trim(txtGrpCoCode)) Then
        'strAppend = strAppend + " And GRPCoCode like '%" & ChkStr(txtGrpCoCode) & "%'"
    'End If
    
    'If Len(Trim(txtCompanyName)) Then
        'strAppend = strAppend + " And GRPCoName like '%" & ChkStr(txtCompanyName) & "%'"
    'End If
    
    'GRPCoCriteria = strAppend
    'Call GRPCompanyListing(strAppend)
    'Screen.MousePointer = Default
'End Sub

'Private Sub lstGrpCoList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

'lstGrpCoList.SortKey = ColumnHeader.Index - 1
'lstGrpCoList.Sorted = True
    'If lstGrpCoList.SortOrder = lvwAscending Then
        'lstGrpCoList.SortOrder = lvwDescending
    'Else
        'lstGrpCoList.SortOrder = lvwAscending
    'End If
'End Sub

'Private Sub PrintGRPCO()
    'RptGrpCompanyList.Show
'End Sub
'**********************************End Group Company**************************************

'**********************************Start Group Clinic**************************************
Private Sub SaveClinicGrp()
Const METHOD_NAME = "SaveClinicGrp"
On Error GoTo debugs
Dim RecordSaved
Dim ClinicGrp As New ADODB.Recordset
Dim SaveClinicGrp As New ADODB.Recordset
Dim StrSql As String
Dim CLNSeq As String

    If txtClnName = "" Then
        MsgBox "Please Enter Clinic Group Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtClnName.SetFocus
    ElseIf txtClnGrpName = "" Then
        MsgBox "Please Enter Clinic Group Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtClnGrpName.SetFocus
    ElseIf txtClnAdd1 = "" Then
        MsgBox "Please Enter Clinic Group Address", vbOKOnly + vbCritical, DialogBoxTitle
        txtClnAdd1.SetFocus
    ElseIf txtClnEffDate = "__/__/____" Then
        MsgBox "Please Enter Effective Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtClnEffDate.SetFocus
    Else
        StrSql = "Select * From GRPClinic Where GRPClnCode = '" & Trim(txtClnGrpCode) & "'"
        ClinicGrp.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
      If ClinicGrp.RecordCount > 0 And EditFlag = False Then
      MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
      ClinicGrp.Close
      txtClnGrpCode.SetFocus
      txtClnGrpCode.SelStart = 0
      txtClnGrpCode.SelLength = Len(txtClnGrpCode)
      Set ClinicGrp = Nothing
      Exit Sub
      Else
      ClinicGrp.Close
      Set ClinicGrp = Nothing
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            If EditFlag = False Then
                ClinicGrp.Open "GRPClinic", MediConnCISDB, adOpenKeyset, adLockOptimistic
                ClinicGrp.AddNew
                CLNSeq = GenerateGRPClnSeq(Mid(txtClnName, 1, 1))
            Else
                StrSql = "Select * From GRPClinic Where GRPClnCode = '" & Trim(txtClnGrpCode) & "'"
                ClinicGrp.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                CLNSeq = txtClnGrpCode
            End If
   
                With ClinicGrp
                    !GRPClnCode = ChkStr(CLNSeq)
                    !GRPClnName = ChkStr(txtClnGrpName)
                    !GRPClnInitial = ChkStr(txtClnName)
                    '!GRPClnPaycode = Mid(cboClnGrpPayor, 1, 2)
                    !GRPClnAdd1 = IIf(Len(txtClnAdd1), ChkStr(txtClnAdd1), Null)
                    !GRPClnAdd2 = IIf(Len(txtClnAdd2), ChkStr(txtClnAdd2), Null)
                    !GRPClnAdd3 = IIf(Len(txtClnAdd3), ChkStr(txtClnAdd3), Null)
                    !GRPClnContact = IIf(Len(txtClnContact), ChkStr(txtClnContact), Null)
                    !GRPClnAdmContact = IIf(Len(txtClnAdmContact), ChkStr(txtClnAdmContact), Null)
                    !GRPClnCEO = IIf(Len(txtClnCEO), ChkStr(txtClnCEO), Null)
                    !GRPClnPhone = IIf(Len(txtClnPhone), ChkStr(txtClnPhone), Null)
                    !GRPClnFax = IIf(Len(txtClnFax), ChkStr(txtClnFax), Null)
                    !GRPClnEmail = IIf(Len(txtClnEmail), ChkStr(txtClnEmail), Null)
                    !GRPClnHomepage = ChkStr(txtClnHomepage)
                    !GRPClnEffDate = IIf(IsDate(txtClnEffDate), txtClnEffDate, Null)
                    !GRPClnTax = Mid(cboClnTax, 1, 1)
                    !GRPClnConversion = Mid(cboClnConversion, 1, 1)
                    !GRPClnPrePost = Mid(cboClnPrePost, 1, 1)
                    !GRPClnLastUpdateUser = Logon
                    !GRPClnLastUpdateDate = Date
                End With
                ClinicGrp.update
                Call ClearForm(Me)
                Call GrpClinicListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
     
            ClinicGrp.Close
            Set SaveClinicGrp = Nothing
        End If
        EditFlag = False
    End If
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub DeleteClinicGrp()
Const METHOD_NAME = "DeleteClinicGrp"
On Error GoTo debugs
Dim DeleteGrpCln As New ADODB.Recordset
Dim rsChkCln As New ADODB.Recordset
Dim StrSql, sqlClinic As String
Dim DeleteRecord

EditFlag = False
If txtClnGrpCode = "" Then
   MsgBox "Please Enter Clinic Group Code", vbOKOnly + vbCritical, DialogBoxTitle
'   txtClnGrpCode.SetFocus
ElseIf lstGrpClinicList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlClinic = "Select GRPClnCode from GRPClinic where GRPClnCode = '" & RemStr(txtClnGrpCode) & "'"

   rsChkCln.MaxRecords = 1
   rsChkCln.Open sqlClinic, MediConnCISDB, adOpenKeyset, adLockOptimistic

   If rsChkCln.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in Clinic Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsChkCln.RecordCount = 0 Then
           StrSql = "Delete From ClinicGrp Where ClinicGrpCode = '" & ChkStr(lstGrpClinicList.SelectedItem.Text) & "'"
           DeleteGrpCln.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call GrpClinicListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtClnGrpCode.SetFocus
         End If
     End If
End If
Exit Sub

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub GrpClinicListing(Optional strAppend As String)
Const METHOD_NAME = "GrpClinicListing"
On Error GoTo debugs
    Dim rst2 As New ADODB.Recordset
    Dim ClinicGrpMaster As ListItem
    Dim sql As String
    Dim TotalGRPCLNCount As Integer
    
    sql = "Select * from GRPClinic where 0=0 "
   
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
              
    TotalGRPCLNCount = 0
    lstGrpClinicList.ListItems.Clear
    rst2.Open sql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set ClinicGrpMaster = lstGrpClinicList.ListItems.Add(, , rst2!GRPClnCode)
                ClinicGrpMaster.SubItems(1) = rst2!GRPClnName
                rst2.MoveNext
        TotalGRPCLNCount = TotalGRPCLNCount + 1
        txtTotalRecord = TotalGRPCLNCount
        Loop
    rst2.Close
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub GRPClinicComboListing()
Dim PAY As New ADODB.Recordset
        
'    PAY.Open "PAY", MediConnCISDB, adOpenKeyset, adLockOptimistic
    
'    cboClnGrpPayor.Clear

'    Do Until PAY.EOF
'        cboClnGrpPayor.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
'        PAY.MoveNext
'    Loop
            
'    PAY.Close
    
    cboClnConversion.Clear
    cboClnConversion.AddItem "N" & " - " & "No"
    cboClnConversion.AddItem "Y" & " - " & "Yes"
    
    cboClnTax.Clear
    cboClnTax.AddItem "N" & " - " & "No"
    cboClnTax.AddItem "Y" & " - " & "Yes"
    
    cboClnPrePost.Clear
    cboClnPrePost.AddItem "N" & " - " & "No"
    cboClnPrePost.AddItem "Y" & " - " & "Yes"
    
End Sub

Private Sub lstGrpClinicList_Click()
Const METHOD_NAME = "lstGrpClinicList_Click"
On Error GoTo debugs
Dim rst3 As New ADODB.Recordset
Dim StrSql As String
        
    If lstGrpClinicList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
    
       ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        
        StrSql = "Select * from GRPClinic Where GRPClnCode = '" & lstGrpClinicList.SelectedItem.Text & "'"
        rst3.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        
        txtClnGrpCode = rst3!GRPClnCode
        txtClnGrpName = rst3!GRPClnName
'        If Len(rst3!GRPClnPaycode) Then
'            cboClnGrpPayor = rst3!GRPClnPaycode
'        End If
        If Len(rst3!GRPClnInitial) Then
            txtClnName = rst3!GRPClnInitial
        End If
        If Len(rst3!GRPClnAdd1) Then
            txtClnAdd1 = rst3!GRPClnAdd1
        End If
        If Len(rst3!GRPClnAdd2) Then
            txtClnAdd2 = rst3!GRPClnAdd2
        End If
        If Len(rst3!GRPClnAdd3) Then
            txtClnAdd3 = rst3!GRPClnAdd3
        End If
        If Len(rst3!GRPClnAdmContact) Then
            txtClnAdmContact = rst3!GRPClnAdmContact
        End If
        If Len(rst3!GRPClnContact) Then
            txtClnContact = rst3!GRPClnContact
        End If
        If Len(rst3!GRPClnCEO) Then
            txtClnCEO = rst3!GRPClnCEO
        End If
        If Len(rst3!GRPClnPhone) Then
            txtClnPhone = rst3!GRPClnPhone
        End If
        If Len(rst3!GRPClnFax) Then
            txtClnFax = rst3!GRPClnFax
        End If
        If Len(rst3!GRPClnEmail) Then
            txtClnEmail = rst3!GRPClnEmail
        End If
        If Len(rst3!GRPClnHomepage) Then
            txtClnHomepage = rst3!GRPClnHomepage
        End If
        If Len(rst3!GRPClnEffDate) Then
            txtClnEffDate = rst3!GRPClnEffDate
        End If
        If Len(rst3!GRPClnTax) Then
            cboClnTax = rst3!GRPClnTax
        End If
        If Len(rst3!GRPClnConversion) Then
            cboClnConversion = rst3!GRPClnConversion
        End If
        If Len(rst3!GRPClnPrePost) Then
            cboClnPrePost = rst3!GRPClnPrePost
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
      '  MediConnCISDB.Close
      '  Set MediConnCISDB = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SearchClinicGrp()
Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(txtClnGrpCode)) Then
        strAppend = strAppend + " And GRPClnCode like '%" & ChkStr(txtClnGrpCode) & "%'"
    End If
    If Len(Trim(txtClnGrpName)) Then
        strAppend = strAppend + " And GRPClnName like '%" & ChkStr(txtClnGrpName) & "%'"
    End If
    
    GRPClnCriteria = strAppend
    Call GrpClinicListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstGrpClinicList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

lstGrpClinicList.SortKey = ColumnHeader.Index - 1
lstGrpClinicList.Sorted = True
    If lstGrpClinicList.SortOrder = lvwAscending Then
        lstGrpClinicList.SortOrder = lvwDescending
    Else
        lstGrpClinicList.SortOrder = lvwAscending
    End If
End Sub

Private Sub PrintGRPCLN()
    RptGrpClinicList.Show
End Sub
'**********************************End Group Clinic**************************************

'**********************************Start Doctor**************************************
Private Sub SaveDoctor()
Const METHOD_NAME = "SaveDoctor"
On Error GoTo debugs
Dim RecordSaved
Dim Dr As New ADODB.Recordset
Dim SaveDoc As New ADODB.Recordset
Dim StrSql As String
Dim DocSeq As String
        
    If txtDocName = "" Then
        MsgBox "Please Enter Doctor Full Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtDocName.SetFocus
    Else
        StrSql = "Select * From Doctor Where DocName = '" & Trim(txtDocName) & "'"
        Dr.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
      If Dr.RecordCount > 0 And EditFlag = False Then
      MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
      Dr.Close
      txtDocCode.SetFocus
      txtDocCode.SelStart = 0
      txtDocCode.SelLength = Len(txtDocCode)
      Set Dr = Nothing
      Exit Sub
      Else
      Dr.Close
      Set Dr = Nothing
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            If EditFlag = False Then
                Dr.Open "Doctor", MediConnCISDB, adOpenKeyset, adLockOptimistic
                Dr.AddNew
                DocSeq = GenerateDocSeq(Mid(txtDocName, 1, 1))
            Else
                StrSql = "Select * From Doctor Where DocCode = '" & Trim(txtDocCode) & "'"
                Dr.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                DocSeq = txtDocCode
            End If
            
                With Dr
                    !DocCode = ChkStr(DocSeq)
                    !DocName = IIf(Len(txtDocName), ChkStr(txtDocName), Null)
                    '!DocInitial = IIf(Len(txtDocInitial), ChkStr(txtDocInitial), Null)
                    !DocLastUpdateUser = Logon
                    !DocLastUpdateDate = Date
                End With
                Dr.update
                Call ClearForm(Me)
                Call DoctorListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                
            Dr.Close
            Set SaveDoc = Nothing
        End If
        EditFlag = False
    End If
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub DeleteDoctor()
Const METHOD_NAME = "DeleteDoctor"
On Error GoTo debugs
Dim DeleteDr As New ADODB.Recordset
Dim rsChkDr As New ADODB.Recordset
Dim StrSql, sqlDoctor As String
Dim DeleteRecord

EditFlag = False
If txtDocCode = "" Then
   MsgBox "Please Enter Doctor Code", vbOKOnly + vbCritical, DialogBoxTitle
   txtDocCode.SetFocus
ElseIf lstDoctorList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlDoctor = "Select DocCode from Doctor where DocCode = '" & RemStr(txtDocCode) & "'"
   
   rsChkDr.MaxRecords = 1
   rsChkDr.Open sqlDoctor, MediConnCISDB, adOpenKeyset, adLockOptimistic
   
   If rsChkDr.RecordCount Then
        MsgBox "This record cannot be deleted because it is used in Doctor Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsChkDr.RecordCount = 0 Then
           StrSql = "Delete From Doctor Where DocCode = '" & ChkStr(lstDoctorList.SelectedItem.Text) & "'"
           DeleteDr.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call DoctorListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtDocCode.SetFocus
         End If
     End If
End If
Exit Sub
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub DoctorListing(Optional strAppend As String)
Const METHOD_NAME = "DoctorListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim DrMaster As ListItem
Dim sql As String
Dim TotalDRCount As Integer
    
    sql = "Select * from Doctor where 0=0 "
   
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
              
    TotalDRCount = 0
    lstDoctorList.ListItems.Clear
    rst2.Open sql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set DrMaster = lstDoctorList.ListItems.Add(, , rst2!DocCode)
                DrMaster.SubItems(1) = rst2!DocName
                'DrMaster.SubItems(2) = rst2!GrpCoPaycode
                rst2.MoveNext
        TotalDRCount = TotalDRCount + 1
        txtTotalRecord = TotalDRCount
        Loop
    rst2.Close
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ClinicAccessListing()
Const METHOD_NAME = "ClinicAccessListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim ClnAccessMaster As ListItem
Dim sql As String
Dim TotalCACount As Integer
    
    sql = "Select TOP 100 ClinicCode, CLNName, EffDateFrom, EffDateTo, CreatedBy, CreatedDate from CLNAccessControl INNER JOIN CLN ON CLN.CLNCode = ClinicCode Order By CreatedDate DESC"
              
    TotalCACount = 0
    lstClinicAccess.ListItems.Clear
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set ClnAccessMaster = lstClinicAccess.ListItems.Add(, , rst2!ClinicCode)
                ClnAccessMaster.SubItems(1) = rst2!CLNName
                ClnAccessMaster.SubItems(2) = rst2!EffDateFrom
                ClnAccessMaster.SubItems(3) = rst2!EffDateTo
                ClnAccessMaster.SubItems(4) = rst2!CreatedBy
                ClnAccessMaster.SubItems(5) = rst2!CreatedDate
                rst2.MoveNext
        TotalCACount = TotalCACount + 1
        txtTotalRecord = TotalCACount
        Loop
    rst2.Close
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

'Private Sub DoctorComboListing()
'Dim PAY As New ADODB.Recordset
        
'    PAY.Open "PAY", MediConnCISDB, adOpenKeyset, adLockOptimistic
    
'    CboCoGrpPayor.Clear

'    Do Until PAY.EOF
'        CboCoGrpPayor.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
'        PAY.MoveNext
'    Loop
            
'    PAY.Close
       
'    cboCoPrePost.Clear
'    cboCoPrePost.AddItem "N" & " - " & "No"
'    cboCoPrePost.AddItem "Y" & " - " & "Yes"
    
'End Sub


Private Sub lstDoctorList_Click()
Const METHOD_NAME = "lstDoctorList_Click"
On Error GoTo debugs
Dim rst3 As New ADODB.Recordset
Dim StrSql As String
        
    If lstDoctorList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            StrSql = "Select * from Doctor Where DocCode = '" & lstDoctorList.SelectedItem.Text & "'"
            rst3.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            
            txtDocCode = rst3!DocCode
            txtDocName = rst3!DocName
            Call EntriesEnable(False)
            cmdSave.Enabled = False
            rst3.Close
            Set rst3 = Nothing
        
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
        
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstDoctorList_KeyDown(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "lstDoctorList_KeyDown"
On Error GoTo debugs
Dim rst3 As New ADODB.Recordset
Dim StrSql As String
        
    If lstDoctorList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            StrSql = "Select * from Doctor Where DocCode = '" & lstDoctorList.SelectedItem.Text & "'"
            rst3.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            
            txtDocCode = rst3!DocCode
            txtDocName = rst3!DocName
            Call EntriesEnable(False)
            cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstDoctorList_KeyUp(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "lstDoctorList_KeyUp"
On Error GoTo debugs
Dim rst3 As New ADODB.Recordset
Dim StrSql As String
        
    If lstDoctorList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            StrSql = "Select * from Doctor Where DocCode = '" & lstDoctorList.SelectedItem.Text & "'"
            rst3.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            
            txtDocCode = rst3!DocCode
            txtDocName = rst3!DocName
            Call EntriesEnable(False)
            cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub SearchDoctor()

Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(txtDocCode)) Then
        strAppend = strAppend + " And DocCode like '%" & ChkStr(txtDocCode) & "%'"
    End If
    
    If Len(Trim(txtDocName)) Then
        strAppend = strAppend + " And DocName like '%" & ChkStr(txtDocName) & "%'"
    End If
    
    DOCCriteria = strAppend
    Call DoctorListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstDoctorList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

lstDoctorList.SortKey = ColumnHeader.Index - 1
lstDoctorList.Sorted = True
    If lstDoctorList.SortOrder = lvwAscending Then
        lstDoctorList.SortOrder = lvwDescending
    Else
        lstDoctorList.SortOrder = lvwAscending
    End If
End Sub

Private Sub PrintDoctor()
    RptDoctorList.Show
End Sub
'**********************************End Doctor**************************************

'**********************************Start Cln Info**************************************
Private Sub SaveClinic()
Const METHOD_NAME = "SaveClinic"
On Error GoTo debugs
'On Error Resume Next
Dim RecordSaved
Dim CLN As New ADODB.Recordset
Dim SaveCLN As New ADODB.Recordset
Dim City As New ADODB.Recordset
Dim state As New ADODB.Recordset
Dim StrSql As String
Dim str2 As String
Dim str3 As String
Dim CLNSeq As String
Dim strAppend As String
Dim CityName As String
Dim StateName As String
Dim CLNName As String
Dim DocName As String
Dim CLNinitial As String
Dim contact As String

'If txtCLNCode = "" Then
'    MsgBox "Please Enter Clinic Code", vbOKOnly + vbCritical, DialogBoxTitle
'    txtCLNCode.SetFocus

If txtInfoClnName = "" Then
    MsgBox "Please Enter Clinic Name", vbOKOnly + vbCritical, DialogBoxTitle
    txtInfoClnName.SetFocus
ElseIf txtInfoClnMIX = "" Then
    MsgBox "Please Enter Clinic Initial", vbOKOnly + vbCritical, DialogBoxTitle
    txtInfoClnMIX.SetFocus
ElseIf cboInfoClnType = "" Then
    MsgBox "Please Enter Clinic Type", vbOKOnly + vbCritical, DialogBoxTitle
    cboInfoClnType.SetFocus
ElseIf txtInfoClnAdd1 = "" Then
    MsgBox "Please Enter Clinic Address", vbOKOnly + vbCritical, DialogBoxTitle
    txtInfoClnAdd1.SetFocus
ElseIf cboInfoClnBNFType = "" Then
    MsgBox " Please Enter Benefit Type", vbOKOnly + vbCritical, DialogBoxTitle
    cboInfoClnBNFType.SetFocus
ElseIf txtInfoClnReceivedDate = "__/__/____" Or IsDate(txtInfoClnReceivedDate) = False Then
    MsgBox "Please Enter Received Date!", vbOKOnly + vbCritical, DialogBoxTitle
    txtInfoClnReceivedDate.SetFocus
ElseIf cboInfoClnState = "" Then
    MsgBox " Please Enter State", vbOKOnly + vbCritical, DialogBoxTitle
    cboInfoClnState.SetFocus
ElseIf txtInfoClnPostCode = "" Then
    MsgBox " Please Enter PostCode", vbOKOnly + vbCritical, DialogBoxTitle
    txtInfoClnPostCode.SetFocus
ElseIf cboInfoClnCity = "" Then
    MsgBox " Please Enter City", vbOKOnly + vbCritical, DialogBoxTitle
    cboInfoClnCity.SetFocus
ElseIf cboInfoClnPanel = "" Then
    MsgBox " Please Enter Panel", vbOKOnly + vbCritical, DialogBoxTitle
    cboInfoClnPanel.SetFocus
ElseIf cboInfoCln24hrs = "" Then
    MsgBox " Please Enter Clinic Operating Hours", vbOKOnly + vbCritical, DialogBoxTitle
    cboInfoCln24hrs.SetFocus
ElseIf cboInfoFFS = "" Then
    MsgBox " Please Enter FFS Status", vbOKOnly + vbCritical, DialogBoxTitle
    cboInfoFFS.SetFocus
ElseIf cboInfoFFS = "Y" And (txtFFSPerc = "" Or txtFFSPerc = "0") Then
    MsgBox " Please Enter FFS Percentage", vbOKOnly + vbCritical, DialogBoxTitle
    txtFFSPerc.SetFocus
    
ElseIf TxtLocation = "" Then
    MsgBox " Please Enter Location", vbOKOnly + vbCritical, DialogBoxTitle
    TxtLocation.SetFocus
ElseIf txtOperatinghr = "" Then
    MsgBox " Please Operating hours", vbOKOnly + vbCritical, DialogBoxTitle
    txtOperatinghr.SetFocus
Else
      StrSql = "Select * From CLN Where CLNAdd1 = '" & txtInfoClnAdd1 & "'"
      CLN.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
      
      If CLN.RecordCount > 0 And EditFlag = False Then
          MsgBox "Duplicate Record!", vbCritical, "Error"
          CLN.Close
          Set CLN = Nothing
    '      txtCLNCode.SetFocus
    '      txtCLNCode.SelStart = 0
    '      txtCLNCode.SelLength = Len(txtCLNCode)
    '      Set CLN = Nothing
          Exit Sub
        Else
        
            CLN.Close
            Set CLN = Nothing

        End If
'      Else

'        str2 = "Select * from CTY Where CTYCode = '" & Mid(cboCLNcity, 1, 3) & "'"
'        City.Open str2, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
       
'        CityName = City!CTYDescription
        
'        City.Close
'        Set City = Nothing
      
'        str3 = "Select * from STA Where STACode = '" & Mid(cboCLNstate, 1, 2) & "'"
'        State.Open str3, medixmasterconn, adOpenKeyset, adLockOptimistic
        
'        StateName = State!STADescription
        
'        State.Close
'        Set State = Nothing
      
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            If EditFlag = False Then
                CLN.Open "CLN", medixmasterconn, adOpenKeyset, adLockOptimistic
                CLN.AddNew
                CLNSeq = GenerateCLNSeq(Mid(txtInfoClnMIX, 1, 1))
            Else
                StrSql = "Select * From CLN Where CLNCode = '" & Trim(txtInfoCLNCode) & "'"
                CLN.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                CLNSeq = txtInfoCLNCode
            End If
            
                With CLN
                    'CLNSeq = Mid(txtMIXClinic, 1, 1)
                    !CLNCode = ChkStr(CLNSeq)
                    'CLNname = Replace(txtCLNName, "'", "''")
                    !CLNName = ChkStr(txtInfoClnName)
                    !ClinicGrpCode = Trim(Mid(CboInfoClnGrp, 1, 5))
                    !CLNPhone = ChkStr(txtInfoClnTelNo)
                    !CLNFax = ChkStr(txtInfoClnFaxNo)
                    !CLNAdd1 = ChkStr(txtInfoClnAdd1)
                    !CLNAdd2 = ChkStr(txtInfoClnAdd2)
                    !CLNAdd3 = ChkStr(txtInfoClnAdd3)
                    !CLNEMAIL = ChkStr(txtInfoClnEmail)
                    !CLNHomepage = ChkStr(txtInfoClnHome)
                    'contact = Replace(TxtCLNContact, "'", "''")
                    !CLNRepresentative = ChkStr(txtInfoClnContact)
                    'DocName = Replace(TxtCLNDocName, "'", "''")
                    !CLNDocName = ChkStr(txtInfoClnDocName)
                    !CLNBNFType = Trim(Mid(cboInfoClnBNFType, 1, 1))
                    !CLNForRef = ChkStr(txtInfoClnFormema)
                    If txtInfoClnNoDoc <> "" Then
                        !CLNNoofDoc = txtInfoClnNoDoc
                    End If
                    !CLNType = Trim(Mid(cboInfoClnType, 1, 1))
                    !CLNXRay = Trim(Mid(cboInfoClnXray, 1, 1))
                    !CLNPanel = Trim(Mid(cboInfoClnPanel, 1, 1))
                    !CLN24Hrs = Trim(Mid(cboInfoCln24hrs, 1, 1))
                    !CLNFFSStatus = Trim(Mid(cboInfoFFS, 1, 1))
                    !CLNMCOProcfeePerc = txtFFSPerc
                    !CLNState = ChkStr(cboInfoClnState)
                    !CLNCity = ChkStr(cboInfoClnCity)
                    !CLNPostCode = ChkStr(txtInfoClnPostCode)
                    'CLNinitial = Replace(txtMIXClinic, "'", "''")
                    !CLNinitial = ChkStr(txtInfoClnMIX)
                    !CLNReceivedDate = txtInfoClnReceivedDate
                    !CLNBankAcc = ChkStr(txtInfoClnBankAcc)
                    !CLNBankName = ChkStr(txtInfoClnBank)
                    !CLNBankBrn = ChkStr(txtInfoClnBranch)
                    !CLNAccPerson = ChkStr(txtInfoClnAccPerson)
                    !CLNRemark = ChkStr(txtInfoClnRemark)
                    !CLNLocation = ChkStr(TxtLocation)
                    !CLNLastUpdateUser = Logon
                    !CLNLastUpdateDate = Date
                    !CLNOperatingHr = ChkStr(txtOperatinghr)
                    
                    !CLNPName = ChkStr(txtPymtInfoClnName)
                    '!CLNPPostCode = Trim(txtPymtInfoClnPostCode)
                    !CLNPAdd1 = ChkStr(txtPymtInfoClnAdd1)
                    !CLNPAdd2 = ChkStr(txtPymtInfoClnAdd2)
                    !CLNPAdd3 = ChkStr(txtPymtInfoClnAdd3)
                    '!CLNPState = ChkStr(cboPymtInfoClnState)
                    '!CLNPcity = ChkStr(cboPymtInfoClnCity)
                    !CLNOnlineStatus = IIf(Len(cboOnlineStatus), Mid(cboOnlineStatus, 1, 1), "")
                    !CLNGLFormat = IIf(Len(cboGLFormat), Mid(cboGLFormat, 1, 1), "")
                    !CLNRegNo = IIf(Len(txtRegNo), txtRegNo, "")
                    
                End With
                CLN.update
                'Call ClearForm(Me)
                Call SearchCLN
                'Call CLNListing(strAppend)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
 
            CLN.Close
            Set SaveCLN = Nothing
        End If
        EditFlag = False
    End If
'End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub DeleteClinic()
Const METHOD_NAME = "DeleteClinic"
On Error GoTo debugs

Dim DeleteCLN As New ADODB.Recordset
Dim StrSql As String
Dim DeleteRecord
Dim CASRec As New ADODB.Recordset
Dim CASSql As String
Dim CASFound As Boolean

    EditFlag = False
    CASFound = False
    If Trim(txtInfoCLNCode) <> "" Then
        CASSql = "SELECT CASNumber from CISCAS Where ClinicCode='" & ChkStr(lstClinicInfoList.SelectedItem.Text) & "'"
        CASRec.Open CASSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not CASRec.EOF
            CASFound = True
            CASRec.MoveNext
        Loop
        CASRec.Close
        Set CASRec = Nothing
    End If
    If txtInfoClnName = "" Then
       MsgBox "Please Enter Clinic Name", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf lstClinicInfoList.ListItems.Count = 0 Then
       MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf CASFound = True Then
       MsgBox "There are case(s) register under this Clinic !", vbOKOnly + vbCritical, DialogBoxTitle
    Else
       DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
         If DeleteRecord = vbYes Then
               StrSql = "Delete From CLN Where CLNcode = '" & ChkStr(lstClinicInfoList.SelectedItem.Text) & "'"
               DeleteCLN.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
               MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
               cmdSave.Enabled = True
               Call ClearForm(Me)
               lstClinicInfoList.ListItems.Clear
               Call DisableCmdButton(Me)
               Call EntriesEnable(True)
               txtInfoClnName.SetFocus
         End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SearchCLN()

    Dim strAppend As String
    Dim cliname As String
    Dim clinitial As String
    Dim clstate As String
    Dim clcity As String
    Dim clpost As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    cliname = ""
    clinitial = ""
    clstate = ""
    clcity = ""
    clpost = ""
    
    If Len(Trim(txtInfoCLNCode)) Then
        strAppend = strAppend + " And CLNCode = '" & ChkStr(Trim(txtInfoCLNCode)) & "'"
    End If
    
    If Len(Trim(txtInfoClnMIX)) Then
        clinitial = Replace(txtInfoClnMIX, "'", "''")
        strAppend = strAppend + " And CLNInitial like '%" & ChkStr(clinitial) & "%'"
    End If
    
    If Len(Trim(txtInfoClnName)) Then
        cliname = Replace(txtInfoClnName, "'", "''")
        strAppend = strAppend + " And CLNName like '%" & ChkStr(cliname) & "%'"
    End If
    
    If Len(Trim(cboInfoClnState)) Then
        clstate = Replace(cboInfoClnState, "'", "''")
        strAppend = strAppend + " And CLNState like '%" & ChkStr(clstate) & "%'"
    End If
    
    If Len(Trim(cboInfoClnCity)) Then
        clcity = Replace(cboInfoClnCity, "'", "''")
        strAppend = strAppend + " And CLNCity like '%" & ChkStr(clcity) & "%'"
    End If
    
    If Len(Trim(txtInfoClnPostCode)) Then
        clpost = Replace(txtInfoClnPostCode, "'", "''")
        strAppend = strAppend + " And CLNPostcode like '%" & ChkStr(clpost) & "%'"
    End If
    
    If Len(Trim(cboInfoClnPanel)) Then
        strAppend = strAppend + " AND CLNPanel = '" & RemStr(Mid(cboInfoClnPanel, 1, 1)) & "'"
    End If
    
    If Len(Trim(CboInfoClnGrp)) Then
        strAppend = strAppend + " AND ClinicGrpCode like '" & RemStr(Mid(CboInfoClnGrp, 1, IIf(InStr(1, CboInfoClnGrp, "-") = 0, 4, InStr(1, CboInfoClnGrp, "-") - 1))) & "%'"
    End If
    
    If Len(Trim(cboInfoClnType)) Then
        strAppend = strAppend + " AND CLNType = '" & RemStr(Mid(cboInfoClnType, 1, 1)) & "'"
    End If
    
    If Len(Trim(cboInfoCln24hrs)) Then
        strAppend = strAppend + " AND CLN24Hrs = '" & RemStr(Mid(cboInfoCln24hrs, 1, 1)) & "'"
    End If
    
    If Len(Trim(cboInfoFFS)) Then
        strAppend = strAppend + " AND CLNFFSStatus = '" & RemStr(Mid(cboInfoFFS, 1, 1)) & "'"
    End If
    
    If Len(Trim(cboOnlineStatus)) Then
    
        If RemStr(Mid(cboOnlineStatus, 1, 1)) = "Y" Then
            strAppend = strAppend + " AND CLNOnlineStatus = 'Y'"
        Else
            strAppend = strAppend + " AND (CLNOnlineStatus = '' OR CLNOnlineStatus IS NULL OR CLNOnlineStatus = 'N')"
        End If
    End If
    
    If Len(Trim(cboGLFormat)) Then
    
        If RemStr(Mid(cboGLFormat, 1, 1)) = "M" Then
            strAppend = strAppend + " AND CLNGLFormat = 'M'"
        Else
            strAppend = strAppend + " AND CLNGLFormat = 'H'"
        End If
    End If
    
    
    
    CLNCriteria = strAppend
    Call CLNListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub CLNListing(Optional strAppend As String)
    Const METHOD_NAME = "CLNListing"
On Error GoTo debugs
    Dim rst2 As New ADODB.Recordset
    Dim ClnMaster As ListItem
    Dim sql As String
    Dim TotalCLNCount As Integer
    
    sql = "Select * from CLN where 0=0"
    
    'sql = "Select CLNCode, CLNName, ClinicGrpCode, CLNAdd1, CLNAdd2, CLNAdd3," & _
          "CLNPhone, CLNFAX, CLNEmail, CLNHOMEPage, CLNRepresentative, CLNDocName, CLNBNFType " & _
          "from CLN where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    TotalCLNCount = 0
    lstClinicInfoList.ListItems.Clear
    
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set ClnMaster = lstClinicInfoList.ListItems.Add(, , rst2!CLNCode)
                If Trim(rst2!CLNName) <> " " Then
                    ClnMaster.SubItems(1) = rst2!CLNName
                End If
                
                If Trim(rst2!ClinicGrpCode) <> " " Then
                    ClnMaster.SubItems(2) = rst2!ClinicGrpCode
                End If
                
                If Trim(rst2!CLNinitial) <> " " Then
                    ClnMaster.SubItems(3) = rst2!CLNinitial
                End If
                
                If Trim(rst2!CLNType) <> " " Then
                    ClnMaster.SubItems(4) = rst2!CLNType
                End If
                
                If Trim(rst2!CLNBNFType) <> " " Then
                    ClnMaster.SubItems(5) = rst2!CLNBNFType
                End If
                
                If Trim(rst2!CLNAdd1) <> " " Then
                    ClnMaster.SubItems(6) = rst2!CLNAdd1
                End If
                
                If Trim(rst2!CLNAdd2) <> " " Then
                    ClnMaster.SubItems(7) = rst2!CLNAdd2
                End If
                
                If Trim(rst2!CLNAdd3) <> " " Then
                    ClnMaster.SubItems(8) = rst2!CLNAdd3
                End If
                
                If Trim(rst2!CLNPostCode) <> " " Then
                    ClnMaster.SubItems(9) = rst2!CLNPostCode
                End If
                
                If Trim(rst2!CLNCity) <> " " Then
                    ClnMaster.SubItems(10) = rst2!CLNCity
                End If
                
                If Trim(rst2!CLNState) <> " " Then
                    ClnMaster.SubItems(11) = rst2!CLNState
                End If
                
                If Trim(rst2!CLNReceivedDate) <> " " Then
                    ClnMaster.SubItems(12) = rst2!CLNReceivedDate
                End If
                
                If Trim(rst2!CLNPanel) <> " " Then
                    ClnMaster.SubItems(13) = rst2!CLNPanel
                End If
                
                If Trim(rst2!CLNXRay) <> " " Then
                    ClnMaster.SubItems(14) = rst2!CLNXRay
                End If
                
                If Trim(rst2!CLNForRef) <> " " Then
                    ClnMaster.SubItems(15) = rst2!CLNForRef
                End If
                
                If Trim(rst2!CLNNoofDoc) <> " " Then
                    ClnMaster.SubItems(16) = rst2!CLNNoofDoc
                End If
                
                If Trim(rst2!CLNPhone) <> " " Then
                    ClnMaster.SubItems(17) = rst2!CLNPhone
                End If
                
                If Trim(rst2!CLNFax) <> " " Then
                    ClnMaster.SubItems(18) = rst2!CLNFax
                End If
                
                If Trim(rst2!CLNEMAIL) <> " " Then
                    ClnMaster.SubItems(19) = rst2!CLNEMAIL
                End If
                
                If Trim(rst2!CLNHomepage) <> " " Then
                    ClnMaster.SubItems(20) = rst2!CLNHomepage
                End If
                
                If Trim(rst2!CLNBankAcc) <> " " Then
                    ClnMaster.SubItems(21) = rst2!CLNBankAcc
                End If
                
                If Trim(rst2!CLNBankName) <> " " Then
                    ClnMaster.SubItems(22) = rst2!CLNBankName
                End If
                
                If Trim(rst2!CLNBankBrn) <> " " Then
                    ClnMaster.SubItems(23) = rst2!CLNBankBrn
                End If
                
                If Trim(rst2!CLNAccPerson) <> " " Then
                    ClnMaster.SubItems(24) = rst2!CLNAccPerson
                End If
                
                If Trim(rst2!CLN24Hrs) <> " " Then
                    ClnMaster.SubItems(25) = rst2!CLN24Hrs
                End If
                
                If Trim(rst2!CLNRemark) <> " " Then
                    ClnMaster.SubItems(26) = rst2!CLNRemark
                End If
                
                If Trim(rst2!CLNFFSStatus) <> " " Then
                    ClnMaster.SubItems(27) = rst2!CLNFFSStatus
                End If
                
                If Len(rst2!CLNOperatingHr) Then
                    ClnMaster.SubItems(28) = rst2!CLNOperatingHr
                End If
                
                If Len(rst2!CLNPName) Then
                    ClnMaster.SubItems(29) = rst2!CLNPName
                End If
                
                If Len(rst2!CLNPAdd1) Then
                    ClnMaster.SubItems(30) = rst2!CLNPAdd1
                End If
                
                If Len(rst2!CLNPAdd2) Then
                    ClnMaster.SubItems(31) = rst2!CLNPAdd2
                End If
                
                If Len(rst2!CLNPAdd3) Then
                    ClnMaster.SubItems(32) = rst2!CLNPAdd3
                End If
                
                If Len(rst2!CLNPPostCode) Then
                    ClnMaster.SubItems(33) = rst2!CLNPPostCode
                End If
                
                If Len(rst2!CLNPcity) Then
                    ClnMaster.SubItems(34) = rst2!CLNPcity
                End If
                
                If Len(rst2!CLNPState) Then
                    ClnMaster.SubItems(35) = rst2!CLNPState
                End If
                
                If Len(rst2!CLNOnlineStatus) Then
                    ClnMaster.SubItems(36) = rst2!CLNOnlineStatus
                End If
                
                rst2.MoveNext
        TotalCLNCount = TotalCLNCount + 1
        txtTotalRecord = TotalCLNCount
        Loop
    rst2.Close
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CLNComboListing()
Const METHOD_NAME = "CLNComboListing"
On Error GoTo debugs
Dim CLNGRP As New ADODB.Recordset
Dim CTY As New ADODB.Recordset
Dim STA As New ADODB.Recordset

    CLNGRP.Open "ClinicGrp", medixmasterconn, adOpenForwardOnly, adLockReadOnly
    CTY.Open "CTY", medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
    STA.Open "STA", medixmasterconn, adOpenForwardOnly, adLockReadOnly
        
    CboInfoClnGrp.Clear
    cboInfoClnBNFType.Clear
    cboInfoClnType.Clear
    cboInfoClnPanel.Clear
    cboInfoClnXray.Clear
    cboInfoClnCity.Clear
    cboInfoClnState.Clear
    cboInfoCln24hrs.Clear
    cboInfoFFS.Clear
    cboPymtInfoClnCity.Clear
    cboPymtInfoClnState.Clear
    cboOnlineStatus.Clear
    cboGLFormat.Clear
    
    Do While Not CLNGRP.EOF
        CboInfoClnGrp.AddItem CLNGRP!ClinicGrpCode & " - " & CLNGRP!ClinicGrpName
        CLNGRP.MoveNext
    Loop
        
    Do While Not CTY.EOF
        cboInfoClnCity.AddItem CTY!CTYDescription
        cboPymtInfoClnCity.AddItem CTY!CTYDescription
        CTY.MoveNext
    Loop
        
    Do While Not STA.EOF
        cboInfoClnState.AddItem STA!STADescription
        cboPymtInfoClnState.AddItem STA!STADescription
        STA.MoveNext
    Loop
    
    cboInfoClnBNFType.AddItem "1 - Standard"
    cboInfoClnBNFType.AddItem "2 - Clinic"
    cboInfoClnBNFType.AddItem "3 - Plan"
    cboInfoClnBNFType.ListIndex = 0
        
    cboInfoClnType.AddItem "G - General Practisional"
    cboInfoClnType.AddItem "S - Specialised"
        
    cboOnlineStatus.AddItem "Y - Yes"
    cboOnlineStatus.AddItem "N - No"

    cboInfoClnXray.AddItem "Y - Yes"
    cboInfoClnXray.AddItem "N - No"
    
    cboInfoClnPanel.AddItem "Y - Yes"
    cboInfoClnPanel.AddItem "N - No"
    cboInfoClnPanel.AddItem "D - Declined"
    cboInfoClnPanel.AddItem "C - Confirmed"
    
    cboInfoCln24hrs.AddItem "Y - Yes"
    cboInfoCln24hrs.AddItem "N - No"
    
    cboInfoFFS.AddItem "Y - Yes"
    cboInfoFFS.AddItem "N - No"
    
    cboGLFormat.AddItem "M - Medix"
    cboGLFormat.AddItem "H - Health Connect"
    
    CLNGRP.Close
    CTY.Close
    STA.Close
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstClinicInfoList_Click()
Const METHOD_NAME = "lstClinicInfoList_Click"
On Error GoTo debugs
Dim rst3 As New ADODB.Recordset
Dim StrSql As String
    
    If lstClinicInfoList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            
            StrSql = "Select * from CLN Where CLNCode = '" & lstClinicInfoList.SelectedItem.Text & "'"
            rst3.Open StrSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            
            txtInfoCLNCode = rst3!CLNCode
            
            If Trim(rst3!CLNName) <> " " Then
                txtInfoClnName = rst3!CLNName
            End If
            
            If Trim(rst3!ClinicGrpCode) <> " " Then
                CboInfoClnGrp = rst3!ClinicGrpCode
            End If
            
            If Trim(rst3!CLNAdd1) <> " " Then
                txtInfoClnAdd1 = rst3!CLNAdd1
            End If
            
            If Trim(rst3!CLNAdd2) <> " " Then
                txtInfoClnAdd2 = rst3!CLNAdd2
            End If
            
            If Trim(rst3!CLNAdd3) <> " " Then
                txtInfoClnAdd3 = rst3!CLNAdd3
            End If
            
            If Trim(rst3!CLNPhone) <> " " Then
                txtInfoClnTelNo = rst3!CLNPhone
            End If
            
            If Trim(rst3!CLNFax) <> " " Then
                txtInfoClnFaxNo = rst3!CLNFax
            End If
            
            If Trim(rst3!CLNEMAIL) <> " " Then
                txtInfoClnEmail = rst3!CLNEMAIL
            End If
            
            If Trim(rst3!CLNHomepage) <> " " Then
                txtInfoClnHome = rst3!CLNHomepage
            End If
            
            If Trim(rst3!CLNRepresentative) <> " " Then
                txtInfoClnContact = rst3!CLNRepresentative
            End If
            
            If Trim(rst3!CLNDocName) <> " " Then
                txtInfoClnDocName = rst3!CLNDocName
            End If
            
            If Trim(rst3!CLNBNFType) <> " " Then
                cboInfoClnBNFType = rst3!CLNBNFType
            End If
            
            If Trim(rst3!CLNForRef) <> " " Then
                txtInfoClnFormema = rst3!CLNForRef
            End If
            
            If Trim(rst3!CLNNoofDoc) <> " " Then
                txtInfoClnNoDoc = rst3!CLNNoofDoc
            End If
                   
            If Trim(rst3!CLNType) <> " " Then
                cboInfoClnType = rst3!CLNType
            End If
            
            If Trim(rst3!CLNXRay) <> " " Then
                cboInfoClnXray = rst3!CLNXRay
            End If
            
            If Trim(rst3!CLNPanel) <> " " Then
                cboInfoClnPanel = rst3!CLNPanel
            End If
            
            If Trim(rst3!CLNState) <> " " Then
                cboInfoClnState = rst3!CLNState
            End If
            
            If Trim(rst3!CLNCity) <> " " Then
                cboInfoClnCity = rst3!CLNCity
            End If
            
            If Trim(rst3!CLNPostCode) <> " " Then
                txtInfoClnPostCode = rst3!CLNPostCode
            End If
                   
            If Trim(rst3!CLNinitial) <> " " Then
                txtInfoClnMIX = rst3!CLNinitial
            End If
                    
            If rst3!CLNReceivedDate <> "__/__/____" Then
                txtInfoClnReceivedDate = rst3!CLNReceivedDate
            End If
                    
            If Trim(rst3!CLNBankAcc) <> " " Then
                txtInfoClnBankAcc = rst3!CLNBankAcc
            End If
            
            If Trim(rst3!CLNBankName) <> " " Then
                txtInfoClnBank = rst3!CLNBankName
            End If
            
            If Trim(rst3!CLNBankBrn) <> " " Then
                txtInfoClnBranch = rst3!CLNBankBrn
            End If
            
            If Trim(rst3!CLNAccPerson) <> " " Then
                txtInfoClnAccPerson = rst3!CLNAccPerson
            End If
            
            If Trim(rst3!CLN24Hrs) <> " " Then
                cboInfoCln24hrs = rst3!CLN24Hrs
            End If
    
            If Trim(rst3!CLNRemark) <> " " Then
                txtInfoClnRemark = rst3!CLNRemark
            End If
            
            If Trim(rst3!CLNLocation) <> " " Then
                TxtLocation = rst3!CLNLocation
            End If
            
            If Trim(rst3!CLNFFSStatus) <> " " Then
                cboInfoFFS = rst3!CLNFFSStatus
            End If
            
            If Trim(rst3!CLNMCOProcfeePerc) <> " " Then
                txtFFSPerc = rst3!CLNMCOProcfeePerc
            End If
            
            If Len(rst3!CLNOperatingHr) Then
                txtOperatinghr = rst3!CLNOperatingHr
            End If
                    
            If Len(rst3!CLNPName) Then
                txtPymtInfoClnName = rst3!CLNPName
            End If
            
            If Len(rst3!CLNPAdd1) Then
                txtPymtInfoClnAdd1 = rst3!CLNPAdd1
            End If
            
            If Len(rst3!CLNPAdd2) Then
                txtPymtInfoClnAdd2 = rst3!CLNPAdd2
            End If
            
            If Len(rst3!CLNPAdd3) Then
                txtPymtInfoClnAdd3 = rst3!CLNPAdd3
            End If
            
            If Len(rst3!CLNPPostCode) Then
                txtPymtInfoClnPostCode = rst3!CLNPPostCode
            End If
            
            If Len(rst3!CLNPcity) Then
                cboPymtInfoClnCity = rst3!CLNPcity
            End If
            
            If Len(rst3!CLNPState) Then
                cboPymtInfoClnState = rst3!CLNPState
            End If
            
            If Len(rst3!CLNOnlineStatus) Then
                cboOnlineStatus = rst3!CLNOnlineStatus
            End If
            
            If Len(rst3!CLNGLFormat) Then
                cboGLFormat = rst3!CLNGLFormat
            End If
            
            If Len(rst3!CLNRegNo) Then
                txtRegNo = rst3!CLNRegNo
            End If
            
            
            
            Call EntriesEnable(False)
            cmdSave.Enabled = False
            rst3.Close
            Set rst3 = Nothing
        
            'medixmasterconn.Close
            'Set medixmasterconn = Nothing
   End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstClinicInfoList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstClinicInfoList.SortKey = ColumnHeader.Index - 1
lstClinicInfoList.Sorted = True
    If lstClinicInfoList.SortOrder = lvwAscending Then
        lstClinicInfoList.SortOrder = lvwDescending
    Else
        lstClinicInfoList.SortOrder = lvwAscending
    End If
End Sub

Public Function ExportToExcelClinic(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Clinic.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                '    Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                '       ' If tag is empty, format as text
                '        Case "string", ""
                '        .NumberFormat = "@"
                '        .HorizontalAlignment = xlLeft
                '        Case "number"
                '        .NumberFormat = "#,##0.00_);(#,##0.00)"
                '        .HorizontalAlignment = xlRight
                '        Case "date"
                '        .NumberFormat = "dd/mm/yyyy"
                '        .HorizontalAlignment = xlCenter 'xlRight
                '    End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelClinic = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

'**********************************End Cln Info**************************************


Private Sub SSTab1_Click(PreviousTab As Integer)
    Screen.MousePointer = vbHourglass
    
    'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
        Select Case SSTab1.Tab
        'Case 0
             'CmdRefresh.Enabled = True
             'lstGrpCoList.ListItems.Clear
             'txtTotalRecord = 0
             'Call ClearForm(Me)
             'Call GRPCompanyComboListing
             'txtGrpCoCode.Enabled = False
        Case 0
             CmdRefresh.Enabled = True
             lstGrpClinicList.ListItems.Clear
             txtTotalRecord = 0
             Call ClearForm(Me)
             Call GRPClinicComboListing
             txtClnGrpCode.Enabled = False
        Case 1
             CmdRefresh.Enabled = True
             lstDoctorList.ListItems.Clear
             txtTotalRecord = 0
             Call ClearForm(Me)
             txtDocCode.Enabled = False
        Case 2
             CmdRefresh.Enabled = True
             lstClinicInfoList.ListItems.Clear
             txtTotalRecord = 0
             Call ClearForm(Me)
             Call CLNComboListing
             txtInfoCLNCode.Enabled = False
        Case 3
             CmdRefresh.Enabled = True
             lstGrpComCLN.ListItems.Clear
             txtTotalRecord = 0
             Call ClearForm(Me)
             'Call GRPCOMCLNComboListing
             CboGrpComName.Enabled = True
        End Select
    
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
    'MediConnCISPayment.Close
    'Set MediConnCISPayment = Nothing
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
    Screen.MousePointer = Default

End Sub

'########################### Start Enteries Enable ###########################'
Private Sub EntriesEnable(status As Boolean)
   
    Dim ctl As Control

    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.Enabled = status
        End If
    Next ctl
   
    
End Sub
'########################### End Entries Enable ###########################'

Private Sub txtClnEffDate_LostFocus()
    If txtClnEffDate <> "__/__/____" Then
        If IsDate(txtClnEffDate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtClnEffDate.SetFocus
        ElseIf Format(txtClnEffDate, "yyyymmdd") > Format(Date, "yyyymmdd") Then
            MsgBox "Date cannot greater than system date ! ", vbCritical
            txtClnEffDate.SetFocus
        End If
    End If
End Sub

Private Sub txtClnFax_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub





Private Sub txtClnPhone_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtEffDateFrom_LostFocus()
    If txtEffDateFrom <> "__/__/____" Then
        If IsDate(txtEffDateFrom) = False Then
            MsgBox "Invalid Effective Date From format!", vbCritical
            txtEffDateFrom.SetFocus
        End If
    End If
End Sub


Private Sub txtEffDateTo_LostFocus()
If txtEffDateTo <> "__/__/____" Then
        If IsDate(txtEffDateTo) = False Then
            MsgBox "Invalid Effective Date To format!", vbCritical
            txtEffDateTo.SetFocus
        End If
    End If
End Sub

'Private Sub txtCoEffDate_LostFocus()
    'If txtCoEffDate <> "__/__/____" Then
        'If IsDate(txtCoEffDate) = False Then
            'MsgBox "Invalid date format!", vbCritical
            'txtCoEffDate.SetFocus
        'ElseIf Format(txtCoEffDate, "yyyymmdd") > Format(Date, "yyyymmdd") Then
            'MsgBox "Date cannot greater than system date ! ", vbCritical
            'txtCoEffDate.SetFocus
        'End If
    'End If
'End Sub

'Private Sub txtInfoClnAdd1_Change()
'    Clipboard.Clear
'End Sub

'Private Sub txtInfoClnAdd2_Change()
'    Clipboard.Clear
'End Sub

'Private Sub txtInfoClnAdd3_Change()
'    Clipboard.Clear
'End Sub

Private Sub txtInfoClnFaxNo_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtInfoClnNoDoc_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtInfoClnPostCode_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtInfoClnReceivedDate_LostFocus()
    If txtInfoClnReceivedDate <> "__/__/____" Then
        If IsDate(txtInfoClnReceivedDate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtInfoClnReceivedDate.SetFocus
        ElseIf Format(txtInfoClnReceivedDate, "yyyymmdd") > Format(Date, "yyyymmdd") Then
            MsgBox "Date cannot greater than system date ! ", vbCritical
            txtInfoClnReceivedDate.SetFocus
        End If
    End If
End Sub

'Private Sub txtCoFax_KeyPress(KeyAscii As Integer)
    'Clipboard.Clear
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

'Private Sub txtCoPhone_KeyPress(KeyAscii As Integer)
    'Clipboard.Clear
    'KeyAscii = KeyNumber(KeyAscii)
'End Sub

Private Sub txtInfoClnTelNo_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub CboGrpComName_Change()
    If InStr(CboGrpComName.Text, "null") Then
       CboGrpComName.Enabled = False
    End If
End Sub

Private Sub CboGrpComName_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboGrpComName.Text, CboGrpComName.SelStart) + Chr(KeyAscii) + Mid(CboGrpComName.Text, CboGrpComName.SelStart + CboGrpComName.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboGrpComName.ListCount - 1
 
        If NewText = LCase(Left(CboGrpComName.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboGrpComName.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboGrpComName.Text = ValidValue
               CboGrpComName.SelStart = 0
               CboGrpComName.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboCLNName_Change()
    If InStr(CboCLNName.Text, "null") Then
       CboCLNName.Enabled = False
    End If
End Sub

Private Sub cboCLNName_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCLNName.Text, CboCLNName.SelStart) + Chr(KeyAscii) + Mid(CboCLNName.Text, CboCLNName.SelStart + CboCLNName.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboCLNName.ListCount - 1
 
        If NewText = LCase(Left(CboCLNName.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCLNName.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboCLNName.Text = ValidValue
               CboCLNName.SelStart = 0
               CboCLNName.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub GRPCOMCLNComboListing()
Const METHOD_NAME = "GRPCOMCLNComboListing"
On Error GoTo debugs
Dim ClinicName As New ADODB.Recordset
Dim GRPCOMName As New ADODB.Recordset

    CboGrpComName.Clear
    CboCLNName.Clear
    
    'If Check2.Value = 1 Then
    
        ClinicName.Open "CLN", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        
        Do Until ClinicName.EOF
            CboCLNName.AddItem Trim(ClinicName!CLNCode) & " - " & Trim(ClinicName!CLNName)
            ClinicName.MoveNext
        Loop
        
        ClinicName.Close
    
    'End If
    
    'If Check1.Value = 1 Then
    
        GRPCOMName.Open "GRPCompany", MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic
        
        Do Until GRPCOMName.EOF
            CboGrpComName.AddItem Trim(GRPCOMName!GRPCoCode) & " - " & Trim(GRPCOMName!GrpCoName)
            GRPCOMName.MoveNext
        Loop
        
        GRPCOMName.Close
            
    'End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveClnAccess()
Const METHOD_NAME = "SaveClnAccess"
On Error GoTo debugs
Dim RecordSaved
Dim ClnAccess As New ADODB.Recordset
Dim SaveClnAccess As New ADODB.Recordset
Dim StrSql As String
Dim COMSeq As String

    If txtCAClnCode.Text = "" Then
        MsgBox "Please Enter Clinic Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtCAClnCode.SetFocus
        Exit Sub
    ElseIf IsDate(txtEffDateFrom.Text) = False Then
        MsgBox "Invalid Effective Date From", vbOKOnly + vbCritical, DialogBoxTitle
        txtEffDateFrom.SetFocus
        Exit Sub
    ElseIf IsDate(txtEffDateTo.Text) = False Then
        MsgBox "InvalidEffective Date To", vbOKOnly + vbCritical, DialogBoxTitle
        txtEffDateTo.SetFocus
        Exit Sub
    ElseIf DateValue(txtEffDateFrom.Text) > DateValue(txtEffDateTo.Text) Then
        MsgBox "Effective Date From cannot be more than Effective Date To!", vbOKOnly + vbCritical, DialogBoxTitle
        txtEffDateTo.SetFocus
        Exit Sub
    Else
        StrSql = " Select CLNName From CLN Where CLNCode = '" & Trim(txtCAClnCode.Text) & "'"
        
        ClnAccess.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If ClnAccess.RecordCount = 0 Then
            MsgBox "Clinic Code does not exists, please enter again!", vbCritical, "Error"
            ClnAccess.Close
            txtCAClnCode.SetFocus
            Set ClnAccess = Nothing
            Exit Sub
        Else
            ClnAccess.Close
            Set ClnAccess = Nothing
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                    ClnAccess.Open "CLNAccessControl", medixmasterconn, adOpenKeyset, adLockOptimistic
                    ClnAccess.AddNew
                End If
                
                With ClnAccess
                    !ClinicCode = Trim(UCase(txtCAClnCode.Text))
                    !EffDateFrom = DateValue(txtEffDateFrom.Text)
                    !EffDateTo = DateValue(txtEffDateTo.Text)
                    !CreatedBy = Logon
                    !CreatedDate = Now()
                End With
                ClnAccess.update

                lstClinicAccess.ListItems.Clear
                Call ClinicAccessListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                    
                ClnAccess.Close
                Set SaveClnAccess = Nothing
            End If
            EditFlag = False
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveGrpComClinic()
Const METHOD_NAME = "SaveGrpComClinic"
On Error GoTo debugs
Dim RecordSaved
Dim CompanyGrp As New ADODB.Recordset
Dim SaveCompanyGrp As New ADODB.Recordset
Dim StrSql As String
Dim COMSeq As String
        
    If CboGrpComName = "" Then
        MsgBox "Please Enter Group Company Name", vbOKOnly + vbCritical, DialogBoxTitle
        CboGrpComName.SetFocus
    ElseIf CboCLNName = "" Then
        MsgBox "Please Enter Clinic Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboCLNName.SetFocus
    Else
        StrSql = " Select * From GRPCOMClinic Where GRPComCode = '" & Trim(Mid(CboGrpComName.Text, 1, IIf(InStr(1, CboGrpComName.Text, "-") = 0, 10, InStr(1, CboGrpComName.Text, "-") - 1))) & "'" & _
                 " AND GRPComClnCode = '" & Trim(Mid(CboCLNName.Text, 1, IIf(InStr(1, CboCLNName.Text, "-") = 0, 10, InStr(1, CboCLNName.Text, "-") - 1))) & "'"
        
        CompanyGrp.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
      If CompanyGrp.RecordCount > 0 Then 'And EditFlag = False Then
      MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
      CompanyGrp.Close
      CboGrpComName.SetFocus
      CboGrpComName.SelStart = 0
      CboGrpComName.SelLength = Len(CboGrpComName)
      Set CompanyGrp = Nothing
      Exit Sub
      Else
      CompanyGrp.Close
      Set CompanyGrp = Nothing
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            If EditFlag = False Then
                CompanyGrp.Open "GRPCOMClinic", MediConnCISDB, adOpenKeyset, adLockOptimistic
                CompanyGrp.AddNew
            Else
                StrSql = " Select * From GRPCOMClinic Where GRPComCode = '" & lstGrpComCLN.SelectedItem.Text & "'" & _
                 " And GrpComClnCode = '" & lstGrpComCLN.SelectedItem.SubItems(2) & "'"
                CompanyGrp.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            End If
            
                With CompanyGrp
                    !GRPComCode = Trim(Mid(CboGrpComName.Text, 1, IIf(InStr(1, CboGrpComName.Text, "-") = 0, 10, InStr(1, CboGrpComName.Text, "-") - 1)))
                    !GrpComClnCode = Trim(Mid(CboCLNName.Text, 1, IIf(InStr(1, CboCLNName.Text, "-") = 0, 10, InStr(1, CboCLNName.Text, "-") - 1)))
                    !GrpComClnLastUIpdateUser = Logon
                    !GrpComClnLastUIpdateDate = Date
                End With
                CompanyGrp.update
                'Call ClearForm(Me)
                'Check1.Value = 1
                'CboGrpComName.Text = ""
                'CboCLNName.Text = ""
                lstGrpComCLN.ListItems.Clear
                Call GRPComCLNListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                
            CompanyGrp.Close
            Set SaveCompanyGrp = Nothing
        End If
        EditFlag = False
    End If
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub GRPComCLNListing(Optional strAppend As String)
Const METHOD_NAME = "GRPComCLNListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim CompanyGrpMaster As ListItem
Dim sql As String
Dim TotalGRPCOMCount As String
    
    sql = "Select * from VIEW_GrpCom_Clinic where 0=0 "
   
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
              
    TotalGRPCOMCount = 0
    lstGrpComCLN.ListItems.Clear
    rst2.Open sql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set CompanyGrpMaster = lstGrpComCLN.ListItems.Add(, , Trim(rst2!GRPComCode))
                If Len(rst2!GrpCoName) Then
                    CompanyGrpMaster.SubItems(1) = Trim(rst2!GrpCoName)
                End If
                If Len(rst2!GrpComClnCode) Then
                    CompanyGrpMaster.SubItems(2) = Trim(rst2!GrpComClnCode)
                End If
                If Len(rst2!CLNName) Then
                    CompanyGrpMaster.SubItems(3) = Trim(rst2!CLNName)
                End If
                
                If Len(rst2!CLNAdd1) Then
                    CompanyGrpMaster.SubItems(4) = Trim(rst2!CLNAdd1 & rst2!CLNAdd2 & rst2!CLNAdd3)
                End If
                If Len(rst2!CLNPostCode) Then
                    CompanyGrpMaster.SubItems(5) = Trim(rst2!CLNPostCode)
                End If
                If Len(rst2!CLNState) Then
                    CompanyGrpMaster.SubItems(6) = Trim(rst2!CLNState)
                End If
                
                If Len(rst2!CLNCity) Then
                    CompanyGrpMaster.SubItems(7) = Trim(rst2!CLNCity)
                End If
                
                If Len(rst2!CLNLocation) Then
                    CompanyGrpMaster.SubItems(8) = Trim(rst2!CLNLocation)
                End If
                
                If Len(rst2!CLN24Hrs) Then
                    CompanyGrpMaster.SubItems(9) = Trim(rst2!CLN24Hrs)
                End If
                
                If Len(rst2!CLNPhone) Then
                    CompanyGrpMaster.SubItems(10) = Trim(rst2!CLNPhone)
                End If
                
                If Len(rst2!CLNFax) Then
                    CompanyGrpMaster.SubItems(11) = Trim(rst2!CLNFax)
                End If
                
                'If Len(rst2!CLNPName) Then
                '    CompanyGrpMaster.SubItems(12) = Trim(rst2!CLNPName)
                'End If
                
                'If Len(rst2!CLNPAdd1) Then
                '    CompanyGrpMaster.SubItems(13) = Trim(rst2!CLNPAdd1)
                'End If
                
                'If Len(rst2!CLNPAdd2) Then
                '    CompanyGrpMaster.SubItems(14) = Trim(rst2!CLNPAdd2)
                'End If
                
                'If Len(rst2!CLNPAdd3) Then
                '    CompanyGrpMaster.SubItems(15) = Trim(rst2!CLNAdd3)
                'End If
                
                'If Len(rst2!CLNPPostCode) Then
                '    CompanyGrpMaster.SubItems(16) = Trim(rst2!CLNPPostCode)
                'End If
                
                'If Len(rst2!CLNPcity) Then
                '    CompanyGrpMaster.SubItems(17) = Trim(rst2!CLNCity)
                'End If
                
                'If Len(rst2!CLNPState) Then
                '    CompanyGrpMaster.SubItems(18) = Trim(rst2!CLNPState)
                'End If
                

                rst2.MoveNext
        TotalGRPCOMCount = CDbl(TotalGRPCOMCount) + 1
        txtTotalRecord = CDbl(TotalGRPCOMCount)
        Loop
    rst2.Close
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub SearchClinicAccess()
    Screen.MousePointer = vbHourglass
    Call ClinicAccessListing
    Screen.MousePointer = Default
End Sub

Private Sub SearchGrpComCLN()
Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
        strAppend = ""
            
        If Len(Trim(CboGrpComName)) Then
            strAppend = strAppend + " And GRPComCode = '" & Trim(Mid(CboGrpComName.Text, 1, IIf(InStr(1, CboGrpComName.Text, "-") = 0, 10, InStr(1, CboGrpComName.Text, "-") - 1))) & "'"
        End If
        
        If Len(Trim(CboCLNName)) Then
            strAppend = strAppend + " And GRPComClnCode = '" & Trim(Mid(CboCLNName.Text, 1, IIf(InStr(1, CboCLNName.Text, "-") = 0, 10, InStr(1, CboCLNName.Text, "-") - 1))) & "'"
        End If
        
        GRPCOMCLNCriteria = strAppend
        Call GRPComCLNListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstGrpComCLN_Click()
Const METHOD_NAME = "lstGrpComCLN_Click"
On Error GoTo debugs
Dim rst3 As New ADODB.Recordset
Dim StrSql As String
        
    If lstGrpComCLN.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        'Call ClearForm(Me)
        'Check1.Value = 1
       ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            CboGrpComName.Text = ""
            CboCLNName.Text = ""
            
            StrSql = " Select * from VIEW_GrpCom_Clinic Where GRPComCode = '" & lstGrpComCLN.SelectedItem.Text & "'" & _
                     " And GrpComClnCode = '" & lstGrpComCLN.SelectedItem.SubItems(2) & "'"
            rst3.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic
            
            CboGrpComName = Trim(rst3!GRPComCode)
            CboCLNName = Trim(rst3!GrpComClnCode)
            Call EntriesEnable(False)
            cmdSave.Enabled = False
            rst3.Close
            Set rst3 = Nothing
       ' MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstGrpComCLN_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rst3 As New ADODB.Recordset
Dim StrSql As String
        
    If lstGrpComCLN.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        'Call ClearForm(Me)
        CboGrpComName.Text = ""
        CboCLNName.Text = ""
        'lstGrpComCLN.ListItems.Clear
       ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            StrSql = " Select * from VIEW_GrpCom_Clinic Where GRPComCode = '" & lstGrpComCLN.SelectedItem.Text & "'" & _
                     " And GrpComClnCode = '" & lstGrpComCLN.SelectedItem.SubItems(2) & "'"
            rst3.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic
            
            CboGrpComName = Trim(rst3!GRPComCode)
            CboCLNName = Trim(rst3!GrpComClnCode)
            Call EntriesEnable(False)
            cmdSave.Enabled = False
            rst3.Close
            Set rst3 = Nothing
      '  MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
    End If
End Sub

Private Sub lstGrpComCLN_KeyUp(KeyCode As Integer, Shift As Integer)
Const METHOD_NAME = "lstGrpComCLN_KeyUp"
On Error GoTo debugs
Dim rst3 As New ADODB.Recordset
Dim StrSql As String
        
    If lstGrpComCLN.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        'Call ClearForm(Me)
        CboGrpComName.Text = ""
        CboCLNName.Text = ""
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            StrSql = " Select * from VIEW_GrpCom_Clinic Where GRPComCode = '" & lstGrpComCLN.SelectedItem.Text & "'" & _
                     " And GrpComClnCode = '" & lstGrpComCLN.SelectedItem.SubItems(2) & "'"
            rst3.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic
            
            CboGrpComName = Trim(rst3!GRPComCode)
            CboCLNName = Trim(rst3!GrpComClnCode)
            Call EntriesEnable(False)
            cmdSave.Enabled = False
            rst3.Close
            Set rst3 = Nothing
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub DeleteGrpComClinic()
Const METHOD_NAME = "DeleteGrpComClinic"
On Error GoTo debugs
Dim DeleteDr As New ADODB.Recordset
Dim rsChkDr As New ADODB.Recordset
Dim StrSql, sqlDoctor As String
Dim DeleteRecord

EditFlag = False
If lstGrpComCLN.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   'sqlDoctor = " Select GRPCOMClinic from Where GRPComCode = '" & Trim(lstGrpComCLN.SelectedItem.Text) & "'" & _
               " AND GRPComClnCode = '" & Trim(lstGrpComCLN.SelectedItem.SubItems(1)) & "'"
   
   'rsChkDr.MaxRecords = 1
   'rsChkDr.Open sqlDoctor, MediConnCISDB, adOpenKeyset, adLockOptimistic
   
   'If rsChkDr.RecordCount Then
        'MsgBox "This record cannot be deleted because it is used in Doctor Table!", vbOKOnly + vbCritical, "Error Message"
   'Else
    DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then 'And rsChkDr.RecordCount = 0 Then
            StrSql = " Delete From GRPCOMClinic Where GRPComCode = '" & Trim(lstGrpComCLN.SelectedItem.Text) & "'" & _
                     " AND GRPComClnCode = '" & Trim(lstGrpComCLN.SelectedItem.SubItems(2)) & "'"
            DeleteDr.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
            cmdSave.Enabled = True
            'Call ClearForm(Me)
            'Check1.Value = 1
            CboGrpComName.Text = ""
            CboCLNName.Text = ""
            lstGrpComCLN.ListItems.Clear
            Call GRPComCLNListing
            'Call DisableCmdButton(Me)
            Call EntriesEnable(True)
            CboGrpComName.SetFocus
        End If
     'End If
End If
Exit Sub
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub PrintGrpComCln()
    RptGrpComCLN.Show
End Sub
