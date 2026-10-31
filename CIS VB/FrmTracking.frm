VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTracking 
   Caption         =   "Clinic Tracking"
   ClientHeight    =   8412
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8412
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   40
      Top             =   7320
      Width           =   11775
      Begin VB.TextBox lblTotalRecord 
         Height          =   285
         Left            =   1560
         TabIndex        =   49
         Top             =   465
         Width           =   975
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   7080
         TabIndex        =   17
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9720
         TabIndex        =   20
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10560
         TabIndex        =   21
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         Height          =   375
         Left            =   8760
         TabIndex        =   19
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   7920
         TabIndex        =   18
         Top             =   160
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "of"
         Height          =   255
         Left            =   1320
         TabIndex        =   44
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1560
         TabIndex        =   41
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Clinic Tracking"
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
         TabIndex        =   39
         Top             =   80
         Width           =   3285
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2295
      Left            =   0
      TabIndex        =   22
      Top             =   360
      Width           =   11775
      Begin VB.TextBox txtCLNCode 
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
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   0
         Top             =   240
         Width           =   930
      End
      Begin VB.TextBox txtCLNName 
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
         Left            =   4020
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   7515
      End
      Begin VB.ComboBox CboClinicGrp 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   500
         Width           =   3615
      End
      Begin VB.TextBox txtCLNpostcode 
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
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   3
         Top             =   760
         Width           =   1455
      End
      Begin VB.TextBox txtMIXClinic 
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
         Left            =   6720
         MaxLength       =   50
         TabIndex        =   8
         Top             =   500
         Width           =   4815
      End
      Begin VB.ComboBox cboCLNType 
         Height          =   315
         Left            =   6720
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   750
         Width           =   1815
      End
      Begin VB.ComboBox cboCLNcity 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1020
         Width           =   3615
      End
      Begin VB.ComboBox cboCLNstate 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1300
         Width           =   3615
      End
      Begin VB.TextBox TxtCLNContact 
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1570
         Width           =   3615
      End
      Begin VB.ComboBox cboCLNBNFType 
         Height          =   315
         Left            =   9600
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   750
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtCLNReceivedDate 
         Height          =   285
         Left            =   6720
         TabIndex        =   11
         Top             =   1050
         Width           =   1815
         _ExtentX        =   3196
         _ExtentY        =   487
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtCLNDocName 
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1830
         Width           =   3615
      End
      Begin VB.ComboBox cboCLNpanel 
         Height          =   315
         Left            =   6720
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   1300
         Width           =   1815
      End
      Begin VB.TextBox txtCLNformema 
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
         Left            =   6720
         MaxLength       =   15
         TabIndex        =   15
         Top             =   1590
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txtCLNReceivedDate2 
         Height          =   285
         Left            =   9600
         TabIndex        =   12
         Top             =   1050
         Width           =   1935
         _ExtentX        =   3408
         _ExtentY        =   487
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboCLNxray 
         Height          =   315
         Left            =   9600
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   1300
         Width           =   1935
      End
      Begin VB.TextBox TxtLocation 
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
         MaxLength       =   100
         TabIndex        =   16
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox cboInfoFFS 
         Height          =   315
         Left            =   6720
         Sorted          =   -1  'True
         TabIndex        =   46
         Top             =   1850
         Width           =   1815
      End
      Begin VB.Label txtFFSStatus 
         Caption         =   "FFS Status"
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
         Left            =   5400
         TabIndex        =   47
         Top             =   1900
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Location"
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
         Left            =   8640
         TabIndex        =   45
         Top             =   1695
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "To"
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
         Left            =   8640
         TabIndex        =   42
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label66 
         Caption         =   "Received Date"
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
         Left            =   5400
         TabIndex        =   37
         Top             =   1100
         Width           =   1455
      End
      Begin VB.Label Label65 
         Caption         =   "Clinic Initial"
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
         Left            =   5400
         TabIndex        =   36
         Top             =   555
         Width           =   1215
      End
      Begin VB.Label Label64 
         Caption         =   "For. Ref."
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
         Left            =   5400
         TabIndex        =   35
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label62 
         Caption         =   "X-Ray"
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
         Left            =   8640
         TabIndex        =   34
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Panel"
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
         Left            =   5400
         TabIndex        =   33
         Top             =   1350
         Width           =   1095
      End
      Begin VB.Label Label48 
         Caption         =   "Doctor Name"
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
         Left            =   240
         TabIndex        =   32
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label49 
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
         Left            =   240
         TabIndex        =   31
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label63 
         Caption         =   "State"
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
         Left            =   240
         TabIndex        =   30
         Top             =   1400
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Clinic Type"
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
         Left            =   5400
         TabIndex        =   29
         Top             =   795
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "Bnf. Type"
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
         Left            =   8640
         TabIndex        =   28
         Top             =   795
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Clinic Grp. Code"
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
         TabIndex        =   27
         Top             =   550
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Clinic Code        "
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
         Index           =   6
         Left            =   240
         TabIndex        =   26
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Clinic Name"
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
         Index           =   6
         Left            =   2760
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label61 
         Caption         =   "Post Code"
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
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label Label60 
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1100
         Width           =   525
      End
   End
   Begin MSComctlLib.ListView lstClinicList 
      Height          =   4560
      Left            =   0
      TabIndex        =   43
      Top             =   2760
      Width           =   11745
      _ExtentX        =   20722
      _ExtentY        =   8043
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
      NumItems        =   29
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Clinic Code"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Clinic Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Clinic Group Code"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Clinic Initial"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Clinic Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Benefit Type"
         Object.Width           =   2117
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
         Object.Width           =   2117
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
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Received Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "Panel"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "X-Ray"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Formema Ref"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "No Of Doctor"
         Object.Width           =   2117
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
         Text            =   "Home Page"
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
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "FFS Status"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FrmTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intExit As Integer
Private m_blnDirection(1 To 28) As Boolean
Private Sub CLNComboListing()
Const METHOD_NAME = "CLNComboListing"
On Error GoTo debugs
Dim CLNGRP As New ADODB.Recordset
Dim CTY As New ADODB.Recordset
Dim STA As New ADODB.Recordset
    CLNGRP.Open "ClinicGrp", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    CTY.Open "CTY", medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
    STA.Open "STA", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    CboClinicGrp.Clear
    cboCLNBNFType.Clear
    cboCLNType.Clear
    cboCLNpanel.Clear
    cboCLNxray.Clear
    cboCLNcity.Clear
    cboCLNstate.Clear
    
    Do Until CLNGRP.EOF
        CboClinicGrp.AddItem CLNGRP!ClinicGrpCode & " - " & CLNGRP!ClinicGrpName
        CLNGRP.MoveNext
    Loop
        
    Do Until CTY.EOF
        cboCLNcity.AddItem CTY!CTYDescription
        CTY.MoveNext
    Loop
        
    Do Until STA.EOF
        cboCLNstate.AddItem STA!STADescription
        STA.MoveNext
    Loop
    
    CLNGRP.Close
    Set CLNGRP = Nothing
    CTY.Close
    Set CTY = Nothing
    STA.Close
    Set STA = Nothing
    cboCLNBNFType.AddItem "1 - Standard"
    cboCLNBNFType.AddItem "2 - Plan"
    cboCLNType.AddItem "G - General Practisional"
    cboCLNType.AddItem "S - Specialised"
    cboCLNxray.AddItem "Y - Yes"
    cboCLNxray.AddItem "N - No"
    cboCLNpanel.AddItem "Y - Yes"
    cboCLNpanel.AddItem "N - No"
    cboCLNpanel.AddItem "C - Confirm"
    cboInfoFFS.AddItem "Y - Yes"
    cboInfoFFS.AddItem "N - No"
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CboClinicGrp_Change()

    If CboClinicGrp.Text = "" Then
        If InStr(CboClinicGrp.Text, "null") Then
           CboClinicGrp.Enabled = False
        End If
    End If
End Sub

Private Sub CboClinicGrp_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboClinicGrp_KeyPress"
On Error GoTo debugs

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClinicGrp.Text, CboClinicGrp.SelStart) + Chr(KeyAscii) + Mid(CboClinicGrp.Text, CboClinicGrp.SelStart + CboClinicGrp.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To CboClinicGrp.ListCount - 1
        If NewText = LCase(Left(CboClinicGrp.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClinicGrp.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboClinicGrp.Text = ValidValue
            CboClinicGrp.SelStart = 0
            CboClinicGrp.SelLength = Len(ValidValue)
        End If
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cboCLNBNFType_Change()
    If cboCLNBNFType.Text = "" Then
        If InStr(cboCLNBNFType.Text, "null") Then
           cboCLNBNFType.Enabled = False
        End If
    End If
End Sub

Private Sub cboCLNBNFType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNBNFType.Text, cboCLNBNFType.SelStart) + Chr(KeyAscii) + Mid(cboCLNBNFType.Text, cboCLNBNFType.SelStart + cboCLNBNFType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCLNBNFType.ListCount - 1
 
        If NewText = LCase(Left(cboCLNBNFType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNBNFType.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCLNBNFType.Text = ValidValue
                cboCLNBNFType.SelStart = 0
                cboCLNBNFType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCLNcity_Change()

    If cboCLNcity.Text = "" Then
        If InStr(cboCLNcity.Text, "null") Then
           cboCLNcity.Enabled = False
        End If
    End If
End Sub

Private Sub cboCLNcity_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNcity.Text, cboCLNcity.SelStart) + Chr(KeyAscii) + Mid(cboCLNcity.Text, cboCLNcity.SelStart + cboCLNcity.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCLNcity.ListCount - 1
 
        If NewText = LCase(Left(cboCLNcity.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNcity.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCLNcity.Text = ValidValue
                cboCLNcity.SelStart = 0
                cboCLNcity.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCLNpanel_Change()

    If cboCLNpanel.Text = "" Then
        If InStr(cboCLNpanel.Text, "null") Then
           cboCLNpanel.Enabled = False
        End If
    End If
End Sub

Private Sub cboCLNpanel_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNpanel.Text, cboCLNpanel.SelStart) + Chr(KeyAscii) + Mid(cboCLNpanel.Text, cboCLNpanel.SelStart + cboCLNpanel.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCLNpanel.ListCount - 1
 
        If NewText = LCase(Left(cboCLNpanel.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNpanel.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCLNpanel.Text = ValidValue
                cboCLNpanel.SelStart = 0
                cboCLNpanel.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCLNstate_Change()

    If cboCLNstate.Text = "" Then
        If InStr(cboCLNstate.Text, "null") Then
           cboCLNstate.Enabled = False
        End If
    End If
End Sub

Private Sub cboCLNstate_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNstate.Text, cboCLNstate.SelStart) + Chr(KeyAscii) + Mid(cboCLNstate.Text, cboCLNstate.SelStart + cboCLNstate.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCLNstate.ListCount - 1
 
        If NewText = LCase(Left(cboCLNstate.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNstate.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCLNstate.Text = ValidValue
                cboCLNstate.SelStart = 0
                cboCLNstate.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCLNType_Change()

    If cboCLNType.Text = "" Then
        If InStr(cboCLNType.Text, "null") Then
           cboCLNType.Enabled = False
        End If
    End If
End Sub

Private Sub cboCLNType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNType.Text, cboCLNType.SelStart) + Chr(KeyAscii) + Mid(cboCLNType.Text, cboCLNType.SelStart + cboCLNType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCLNType.ListCount - 1
 
        If NewText = LCase(Left(cboCLNType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNType.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCLNType.Text = ValidValue
                cboCLNType.SelStart = 0
                cboCLNType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCLNxray_Change()

    If cboCLNxray.Text = "" Then
        If InStr(cboCLNxray.Text, "null") Then
           cboCLNxray.Enabled = False
        End If
    End If
End Sub

Private Sub cboCLNxray_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNxray.Text, cboCLNxray.SelStart) + Chr(KeyAscii) + Mid(cboCLNxray.Text, cboCLNxray.SelStart + cboCLNxray.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCLNxray.ListCount - 1
 
        If NewText = LCase(Left(cboCLNxray.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNxray.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCLNxray.Text = ValidValue
                cboCLNxray.SelStart = 0
                cboCLNxray.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboInfoFFS_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

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

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lblCurrent = 0
    lblTotalRecord = 0
    lstClinicList.ListItems.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
    
    If lstClinicList.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstClinicList)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
Const METHOD_NAME = "cmdSearch_Click"
On Error GoTo debugs

Dim strAppend As String
   
Screen.MousePointer = vbHourglass
strAppend = ""
    
    If Len(txtCLNCode) Then
        strAppend = strAppend + " And CLNCode = '" & Trim(txtCLNCode) & "'"
    End If
    
    If Len(Trim(txtCLNName)) Then
        strAppend = strAppend + " And CLNName like '%" & ChkStr(txtCLNName) & "%'"
    End If
    
    If Len(Trim(CboClinicGrp)) Then
        strAppend = strAppend + " AND ClinicGrpCode like '" & RemStr(Mid(CboClinicGrp, 1, IIf(InStr(1, CboClinicGrp, "-") = 0, 4, InStr(1, CboClinicGrp, "-") - 1))) & "%'"
    End If
    
    If Len(txtCLNpostcode) Then
        strAppend = strAppend + " And CLNPostCode = '" & Trim(txtCLNpostcode) & "'"
    End If
    
    If Len(cboCLNcity) Then
        strAppend = strAppend + " And CLNCity = '" & Trim(cboCLNcity) & "'"
    End If
    
    If Len(cboCLNstate) Then
        strAppend = strAppend + " And CLNState = '" & Trim(cboCLNstate) & "'"
    End If
    
    If Len(Trim(TxtCLNContact)) Then
        strAppend = strAppend + " And CLNRepresentative like '%" & ChkStr(TxtCLNContact) & "%'"
    End If
    
    If Len(Trim(TxtCLNDocName)) Then
        strAppend = strAppend + " And CLNDocName like '%" & ChkStr(TxtCLNDocName) & "%'"
    End If
       
    If Len(Trim(txtMIXClinic)) Then
        strAppend = strAppend + " And CLNInitial like '%" & ChkStr(txtMIXClinic) & "%'"
    End If
    
    If Len(cboCLNBNFType) Then
        strAppend = strAppend + " And CLNBNFType = '" & Mid(cboCLNBNFType, 1, 1) & "'"
    End If
    
    If Len(Trim(txtCLNReceivedDate)) > 0 And IsDate(txtCLNReceivedDate) = True Then
        strAppend = strAppend + " And CLNReceivedDate >= '" & Format(txtCLNReceivedDate, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtCLNReceivedDate2)) > 0 And IsDate(txtCLNReceivedDate2) = True Then
        strAppend = strAppend + " And CLNReceivedDate <= '" & Format(txtCLNReceivedDate2, "mm/dd/yyyy") & "'"
    End If
        
    If Len(Trim(cboCLNpanel)) Then
        strAppend = strAppend + " AND CLNPanel = '" & RemStr(Mid(cboCLNpanel, 1, 1)) & "'"
    End If
    
    If Len(Trim(cboCLNType)) Then
        strAppend = strAppend + " AND CLNType = '" & RemStr(Mid(cboCLNType, 1, 1)) & "'"
    End If
    
    If Len(Trim(cboCLNxray)) Then
        strAppend = strAppend + " AND CLNXRay = '" & RemStr(Mid(cboCLNxray, 1, 1)) & "'"
    End If
    
    If Len(Trim(txtCLNformema)) Then
        strAppend = strAppend + " And CLNForRef = '" & Trim(txtCLNformema) & "'"
    End If
    
    If Len(Trim(TxtLocation)) Then
        'strAppend = strAppend + " And CLNLocation = '" & Trim(TxtLocation) & "'"
        strAppend = strAppend + " and ((CLNAdd1 LIKE '%" & TxtLocation & "%') OR" & _
                     " (CLNAdd2 LIKE '%" & TxtLocation & "%') OR " & _
                     " (CLNAdd3 LIKE '%" & TxtLocation & "%'))"
    End If
    
    If Len(Trim(cboInfoFFS)) Then
        strAppend = strAppend + " AND CLNFFSStatus = '" & RemStr(Mid(cboInfoFFS, 1, 1)) & "'"
    End If
    
    Call CLNListing(strAppend)
    
Screen.MousePointer = Default
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CLNListing(Optional strAppend As String)
Const METHOD_NAME = "CLNListing"
On Error GoTo debugs

Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim current As String
Dim total As String
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
        
    lstClinicList.ListItems.Clear
    current = 0
    total = 0
    strAppendCase = "182"
    sql = sql + strAppend
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "MBMTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    If rst1.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    Else
    
    Do
        Do Until rst1.EOF
            With rst1
                Set Record = lstClinicList.ListItems.Add(, , !CLNCode)
                If Trim(!CLNName) <> "" Then
                    Record.SubItems(1) = Trim(!CLNName)
                End If
                If Trim(!ClinicGrpCode) <> "" Then
                    Record.SubItems(2) = !ClinicGrpCode
                End If
                If Trim(!CLNinitial) <> "" Then
                    Record.SubItems(3) = !CLNinitial
                End If
                If Trim(!CLNType) <> "" Then
                    Record.SubItems(4) = !CLNType
                End If
                If Trim(!CLNBNFType) <> "" Then
                    Record.SubItems(5) = !CLNBNFType
                End If
                If Trim(!CLNAdd1) <> "" Then
                    Record.SubItems(6) = !CLNAdd1
                End If
                If Trim(!CLNAdd2) <> "" Then
                    Record.SubItems(7) = !CLNAdd2
                End If
                If Trim(!CLNAdd3) <> "" Then
                    Record.SubItems(8) = !CLNAdd3
                End If
                
                If Trim(!CLNPostCode) <> "" Then
                    Record.SubItems(9) = !CLNPostCode
                End If
                
                If Trim(!CLNCity) <> "" Then
                    Record.SubItems(10) = !CLNCity
                End If
                
                If Trim(!CLNState) <> "" Then
                    Record.SubItems(11) = !CLNState
                End If
                
                If Trim(!CLNReceivedDate) <> "" Then
                    Record.SubItems(12) = !CLNReceivedDate
                End If
                
                If Trim(!CLNPanel) <> "" Then
                    Record.SubItems(13) = !CLNPanel
                End If
                
                If Trim(!CLNXRay) <> "" Then
                    Record.SubItems(14) = !CLNXRay
                End If
                
                If Trim(!CLNForRef) <> "" Then
                    Record.SubItems(15) = !CLNForRef
                End If
                
                If Trim(!CLNNoofDoc) <> "" Then
                    Record.SubItems(16) = !CLNNoofDoc
                End If
                
                If Trim(!CLNPhone) <> "" Then
                    Record.SubItems(17) = !CLNPhone
                End If
                
                If Trim(!CLNFax) <> "" Then
                    Record.SubItems(18) = !CLNFax
                End If
                
                If Trim(!CLNEmail) <> "" Then
                    Record.SubItems(19) = !CLNEmail
                End If
                
                If Trim(!CLNHomepage) <> "" Then
                    Record.SubItems(20) = !CLNHomepage
                End If
                
                If Trim(!CLNBankAcc) <> "" Then
                    Record.SubItems(21) = !CLNBankAcc
                End If
                
                If Trim(!CLNBankName) <> "" Then
                    Record.SubItems(22) = !CLNBankName
                End If
                
                If Trim(!CLNBankBrn) <> "" Then
                    Record.SubItems(23) = !CLNBankBrn
                End If
                
                If Trim(!CLNAccPerson) <> "" Then
                    Record.SubItems(24) = !CLNAccPerson
                End If
                If Trim(!CLN24Hrs) <> "" Then
                    Record.SubItems(25) = !CLN24Hrs
                End If
                If Trim(!CLNRemark) <> "" Then
                    Record.SubItems(26) = Trim(!CLNRemark)
                End If
                If Trim(!CLNLocation) <> "" Then
                    Record.SubItems(27) = Trim(!CLNLocation)
                End If
                If Trim(!CLNFFSStatus) <> "" Then
                    Record.SubItems(28) = Trim(!CLNFFSStatus)
                End If
                current = current + 1
                lblCurrent = current
            End With
            rst1.MoveNext
            DoEvents
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    rst1.Close
    Set rst1 = Nothing
    End If
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    cmdExit.Enabled = True
    CmdClear.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "FrmTracking")
    Call CLNComboListing
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


Public Sub ExportListViewtoExcel(lvwList As Control)
Const METHOD_NAME = "ExportListViewtoExcel"
On Error GoTo debugs
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
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub ExportRecords(vntData As Variant, vntHeader As Variant, ws As Worksheet)
Const METHOD_NAME = "ExportListViewtoExcel"
On Error GoTo debugs
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
        ws.Rows(2).Select
        ws.Activate
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
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub OpenExcel(vntData As Variant, vntHeader As Variant)
    Const METHOD_NAME = "OpenExcel"
    On Error GoTo Err_OpenExcel
    Dim objExcel As Excel.Application
    Dim objWrkSht As Worksheet
    Dim X As Integer
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Workbooks.Add
    '========== Set Template ===============
        'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\DRCRListing\DRCRListing.XLT")
    '========== Set Template ===============
    Set objWrkSht = objExcel.ActiveWorkbook.Sheets(1)
    objExcel.Visible = True
    ExportRecords vntData, vntHeader, objWrkSht
    objExcel.Interactive = True
    Set objExlSht = Nothing
    Set objExcel = Nothing
Err_OpenExcel:

    If Err.Description <> "" Then
        Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
    End If
    Select Case Err
        Case 0
        Case 439
        MsgBox "You must have Microsoft Excel installed on your PC.", vbCritical, "Application Not Found"
        Case Else
        MsgBox Err & ": " & error, vbCritical, "OpenExcel Error"
    End Select
End Sub

Private Sub lstClinicList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
   

    Select Case ColumnHeader.Index
    Case 1
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstClinicList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstClinicList, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstClinicList, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
    Case 14
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstClinicList, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
    Case 18
        SortListView lstClinicList, ColumnHeader.Index, ldtNumber, m_blnDirection(18)
    Case 19
        SortListView lstClinicList, ColumnHeader.Index, ldtNumber, m_blnDirection(19)
    Case 20
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(21)
    Case 22
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(22)
    Case 23
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(23)
    Case 24
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(24)
    Case 25
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(25)
    Case 26
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(26)
    Case 27
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(27)
    Case 28
        SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(28)
    End Select
        
End Sub

Private Sub txtCLNReceivedDate_LostFocus()
    If IsDate(txtCLNReceivedDate) Then
    Else
        If txtCLNReceivedDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtCLNReceivedDate.SetFocus
        End If
    End If
End Sub

Private Sub txtCLNReceivedDate2_LostFocus()
    If IsDate(txtCLNReceivedDate2) Then
    Else
        If txtCLNReceivedDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtCLNReceivedDate2.SetFocus
        End If
    End If
End Sub
