VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewWSAjd 
   Caption         =   "Worksheet Adjustment"
   ClientHeight    =   8088
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8088
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   11775
      Begin VB.TextBox txtCASNumber 
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
         Left            =   960
         TabIndex        =   0
         Top             =   150
         Width           =   1380
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdSearchCase 
         Caption         =   "&Search"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lblClaimability 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10440
         TabIndex        =   17
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Claimability"
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
         Left            =   9600
         TabIndex        =   16
         Top             =   150
         Width           =   900
      End
      Begin VB.Label txtPolStatus 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6360
         TabIndex        =   15
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label48 
         Caption         =   "Pol Status"
         Height          =   255
         Left            =   5520
         TabIndex        =   14
         Top             =   150
         Width           =   795
      End
      Begin VB.Label Label9 
         Caption         =   "Case Status"
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
         Left            =   7560
         TabIndex        =   13
         Top             =   150
         Width           =   900
      End
      Begin VB.Label lblCASStatus 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8520
         TabIndex        =   12
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label46 
         Caption         =   "Case Num"
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
         TabIndex        =   11
         Top             =   150
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   15
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Worksheet Adjustment"
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
         Left            =   180
         TabIndex        =   7
         Top             =   60
         Width           =   6615
      End
   End
   Begin VB.TextBox txtAdd1 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAdd2 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   8160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtState 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Text            =   "Text4"
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   18
      Top             =   960
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   12510
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Invoices && Diagnosis"
      TabPicture(0)   =   "frmNewWSAjd.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdExit2"
      Tab(0).Control(1)=   "GeneralFrame"
      Tab(0).Control(2)=   "txtMemType"
      Tab(0).Control(3)=   "lblHLTcode"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&WorkSheet"
      TabPicture(1)   =   "frmNewWSAjd.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblCDNoteNo"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblCDNo"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label11"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label10"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label12"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblVisitLimit"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtINvDate"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Fgrid"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtBalLimit"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtINVNo"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame6"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame7"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame4"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Not Covered"
      TabPicture(2)   =   "frmNewWSAjd.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label16"
      Tab(2).Control(1)=   "FGridNCover"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Remarks"
      TabPicture(3)   =   "frmNewWSAjd.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame11"
      Tab(3).Control(1)=   "Frame5"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame4 
         Height          =   495
         Left            =   6240
         TabIndex        =   87
         Top             =   6480
         Width           =   5415
         Begin VB.CommandButton cmdExport 
            Caption         =   "Export"
            Height          =   255
            Left            =   1200
            TabIndex        =   92
            Top             =   165
            Width           =   1095
         End
         Begin VB.CommandButton cmdApprove 
            Caption         =   "&Save WS"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   165
            Width           =   1095
         End
         Begin VB.CommandButton cmdExit 
            Cancel          =   -1  'True
            Caption         =   "Exit"
            Height          =   255
            Left            =   4320
            TabIndex        =   90
            Top             =   165
            Width           =   975
         End
         Begin VB.CommandButton cmdCredit 
            Caption         =   "Credit Note"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2280
            TabIndex        =   89
            Top             =   165
            Width           =   975
         End
         Begin VB.CommandButton cmdDebit 
            Caption         =   "Debit Note"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3240
            TabIndex        =   88
            Top             =   165
            Width           =   1095
         End
         Begin VB.CommandButton cmdExcLetter 
            Caption         =   "Exc Letter"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4080
            TabIndex        =   93
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Additional Excess"
         Height          =   540
         Left            =   120
         TabIndex        =   70
         Top             =   6500
         Visible         =   0   'False
         Width           =   3615
         Begin VB.OptionButton optExcessMedix 
            Caption         =   "By Medix"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optExcessPayor 
            Caption         =   "By Payor"
            Height          =   255
            Left            =   1200
            TabIndex        =   72
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optExcessCLN 
            Caption         =   "By Clinic"
            Height          =   255
            Left            =   2280
            TabIndex        =   71
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Limit Exceeded"
         Height          =   540
         Left            =   240
         TabIndex        =   66
         Top             =   6480
         Visible         =   0   'False
         Width           =   3495
         Begin VB.OptionButton optLimitMedix 
            Caption         =   "By Medix"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optLimitPayor 
            Caption         =   "By Payor"
            Height          =   255
            Left            =   1200
            TabIndex        =   68
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optLimitCLN 
            Caption         =   "By Clinic"
            Height          =   255
            Left            =   2160
            TabIndex        =   67
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtINVNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   65
         Top             =   360
         Width           =   6615
      End
      Begin VB.TextBox txtBalLimit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8880
         TabIndex        =   64
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdExit2 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   -64320
         TabIndex        =   63
         Top             =   6240
         Width           =   855
      End
      Begin VB.Frame GeneralFrame 
         Height          =   5775
         Left            =   -74880
         TabIndex        =   27
         Top             =   360
         Width           =   11535
         Begin VB.Frame Frame1 
            Caption         =   "Invoices"
            Height          =   1935
            Left            =   120
            TabIndex        =   82
            Top             =   3720
            Width           =   11335
            Begin MSComctlLib.ListView lstInvoice 
               Height          =   1575
               Left            =   120
               TabIndex        =   83
               Top             =   240
               Width           =   11175
               _ExtentX        =   19706
               _ExtentY        =   2773
               View            =   3
               LabelEdit       =   1
               Sorted          =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               AllowReorder    =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   38
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Index"
                  Object.Width           =   88
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Bordx"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Case No"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Version"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Inv No"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Inv Date"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "INV Amount"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "Pay To"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "Payee"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   9
                  Text            =   "Claim Type"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   10
                  Text            =   "Fol. Up Date"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   11
                  Text            =   "SPConsAmt"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   12
                  Text            =   "GPConsAmt"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   13
                  Text            =   "Medicine"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   14
                  Text            =   "Procedures"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   15
                  Text            =   "Lab"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   16
                  Text            =   "XRay"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   17
                  Text            =   "Dental"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   18
                  Text            =   "ChineseMed"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   19
                  Text            =   "TraditionalMed"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   20
                  Text            =   "PhyTheraphy"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   21
                  Text            =   "ACC"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   22
                  Text            =   "PRVCheckup"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   23
                  Text            =   "Vision"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   24
                  Text            =   "ID"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   25
                  Text            =   "TSSTO"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   26
                  Text            =   "Injection"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   27
                  Text            =   "Others1Desc"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   28
                  Text            =   "Others1"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   29
                  Text            =   "Others1Desc"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   30
                  Text            =   "NCSPfee"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   31
                  Text            =   "NCMedicine"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   32
                  Text            =   "NCDental"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   33
                  Text            =   "NCPhysio"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   34
                  Text            =   "INVPaidByMem"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   35
                  Text            =   "Maternity"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   36
                  Text            =   "PRVCheckup"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   37
                  Text            =   "Vision"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   " Diagnosis"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1680
            Left            =   120
            TabIndex        =   45
            Top             =   2040
            Width           =   11355
            Begin MSComctlLib.ListView lstFinalDiag 
               Height          =   1335
               Left            =   120
               TabIndex        =   46
               Top             =   240
               Width           =   11145
               _ExtentX        =   19664
               _ExtentY        =   2350
               View            =   3
               LabelEdit       =   1
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
                  Text            =   "Case Number"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Diag Desc"
                  Object.Width           =   5292
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "ICD Code"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.TextBox txtMBMNumber 
            Enabled         =   0   'False
            Height          =   285
            Left            =   8040
            TabIndex        =   32
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtMBMName 
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
            Height          =   285
            Left            =   1200
            TabIndex        =   31
            Top             =   240
            Width           =   3495
         End
         Begin VB.TextBox txtINSType 
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
            Height          =   285
            Left            =   5880
            TabIndex        =   29
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtPatientID 
            Enabled         =   0   'False
            Height          =   285
            Left            =   10900
            TabIndex        =   28
            Top             =   240
            Width           =   585
         End
         Begin VB.TextBox txtPatientName 
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
            Height          =   285
            Left            =   1200
            TabIndex        =   33
            Top             =   480
            Width           =   3495
         End
         Begin VB.TextBox txtClinic 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   34
            Top             =   720
            Width           =   3495
         End
         Begin VB.TextBox txtDisability 
            Enabled         =   0   'False
            Height          =   765
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   35
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txtCASRemark 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   36
            Top             =   1680
            Width           =   3495
         End
         Begin VB.TextBox txtPLNCode 
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
            Height          =   285
            Left            =   5880
            TabIndex        =   30
            Top             =   480
            Width           =   1065
         End
         Begin VB.TextBox TxtPayor 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5880
            TabIndex        =   37
            Top             =   720
            Width           =   1065
         End
         Begin VB.TextBox txtPolEffDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5880
            TabIndex        =   38
            Top             =   960
            Width           =   1065
         End
         Begin VB.TextBox txtExpDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5880
            TabIndex        =   39
            Top             =   1200
            Width           =   1065
         End
         Begin VB.TextBox txtICDCode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5880
            MultiLine       =   -1  'True
            TabIndex        =   40
            Top             =   1440
            Width           =   1065
         End
         Begin VB.TextBox txtVisitDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5880
            TabIndex        =   41
            Top             =   1680
            Width           =   1065
         End
         Begin VB.TextBox txtMBMPolNo 
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
            Height          =   285
            Left            =   8040
            TabIndex        =   42
            Top             =   480
            Width           =   3445
         End
         Begin VB.TextBox txtType 
            Enabled         =   0   'False
            Height          =   285
            Left            =   8040
            MultiLine       =   -1  'True
            TabIndex        =   43
            Top             =   720
            Width           =   3445
         End
         Begin VB.TextBox txtFolUpDate 
            Enabled         =   0   'False
            Height          =   765
            Left            =   8040
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   44
            Top             =   960
            Width           =   3445
         End
         Begin VB.Label Label2 
            Caption         =   "Mem Name"
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
            TabIndex        =   62
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label8 
            Caption         =   "Patient Name"
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
            TabIndex        =   61
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label Label5 
            Caption         =   "Plan Code"
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
            Left            =   4920
            TabIndex        =   60
            Top             =   480
            Width           =   795
         End
         Begin VB.Label Label47 
            Caption         =   "Eff Date"
            Height          =   255
            Left            =   4920
            TabIndex        =   59
            Top             =   975
            Width           =   1335
         End
         Begin VB.Label Label54 
            Caption         =   "Exp Date"
            Height          =   255
            Left            =   4920
            TabIndex        =   58
            Top             =   1215
            Width           =   1335
         End
         Begin VB.Label Label55 
            Caption         =   "Payor"
            Height          =   255
            Left            =   4920
            TabIndex        =   57
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label44 
            Caption         =   "Claim Types"
            Height          =   255
            Left            =   7080
            TabIndex        =   56
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label27 
            Caption         =   "Pol Num"
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
            Left            =   7080
            TabIndex        =   55
            Top             =   480
            Width           =   795
         End
         Begin VB.Label Label26 
            Caption         =   "Mem Num"
            Height          =   255
            Left            =   7080
            TabIndex        =   54
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "ICD Code"
            Height          =   255
            Left            =   4920
            TabIndex        =   53
            Top             =   1455
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Disability"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Fol-up Dates"
            Height          =   255
            Left            =   7080
            TabIndex        =   51
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label51 
            Caption         =   "Clinic"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label49 
            Caption         =   "DOV"
            Height          =   255
            Left            =   4920
            TabIndex        =   49
            Top             =   1695
            Width           =   1335
         End
         Begin VB.Label Label60 
            Caption         =   "Remark"
            Height          =   375
            Left            =   120
            TabIndex        =   48
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "INS Type"
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
            Left            =   4920
            TabIndex        =   47
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1860
         Left            =   -69000
         TabIndex        =   23
         Top             =   360
         Width           =   5655
         Begin VB.TextBox txtCDRemarks 
            Height          =   1455
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   5475
         End
         Begin VB.Label Label98 
            Caption         =   "Credit/Debit Note Remarks"
            Height          =   495
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1875
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   5775
         Begin VB.TextBox txtWSRemarks 
            Height          =   1455
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   5535
         End
         Begin VB.Label Label14 
            Caption         =   "Worksheet remarks"
            Height          =   495
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   2655
         End
      End
      Begin VB.TextBox txtMemType 
         Height          =   285
         Left            =   -74400
         TabIndex        =   19
         Top             =   6240
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid Fgrid 
         Height          =   5535
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   11535
         _ExtentX        =   20341
         _ExtentY        =   9758
         _Version        =   393216
         Rows            =   27
         Cols            =   8
         FixedCols       =   3
         BackColorFixed  =   -2147483638
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtINvDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   74
         Top             =   600
         Width           =   6615
      End
      Begin MSFlexGridLib.MSFlexGrid FGridNCover 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   84
         Top             =   600
         Width           =   3735
         _ExtentX        =   6583
         _ExtentY        =   3620
         _Version        =   393216
         Rows            =   8
      End
      Begin VB.Label lblVisitLimit 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8880
         TabIndex        =   86
         Top             =   630
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Visit Limit"
         Height          =   255
         Left            =   8040
         TabIndex        =   85
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Invoice Date"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Bal Limit"
         Height          =   255
         Left            =   8040
         TabIndex        =   80
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Invoice Num"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblCDNo 
         BackColor       =   &H0080FFFF&
         Height          =   255
         Left            =   10680
         TabIndex        =   77
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblHLTcode 
         BackColor       =   &H0080FFFF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   76
         Top             =   6240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Not Claimable Items"
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
         Left            =   -74760
         TabIndex        =   75
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblCDNoteNo 
         Caption         =   "Credit No :"
         Height          =   255
         Left            =   10680
         TabIndex        =   78
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmNewWSAjd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBlimit, ICUlimit, KidneyLimit, CashLimit, GuardianLimit, totalFMemOriginal As Double
Dim SurgicalLimit, AnaLimit, OrganTransLimit As Double
Dim TotalRB, TotalICU, TotalCash, TotalGuardian, totalKidney As Double
Dim TotalActual, totalApprove, totalCLN, totalMem, totalExceed, Excess As Double
Dim TotalFActual, totalFApprove, totalFCLN, totalFMem As Double
Dim BordxAmount As Double
Dim BalanceLimit, ExceededLimit, PreviousLimit As Double
Dim TotalReimburse, totalExcessMemNeed2Pay   As Double
Dim BoardTaxLmt, ICUTaxLmt, LodgerTaxLmt As Double
Dim BoardTaxAct, ICUTaxAct, LodgerTaxAct As Double
Dim BoardTaxApp, ICUTaxApp, LodgerTaxApp As Double
Dim BoardTaxHOS, ICUTaxHOS, LodgerTaxHOS As Double
Dim BoardTaxMBM, ICUTaxMBM, LodgerTaxMBM As Double
Dim InHosPhycisianLimit, KidneyDCLimit As Double
Dim KidneyMonths, lodgerday, daycash, PhyVisit As Double
Dim DB1, ICU1 As Double
Dim KidneyMonthsLmt, lodgerdayLmt, daycashLmt, PhyVisitLmt As Double
Dim DBDaysLmt, ICUDaysLmt As Double
Dim CoPayApp, CoPayMBM, TotalExcess As Double
Dim PLNCode As String
Dim ClickTExcess As Boolean
Dim ScanCode As String
Dim RBComAmt, RBCOMExcess As Double
Dim NewPLNCode As String
Dim NewPolEffDate, AdmissionDate, PolEffDate As String
Dim ClaimStatus As Boolean
Dim strsqlCLNSameLmt As String
Dim SameCLNLmt As New ADODB.Recordset
Dim CLNCode As String
Dim VISInd As Boolean
Dim ANNInd As Boolean

Dim CB1, CB2, CB3, CB4, CB5, CB6 As Integer
Dim CB7, CB8, CB9, CB10, CB11, CB12 As Integer
Dim CB13, CB14, CB15, CB16, CB17, CB18 As Integer


Dim I1, I2, I3, I4, I5, I6 As Integer
Dim I7, I8, I9, I10, I11, I12 As Integer
Dim I13, I14, I15, I16, I17, I18 As Integer


Dim ComPLNCode As String
Dim CLMPByM2M As Double
Dim FTotalMember As Double
Dim BNFStatus As Boolean
Dim ClaimOpen As String
Dim bExit As Boolean
Dim CLMVersionNo As Integer
Dim debitNo As String
Dim creditno As String
Dim HdebitNo As String
Dim HCreditNo As String
Dim MCreditNo As String
Dim bChangesMade As Boolean
Dim RBDays As String
Dim ICUDays As String
Dim CashDays, GuardianDays, InHosPhycisianDays, KidneyDCDays As String
Dim strsqlSameLmt As String
Dim SameLmt As New ADODB.Recordset
Dim EXCNumber As String
Dim HOSCode As String
Dim COPreHosp, COMPreHosp As Double
Dim PLNPercentage As String

' Approved
Dim App1, App2, App3, App4, App5 As String
Dim App6, App7, App8, App9, App10 As String
Dim App11, App12, App13, App14, App15 As String
Dim App16, App17, App18, App19, App20 As String
Dim App21, App22, App23, App24, APP25, APP26 As String
Dim APPCOPreHosp, APPCOMPreHosp As String

'CalReimburse
Dim RGPConsAmt As Double
Dim RSPConsAmt As Double
Dim RMedicine As Double
Dim RProcedures As Double
Dim RLab As Double
Dim RXRay As Double
Dim RMRI As Double
Dim RDental As Double
Dim RChineseMed As Double
Dim RTraditionalMed As Double
Dim RPhysioTheraphy As Double
Dim RAccident As Double
Dim RPRVCheckup As Double
Dim RVision As Double
Dim RID As Double
Dim RTSSTO As Double
Dim ROthers1 As Double
Dim ROthers2 As Double

'CalCashless
Dim CGPConsAmt As Double
Dim CSPConsAmt As Double
Dim CMedicine As Double
Dim CProcedures As Double
Dim CLab As Double
Dim CXRay As Double
Dim CMRI As Double
Dim CDental As Double
Dim CChineseMed As Double
Dim CTraditionalMed As Double
Dim CPhysioTheraphy As Double
Dim CAccident As Double
Dim CPRVCheckup As Double
Dim CVision As Double
Dim CID As Double
Dim CTSSTO As Double
Dim COthers1 As Double
Dim COthers2 As Double

'CalTotal
Dim TRBDay As Double
Dim TRBLmt As Double
Dim TICUDay As Double
Dim TICULmt As Double
Dim TGurdianDay As Double
Dim TGurdianLmt As Double
Dim TGorvDay As Double
Dim TGorvLmt As Double
Dim GenCB As Boolean


Private Sub BNFcombolisting()
Screen.MousePointer = vbHourglass

Dim subCat As New ADODB.Recordset
Dim Cnt As Integer
Call DoInitialSettings

'subCat.Open "select BNFSubCode , BNFSubNAme , CATDescription from view_bnf_cat_subcat1 order by BNFsubCode", medixmasterconn, adOpenKeyset, adLockOptimistic
'cboVersion.Clear
'cboVersion.AddItem "1"
'cboVersion.AddItem "Master"



'If cboVersion = "" Then
'    cboVersion.ListIndex = 0
'End If


    subCat.Open "select BNFCATCode , BNFCATDescription   from BNFCategory order by BNFCATCode", MediConnCISDB, adOpenKeyset, adLockOptimistic
    Cnt = 1
    Fgrid.Rows = subCat.RecordCount + 10
    
    Do Until subCat.EOF
        ' cboCatCode.AddItem subCat!BNFSubCode & " -- " & subCat!BNFSubName
         Fgrid.Col = 0
         Fgrid.Row = Cnt
         Fgrid.Text = subCat!BNFCatCode
         
         'Fgrid.Col = 1
         'Fgrid.Row = Cnt
         'Fgrid.Text = subCat!BNFCATDescription
    
         Fgrid.Col = 1
         Fgrid.Row = Cnt
         Fgrid.Text = subCat!BNFCATDescription
         Cnt = Cnt + 1
         subCat.MoveNext
    Loop
Screen.MousePointer = Default
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
' Benefits
Sub DoInitialSettings()
    Dim i, Y    As Integer
    Fgrid.MergeCol(1) = True     ' Allow merge on Columns 0 thru 3
    
    Fgrid.ColAlignment(1) = 1
    Fgrid.MergeCells = flexMergeRestrictColumns
    Fgrid.ColAlignment(3) = 7
    Fgrid.ColAlignment(4) = 7
    Fgrid.ColAlignment(5) = 7
    Fgrid.ColAlignment(6) = 7
    Fgrid.ColAlignment(7) = 7
    
    Fgrid.Row = 0
    'Fgrid.Font = Bold
    Fgrid.ColWidth(0) = 0     ' Set column's width
    Fgrid.ColWidth(1) = 2200
    Fgrid.ColWidth(2) = 2000
    Fgrid.ColWidth(3) = 1200
    Fgrid.ColWidth(4) = 1200
    Fgrid.ColWidth(5) = 1200
    Fgrid.ColWidth(6) = 1500
    Fgrid.ColWidth(7) = 1500

    Fgrid.Col = 1
    Fgrid = "Benefits"
    Fgrid.Col = 2
    Fgrid = "Days/Mths@Limit(MAX)"
    Fgrid.Col = 3
    Fgrid = "Limit"
    Fgrid.Col = 4
    Fgrid = "Actual"
    Fgrid.Col = 5
    Fgrid = "Approved"
    Fgrid.Col = 6
    Fgrid = "Payable to CLN"
    Fgrid.Col = 7
    Fgrid = "Due to/(frm)Mem"
    'Fgrid.Col = 7
    'Fgrid = "Due to/(frm)Mem"
    
    Fgrid.Col = 1
    Fgrid.Row = 19
    Fgrid = "Total"
    
    Fgrid.Row = 20
    Fgrid = "Co-Payment"
    
    Fgrid.Row = 21
    Fgrid = "Total Excess"

    Fgrid.Row = 22
    Fgrid = "Excess Collected By Clinic"
    
    Fgrid.Row = 23
    Fgrid = "Limit Exceeded"
    
    Fgrid.Row = 24

    Fgrid = "Clinic / Payor To Collect"
    
    Fgrid.Row = 25
    Fgrid = "Final Total"
    
    Fgrid.Row = 26
    Fgrid = "Bordereaux Amount"
    
    ' For Not covered Items
    FGridNCover.ColWidth(0) = 1800
    FGridNCover.ColWidth(1) = 1800
    FGridNCover.Col = 1
    FGridNCover.Row = 0
    FGridNCover = " Amount (RM) "
    FGridNCover.Col = 0
    FGridNCover.Row = 0
    FGridNCover = "Items "
    FGridNCover.Row = 1
    FGridNCover = "SP Fees "
    FGridNCover.Row = 2
    FGridNCover = "Medicine "
    FGridNCover.Row = 3
    FGridNCover = "Dental "
    FGridNCover.Row = 4
    FGridNCover = "Physio Theraphy "
    FGridNCover.Row = 5
    FGridNCover = "Maternity "
    FGridNCover.Row = 6
    FGridNCover = "Prev Checkup "
    FGridNCover.Row = 7
    FGridNCover = "Vision "

End Sub

Private Sub cmdApprove_Click()
On Error GoTo errSave
    Dim RecordSaved
    Dim startTrans As Boolean
    Dim strsql As String
    Dim strsqlFinDiag As String
    Dim FinDiag As New ADODB.Recordset
    Dim strsqlINV As String
    Dim INV As New ADODB.Recordset
   ' Dim CheckClaim As New ADODB.Recordset
    'strsql = "select CASNumber from Claim where CAsNUmber ='" & Trim(txtCASNumber) & "'"
    'CheckClaim.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'If CheckClaim.recordCount Then
     ' MsgBox "Worksheet for this case has been opened before!" & vbNewLine & _
      '       " Please use Claim adjustment for further editting!", vbCritical + vbOKOnly
    'Else
    strsqlFinDiag = "Select * from DIAFinal where CASNumber = '" & Trim(txtCASNumber) & "'"
    FinDiag.Open strsqlFinDiag, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        If Len(Trim(txtCASNumber)) = 0 Then
            MsgBox "Please enter Case Number!", vbOKOnly + vbCritical, DialogBoxTitle
            txtCASNumber.SetFocus
        ElseIf Len(Trim(txtMBMNumber)) = 0 Then
            MsgBox "Please select a record before proceeding!", vbOKOnly + vbCritical, DialogBoxTitle
            'txtMBMNumber.SetFocus
        ElseIf lstInvoice.ListItems.Count = 0 Then
            MsgBox "Worksheet cannot be saved without any invoice!", vbOKOnly + vbCritical, DialogBoxTitle
            SSTab1.Tab = 1
        'ElseIf FinDiag.RecordCount = 0 Then
        '    MsgBox "Worksheet cannot be saved without Final Diagnosis!", vbOKOnly + vbCritical, DialogBoxTitle
        '    SSTab1.Tab = 1
        Else
            RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
               Screen.MousePointer = vbHourglass
               startTrans = True
               medixmasterconnWS.beginTrans
               
               Call SaveClaim("A")
               Call saveLimit
               Call saveCLNLimit
               Call saveActual
               Call saveApprove
               Call saveCLN
               Call saveMBM
               
               ' strsqlINV = "select  * from ClaimInvoice where  INVPayTo = 'H' And CASNumber = '" & Trim(txtCASNumber) & "'"
               ' INV.Open strsqlINV, medixmasterconnWS, adOpenKeyset, adLockOptimistic
               ' If INV.RecordCount Then
                '    Call saveExcessLetter(INV!INVAmount)
                'End If
                'INV.Close
                'Set INV = Nothing

               'Call saveSameLimit
               Call UpdateInvoice
               'Call UpdateCasReserve
               Call UpdateMBMBalance
               'Call CalculateReim
               'Call CalculateCashless
               'Call CalculateTotal
               'lblWSNo = ChkStr(txtCASNumber) & "-" & VersionNo
                ' ready for debit or credit note
                If totalFMem < 0 Then
                    cmdDebit.Enabled = True
                    cmdCredit.Enabled = False
                    'lblCDNoteNo.Visible = True
                    'lblCDNo.Visible = True
                    'lblCDNoteNo.Caption = "Debit No : "
                    'If lblCDNo = "" Then
                    '    debitNo = GenerateDebitNo(TxtPayor)
                    '    lblCDNo = debitNo
                    '    Call UpdateCreditDebit(debitNo)
                    'End If
                ElseIf totalFMem > 0 Then
                    cmdDebit.Enabled = False
                    cmdCredit.Enabled = True
                    'lblCDNoteNo.Caption = "Credit No : "
                    'If lblCDNo = "" Then
                    '    creditno = GenerateCreditNo(TxtPayor)
                    '    lblCDNoteNo.Visible = True
                    '    lblCDNo.Visible = True
                    '    lblCDNo = creditno
                    '    Call UpdateCreditDebit(creditno)
                    'End If
                Else
                    cmdDebit.Enabled = False
                    cmdCredit.Enabled = False
                    lblCDNo.Visible = False
                    lblCDNoteNo.Visible = False
                End If
                        
               medixmasterconnWS.commitTrans
               cmdExport.Enabled = True
               'CmdCheque.Enabled = True
               cmdApprove.Enabled = False
               
               'Call HCreditDebit
               
               Screen.MousePointer = Default
               MsgBox "Record Save", vbOKOnly + vbInformation, DialogBoxTitle
              ' cmdSave.Enabled = False
                cmdApprove.Enabled = False
                Frame7.Enabled = False
                Frame6.Enabled = False
               
             End If
        End If
    FinDiag.Close
    Set FinDiag = Nothing
    Exit Sub
errSave:
    MsgBox Err.Description
    If startTrans = True Then
        medixmasterconnWS.RollbackTrans
    End If
End Sub

Private Sub cmdCredit_Click()
On Error Resume Next
  Dim objExcel As Excel.Application
  Dim objWorkbook As Excel.Workbook
  Dim objWorksheet As Excel.Worksheet
  'Dim creditno As String
    
    Screen.MousePointer = vbHourglass
    'Set objExcel = New excel.Application
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If
    
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

    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "dc\CreditNote.xlt")
    Set objWorksheet = objWorkbook.Sheets(1)
    'Turn off the alerts, otherwise user will have to confirm my actions.
    objExcel.DisplayAlerts = False
    'Do While objWorkbook.Worksheets.Count > 1
     '   Set objWorksheet = objWorkbook.Worksheets.Item(objWorkbook.Worksheets.Count)
      '  objWorksheet.Delete
    'Loop
    Set objWorksheet = objWorkbook.Worksheets.Item(1)
    objWorksheet.Cells(19, "E") = txtPatientName
    objWorksheet.Cells(21, "E") = txtMBMNumber
    objWorksheet.Cells(23, "E") = txtCASNumber
    objWorksheet.Cells(25, "E") = txtClinic
    objWorksheet.Cells(27, "E") = txtVisitDate
    objWorksheet.Cells(29, "E") = txtDisability
        
    'objWorksheet.Cells(29, "E") = txtINVNo
    'objWorksheet.Cells(31, "E") = txtDisability
    'objWorksheet.Cells(33, "E") = txtCDRemarks
    
    
    
    'CreditNo = GenerateCreditNo(TxtPayor)
    objWorksheet.Cells(10, "J") = lblCDNo
    objWorksheet.Cells(11, "J") = Date
    objWorksheet.Cells(19, "J") = totalFMem
    objWorksheet.Cells(10, "C") = txtMBMName
    objWorksheet.Cells(11, "C") = txtAdd1
    objWorksheet.Cells(12, "C") = txtAdd2
    objWorksheet.Cells(13, "C") = txtPhone
    'DebitCreditRefNo
    Call UpdateCreditDebit(creditno)
    cmdCredit.Enabled = False
    objExcel.DisplayAlerts = True
    'objWorksheet.Columns.AutoFit
    objExcel.Interactive = True
    Screen.MousePointer = vbDefault
    Exit Sub
ExitFunction:
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    objExcel.Quit
    Screen.MousePointer = vbDefault
    Exit Sub
HANDLE_ERROR:
    MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
    Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    GoTo ExitFunction
End Sub

Private Sub cmdDebit_Click()
On Error Resume Next
  Dim objExcel As Excel.Application
  Dim objWorkbook As Excel.Workbook
  Dim objWorksheet As Excel.Worksheet
  'Dim debitNo As String
    
    Screen.MousePointer = vbHourglass
    'Set objExcel = New excel.Application
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If
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

    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "dc\DebitNote.xlt")
    Set objWorksheet = objWorkbook.Sheets(1)
    'Turn off the alerts, otherwise user will have to confirm my actions.
    objExcel.DisplayAlerts = False
    'Do While objWorkbook.Worksheets.Count > 1
     '   Set objWorksheet = objWorkbook.Worksheets.Item(objWorkbook.Worksheets.Count)
      '  objWorksheet.Delete
    'Loop
    Set objWorksheet = objWorkbook.Worksheets.Item(1)
    objWorksheet.Cells(19, "E") = txtPatientName
    objWorksheet.Cells(21, "E") = txtMBMNumber
    objWorksheet.Cells(23, "E") = txtCASNumber
    objWorksheet.Cells(25, "E") = txtClinic
    objWorksheet.Cells(27, "E") = txtVisitDate
    objWorksheet.Cells(29, "E") = txtDisability
        
    'objWorksheet.Cells(29, "E") = txtINVNo
    'objWorksheet.Cells(31, "E") = txtDisability
    'objWorksheet.Cells(33, "E") = txtCDRemarks
    
    'debitNo = GenerateDebitNo(TxtPayor)
    objWorksheet.Cells(10, "J") = lblCDNo
    objWorksheet.Cells(11, "J") = Date
    objWorksheet.Cells(19, "J") = Format(totalFMem, "#,#.00")
    objWorksheet.Cells(10, "C") = txtMBMName
    objWorksheet.Cells(11, "C") = txtAdd1
    objWorksheet.Cells(12, "C") = txtAdd2
    objWorksheet.Cells(13, "C") = txtPhone
    'DebitCreditRefNo
    'Call UpdateCreditDebit(debitNo)
    cmdDebit.Enabled = False
    Screen.MousePointer = vbDefault
    'objWorksheet.Columns.AutoFit
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    Exit Sub
ExitFunction:
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    objExcel.Quit
    Screen.MousePointer = vbDefault
    Exit Sub
HANDLE_ERROR:
    MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
    Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    GoTo ExitFunction

End Sub

Private Sub cmdDisplay_Click()
'On Error GoTo ErrorShow
    Dim CAS As New ADODB.Recordset
    Dim strsql, sqlPolStatus As String
    Dim PolStatus As New ADODB.Recordset
    Dim OpenNew
    Dim strsqlannlmt As String
    Dim AnnLmt As New ADODB.Recordset
    Dim strsqlCOV As String
    Dim MBMCov As New ADODB.Recordset
    Dim strsqlCLM As String
    Dim CLM As New ADODB.Recordset
    Dim strsqlDiag As String
    Dim Diag As New ADODB.Recordset
    'cmdExcLetter.Enabled = False
    'PLNPercentage = ""
    'NewPLNCode = ""
    'PolEffDate = ""
    'AdmissionDate = ""

   ' cmdSave.Enabled = True
   'ClickTExcess = True
    strsql = "SELECT CASNumber, MBMStatus,  MBMNumber, PAycode, " & _
        " CoverID, MBMPolicyNo, MBMName, MBMCName , HLTCode, " & _
        " PLNCode, MBMPayorExpDate , CASStatus,Claimability," & _
        " MBMPayorEffDate , HLTCode ,ClinicCode," & _
        " MBMAvailableLimit,MBMCAvailableLimit, Claimability, CASStatus, " & _
        "  INSCOde ,  MBMAddress1, MBMAddress2, MBMMemberType, MBMCLNVisitLmt, MBMCLNBalLmt,MBMClinical," & _
        " MBMAddress3, MBMPostCode, CTYCode, STACode,DateVisit, DebitCreditRefNo,ANNLmtInd,VISLmtInd, " & _
        " MBMTelnoH, MBMTelNoM, MBMTelNoO, CLMPreviousBalance,CLMLimitExceeded, CLMWSRemarks"
    strsql = strsql & " from VIEW_MBM_CAS_Claim_WSAdjusment where 0=0"

'    If Trim(cboVersion) = "" Then
'        MsgBox "Please select a version nombor!", vbCritical
    If Trim(txtCASNumber) = "" Then
        MsgBox "Please select a Case Number ! ", vbCritical
        txtCASNumber.SetFocus
    Else
    
    strsql = strsql & " And CASNumber = '" & Trim(txtCASNumber) & "' "
    'If IsNumeric(cboVersion) Then
    '    strsql = strsql & " and CLMVersion = " & Trim(cboVersion)
    'End If

    'ScanCode = Trim(txtCASNumber)
    'If Mid(ScanCode, 1, 2) = "LO" Then
    '    txtCASNumber = ""
    '    Call SearchForm(ScanCode)
    'Else
        
        Screen.MousePointer = vbHourglass
        Call ClearEntries
        
        strsql = strsql & " And CASNumber = '" & Trim(txtCASNumber) & "'"
        CAS.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            If CAS.RecordCount Then
                'Do Until CAS.EOF
                'If CAS!PLNPercentage = "" Or txtSurLmt = "" Then.
                'If Len(CAS!MBMNewPolEffDate) Then
                '    NewPolEffDate = Format(CAS!MBMNewPolEffDate, "dd/mm/yyyy")
                'End If
                'If Len(CAS!MBMNewPlanCode) Then
                '    NewPLNCode = CAS!MBMNewPlanCode
                'End If
                    PolEffDate = CAS!MBMPayorEffDate
                    AdmissionDate = CAS!DateVisit
                    If Len(CAS!PLNCode) Then
                        PLNCode = CAS!PLNCode
                    End If
                    If CAS!ANNLmtInd = "N" Then
                        ANNInd = False
                    Else
                        ANNInd = True
                    End If
                    If CAS!VISLmtInd = "N" Then
                        VISInd = False
                    Else
                        VISInd = True
                    End If
                    
                    Call Display_Diagnosis("F", txtCASNumber)
                    
    '                Call Display_Hospital(txtCASNumber)
                    sqlPolStatus = " select pstatuscode ,pstatusDesc from PolicyStatus where PStatusCode ='" & Trim(CAS!MBMStatus) & "'"
                    PolStatus.Open sqlPolStatus, medixmasterconn, adOpenKeyset, adLockOptimistic
                    If PolStatus.RecordCount Then
                        txtPolStatus = PolStatus!pstatusDesc
                    End If
                    
                    PolStatus.Close
                    Set PolStatus = Nothing
                    
                    'If Len(CAS!MBMExpiredStatus) Then
                    '    lblEXPStatus = CAS!MBMExpiredStatus
                    'End If
                    If Len(CAS!HLTCode) Then
                        lblHLTcode = CAS!HLTCode
                    End If
                    If ChkStr(CAS!CASStatus) = "O" Then
                        lblCASStatus = "Open"
                    ElseIf ChkStr(CAS!CASStatus) = "C" Then
                        lblCASStatus = "Close"
                    ElseIf ChkStr(CAS!CASStatus) = "R" Then
                        lblCASStatus = "Repudiated"
                    End If
                    
                    If Len(CAS!CASNumber) Then
                        txtCASNumber = CAS!CASNumber
                    End If
                    If Len(CAS!MBMNumber) Then
                        txtMBMNumber = CAS!MBMNumber
                    End If
                    If Len(CAS!MBMName) Then
                        txtMBMName = CAS!MBMName
                    End If
                    If Len(CAS!MBMPolicyNo) Then
                        txtMBMPolNo = CAS!MBMPolicyNo
                    End If
                    If Len(CAS!PLNCode) Then
                        txtPLNCode = CAS!PLNCode
                    End If
                    If CAS!INSCode <> "" Then
                        txtINSType = CAS!INSCode
                    End If
                    
                    If Len(CAS!MBMPayorEffDate) Then
                        txtPolEffDate = CAS!MBMPayorEffDate
                    End If
                    If Len(CAS!MBMPayorExpDate) Then
                        txtExpDate = CAS!MBMPayorExpDate
                    End If
                    If Len(CAS!PAYCode) Then
                        TxtPayor = CAS!PAYCode
                    End If
                    
                    If Len(CAS!MBMAddress1) Or Len(CAS!MBMAddress2) Then
                        txtAdd1 = CAS!MBMAddress1 & CAS!MBMAddress2
                    End If
                    
                    txtAdd2 = CAS!MBMAddress3 & ", " & CAS!MBMPostCode & " " & CAS!CTYCode  '& ", " & CAS!STACode
                                                   
                    txtPhone = "(H)" & CAS!MBMTelnoH & "  (M)" & CAS!MBMTelNoM & "  (O)" & CAS!MBMTelNoO
                    If Len(CAS!STACode) Then
                        txtState = Trim(CAS!STACode)
                    End If
                    
                    txtPatientID = CAS!coverid
                    
                    If Len(CAS!coverid) Then
                        If CAS!coverid <> "00" Then
                            If Len(CAS!MBMCName) Then
                                txtPatientName = CAS!MBMCName
                            End If
                        Else
                            txtPatientName = txtMBMName
                        End If
                    End If
                     'p= separated  c = combined
                    If Len(CAS!MBMMemberType) Then
                        txtMemType = CAS!MBMMemberType
                    End If
                    
                    If ANNInd = True Then
                        strsqlannlmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(txtPLNCode) & "'"
                        AnnLmt.Open strsqlannlmt, medixmasterconn, adOpenKeyset, adLockOptimistic
        
                        If ChkStr(txtMemType) = "Q" Then
                            If AnnLmt!SuppLimitStatus = "B" And CAS!mbmpatientcovid <> "00" Then
                                strsqlCOV = "Select MBMCAvailableLimit from MBMCoveredPersons where MBMCnumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & "01" & "'"
                                MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
                                If Len(CAS!mbmpatientcovid) Then
                                        txtBalLimit = Format(MBMCov!MBMCAvailableLimit, "#,#.00")
                                MBMCov.Close
                                Set MBMCov = Nothing
                                AnnLmt.Close
                                Set AnnLmt = Nothing
                                End If
                            Else
                                If Len(CAS!mbmpatientcovid) Then
                                    If CAS!mbmpatientcovid <> "00" Then
                                        txtBalLimit = Format(CAS!MBMCAvailableLimit, "#,#.00")
                                    Else
                                        txtBalLimit = Format(CAS!MBMCLNBalLmt, "#,#.00")
                                    End If
                                End If
                            End If
                        Else
                            If Len(CAS!MBMCLNBalLmt) Then
                                txtBalLimit = Format(CAS!MBMCLNBalLmt, "#,#.00")
                            End If
                        End If

                        If Trim(CAS!CLMPreviousBalance) <> "" Then
                            PreviousLimit = CAS!CLMPreviousBalance
                        Else
                            PreviousLimit = 0
                        End If
                    End If
                    
                    If VISInd = True Then
                        lblVisitLimit = Format(CAS!MBMCLNVisitLmt, "#,#.00")
                    End If
                    
                    If CAS!CLMLimitExceeded <> "" Then
                        ExceededLimit = Format(CAS!CLMLimitExceeded, "#,#.00")
                    Else
                        ExceededLimit = 0
                    End If
                    If Trim(CAS!DebitCreditRefNo) <> "" Then
                       ' DebitCreditRef =
                        lblCDNo.Visible = True
                        lblCDNoteNo.Visible = True
                        lblCDNo = CAS!DebitCreditRefNo
                    Else
                        lblCDNo = ""
                        lblCDNo.Visible = False
                        lblCDNoteNo.Visible = False
                    End If
                    
                    If Len(CAS!CLMWSRemarks) Then
                        txtWSRemarks = CAS!CLMWSRemarks
                    End If
                    'If Trim(CAS!PayTypeCode) <> "" Then
                    '    txtType = CAS!PayTypeCode
                    'End If
                    If Len(CAS!DateVisit) Then
                        txtVisitDate = CAS!DateVisit
                    End If
                    If Len(Trim(CAS!ClinicCode)) <> 0 Then
                    '    txtHOS = FindHosName(CAS!CASHOSCode)
                        CLNCode = CAS!ClinicCode
                        txtClinic = FindCLNName(CAS!ClinicCode)
                    End If
                   If Len(CAS!Claimability) Then
                        If CAS!Claimability = "A" Then
                            lblClaimability = "Accepted"
                        Else
                            lblClaimability = "Repudiated"
                        End If
                    End If
                    'Call getVersion
                    Call showInvoice
            strsqlSameLmt = " select CASnumber ,  * from CLMLimit "
            strsqlSameLmt = strsqlSameLmt & " where  CASNumber = '" & Trim(txtCASNumber) & "'"
            SameLmt.Open strsqlSameLmt, MediConnCISWS, adOpenKeyset, adLockOptimistic
                If SameLmt.RecordCount Then
                    Call displaySameLmt
                Else
                    SameLmt.Close
                    Set SameLmt = Nothing
                strsqlDiag = " select PLNCode  from DiagNotPayable "
                strsqlDiag = strsqlDiag & " where  PLNCode = '" & Trim(txtPLNCode) & "'"
                Diag.Open strsqlDiag, MediConnCISDB, adOpenKeyset, adLockOptimistic
                
                    If Diag.RecordCount = 0 Then
                       Call display_benefit
                    End If
                    Diag.Close
                    Set Diag = Nothing
                    
                   'Call display_benefit
                End If
            
            'strsqlCLNSameLmt = " select CASnumber ,  * from CLMCLNLimit "
            'strsqlCLNSameLmt = strsqlCLNSameLmt & " where  CASNumber = '" & Trim(txtCASNumber) & "'"
            'SameCLNLmt.Open strsqlCLNSameLmt, MediConnCISWS, adOpenKeyset, adLockOptimistic
            '    If SameCLNLmt.RecordCount Then
            '        Call displaySameCLNLmt
            '    Else
            '        SameCLNLmt.Close
            '        Set SameCLNLmt = Nothing
                    
                    Call display_CLNbenefit
            '    End If
                    
                    Call CalculateReim
                    Call CalculateCashless
                    'Call CalculateTotal
                    Call getDatatoWS
            
            ' Get DR Note
            'strsqlCLM = " select CASnumber , DebitCreditRefNo,CLMPayableByMedix2M from Claim "
            'strsqlCLM = strsqlCLM & " where  CASNumber = '" & Trim(txtCASNumber) & "'and CLMVersion = '" & Version & "'"
            'CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
            '    If CLM.RecordCount Then
                    'If Len(CLM!DebitCreditRefNo) Then
            '            lblCDNo = CLM!DebitCreditRefNo
            '            If CLM!CLMPayableByMedix2M < 0 Then
            '                cmdDebit.Enabled = True
            '                cmdCredit.Enabled = False
            '            ElseIf CLM!CLMPayableByMedix2M > 0 Then
            '                cmdCredit.Enabled = True
            '                cmdDebit.Enabled = False
            '            ElseIf CLM!CLMPayableByMedix2M > 0 Then
            '                cmdCredit.Enabled = False
            '                cmdDebit.Enabled = False
            '                lblCDNo = ""
            '            End If
             '       End If
             '   End If
            'CLM.Close
            'Set CLM = Nothing
                   cmdApprove.Enabled = True
                    Frame7.Enabled = True
                    Frame6.Enabled = True
            Else
                MsgBox "Record not found! ", vbCritical
            End If
    End If
        
        'If cboVersion > 1 Then
        '    If Trim(ClaimOpen) = "C" Then
        '        MsgBox "This would be Worksheet No " & cboVersion
        '    Else
        '        OpenNew = MsgBox("There is an outstading worksheet! " & vbNewLine & _
        '                "Open outstanding Worksheet? ", vbYesNo)
        '        If OpenNew = vbYes Then
                   'If FormLoadedByName("frmNewWSAdjust") = True Then
                   '     frmNewWSAdjust.txtCASNumber = txtCASNumber
                   '     frmNewWSAdjust.cmdDisplay_Click
                   '     frmNewWSAdjust.SetFocus
                   ' Else
                   '     frmNewWSAdjust.txtCASNumber = txtCASNumber
                   '     frmNewWSAdjust.cmdDisplay_Click
                   '     frmNewWSAdjust.Show
                   ' End If
        '        End If
        '    End If
        'End If

    Screen.MousePointer = vbDefault
    Exit Sub
ErrorShow:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description

End Sub

Private Sub cmdExit_Click()
    Unload Me
    'Set frmNewWSReg = Nothing
End Sub

Private Sub cmdExport_Click()
On Error Resume Next
  Dim objExcel As Excel.Application
  Dim objWorkbook As Excel.Workbook
  Dim objWorksheet As Excel.Worksheet
  Dim creditno As String
    Screen.MousePointer = vbHourglass
    cmdExport.Enabled = False
    cmdApprove.Enabled = False
    'Set objExcel = New excel.Application
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If


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

    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\Worksheet.xls")
    Set objWorksheet = objWorkbook.Sheets(1)
    'Turn off the alerts, otherwise user will have to confirm my actions.
    objExcel.DisplayAlerts = False
    'Do While objWorkbook.Worksheets.Count > 1
     '   Set objWorksheet = objWorkbook.Worksheets.Item(objWorkbook.Worksheets.Count)
      '  objWorksheet.Delete
    'Loop
    Set objWorksheet = objWorkbook.Worksheets.Item(1)
    objWorksheet.Cells(5, "C") = txtMBMName
    objWorksheet.Cells(5, "I") = txtPLNCode
    objWorksheet.Cells(5, "K") = txtMBMNumber
    
    objWorksheet.Cells(6, "C") = txtPatientName
    objWorksheet.Cells(6, "I") = TxtPayor
    objWorksheet.Cells(6, "K") = txtCASNumber
  
    objWorksheet.Cells(7, "C") = txtClinic
    objWorksheet.Cells(7, "I") = txtPolEffDate
    
    objWorksheet.Cells(7, "K") = txtMBMPolNo
    
    objWorksheet.Cells(8, "C") = txtDisability
    objWorksheet.Cells(8, "I") = txtICDCode
    objWorksheet.Cells(8, "K") = "-"
    objWorksheet.Cells(9, "C") = txtWSRemarks
    objWorksheet.Cells(9, "I") = txtVisitDate
    objWorksheet.Cells(9, "K") = txtFolUpDate
    'objWorksheet.Cells(10, "I") = txtDisChargeDate
    
    objWorksheet.Cells(42, "D") = Logon
    objWorksheet.Cells(43, "J") = Date
    
    Call GetInvoiceToExcel(objWorksheet)
    Call InputMBMToExcel(objWorksheet)
    
    Set objWorksheet = objWorkbook.Sheets(2)
    Set objWorksheet = objWorkbook.Worksheets.Item(2)
    Call InputCActualToExcel(objWorksheet)
    If lstInvoice.SelectedItem.SubItems(7) = "C" Then
        Call InputMBMToExcel(objWorksheet)
    End If
    
    
    Set objWorksheet = objWorkbook.Sheets(3)
    Set objWorksheet = objWorkbook.Worksheets.Item(3)
    Call InputReimToExcel(objWorksheet)
    Call InputReimActualToExcel(objWorksheet)
    'If optExcessPayor = True Then
    '    objWorksheet.Application.ActiveSheet.optExcessPayor.value = True
    'ElseIf optExcessHos = True Then
    '   objWorksheet.Application.ActiveSheet.optExcessHos.value = True
    'Else
    '   objWorksheet.Application.ActiveSheet.optExcessMedix.value = True
    'End If
    
    'If optLimitPayor = True Then
    '   objWorksheet.Application.ActiveSheet.optLimitPayor.value = True
    'ElseIf optLimitHos = True Then
    '   objWorksheet.Application.ActiveSheet.optLimitHos.value = True
    'Else
    '  objWorksheet.Application.ActiveSheet.optLimitMedix.value = True
    'End If

    
    Screen.MousePointer = vbDefault
    cmdExport.Enabled = True
    'objWorksheet.Columns.AutoFit
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    cmdExport.Enabled = True
    Exit Sub
ExitFunction:
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    Screen.MousePointer = vbDefault
    objExcel.Quit
    Exit Sub
HANDLE_ERROR:
    MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
    Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    GoTo ExitFunction

End Sub

Private Sub FGrid_Click()
    On Error Resume Next
    Dim temp, tempCol, tempRow, tempTitle As String
    Dim Appvalue, PayHosValue, ActualValue As Double
    Dim Col6, Row22 As Integer
    
    If (Fgrid.Col = 3 And Fgrid.Row <= 17) Or (Fgrid.Col = 6 And Fgrid.Row = 22) And Fgrid.Row <> 0 Then
        'If (Fgrid.Col = 4 And Fgrid.Row <> 5) Then
    
            If Fgrid.Col = 6 And Fgrid.Row = 22 Then
                 Col6 = "6"
                 Row22 = "22"
                 ClickTExcess = True
            Else
                ClickTExcess = False
            End If
            tempCol = Fgrid.Col
            Fgrid.Col = 1
            tempTitle = Fgrid.Text
            temp = InputBox(tempTitle, Fgrid.Text)
            Fgrid.Col = tempCol
            
            
            If temp <> "" Then
                Excess = temp
                If Col6 = 6 And Row22 = 22 Then
                    temp = temp * -1
                End If

                If IsNumeric(temp) = True Then
                    Fgrid = Format(temp, "#,#0.00;(#,#0.00)")
                Else
                    Fgrid = UCase(temp)
                End If
                
                If ClickTExcess = True Then
                    Fgrid.Col = 7
                    If IsNumeric(temp) = True Then
                        Fgrid = Format(temp * -1, "#,#0.00;(#,#0.00)")
                    Else
                        Fgrid = UCase(temp)
                    End If
                End If
                
                Fgrid.Col = 4
                If IsNumeric(Fgrid.Text) Then
                    ActualValue = Fgrid.Text
                Else
                    ActualValue = 0
                End If
                
                If Col6 = 6 And Row22 = 22 Then
                    temp = temp * -1
                End If

                Fgrid.Col = 5
                If IsNumeric(temp) = True Then
                    Fgrid.Text = GetMin(CDbl(temp), CDbl(ActualValue))
                 Else
                    Fgrid.Text = Format(ActualValue, "#,#.00")
                End If
                

                
                Appvalue = Fgrid.Text
                Fgrid.Col = 6
                PayHosValue = Fgrid.Text
                
                If ClickTExcess = False Then
                    Fgrid.Col = 7
                    Fgrid.Text = DueToFromMBM((Appvalue), (PayHosValue))
                End If
                'Fgrid.Row = 29
                
                'Approve
                Fgrid.Col = 5
                Fgrid = getTotal(5)
                totalApprove = Fgrid
                
                Fgrid.Col = 6
                Fgrid = getTotal(6)
                totalCLN = Fgrid
                
                Fgrid.Col = 7
                Fgrid = getTotal(7)
                If IsNumeric(Fgrid) Then
                    totalMem = Fgrid
                Else
                    totalMem = 0
                End If
                TotalExcess = CDbl(totalMem) + CDbl(CoPayMBM) + Format(CDbl(Excess), "#,#0.00;(#,#0.00)")

                Call getFinalTotal
            End If
        End If
'    End If
End Sub
Private Function getTotal(CurCol As Integer)
    Dim X As Integer
    
    getTotal = 0
    Fgrid.Col = CurCol
    For X = 1 To 17
        Fgrid.Row = X
        If IsNumeric(Fgrid.Text) = True Then
            getTotal = getTotal + CDbl(Fgrid.Text)
            If Fgrid.Col = 8 Then
                If CDbl(Fgrid.Text) >= 0 Then
                   TotalReimburse = TotalReimburse + CDbl(Fgrid.Text)
                Else
                    totalExcessMemNeed2Pay = totalExcessMemNeed2Pay + CDbl(Fgrid.Text)
                End If
            End If
        End If
    Next
        
    If getTotal <> 0 Then
        getTotal = Format(getTotal, "#,#.00;(#,#.00)")
    Else
        getTotal = "-"
    End If
    Fgrid.Row = 19
    'Fgrid.Row = 25
    'Fgrid.Row = 24
End Function
Private Sub getFinalTotal()
    Dim FinalApprove As Double
    Dim curRow  As Integer
    If IsNumeric(totalExceed) = False Then
        totalExceed = 0
    End If
    If IsNumeric(totalApprove) = False Then
        totalApprove = 0
    End If
    If IsNumeric(txtBalLimit) = False Then
        txtBalLimit = 0
    End If
    
    
    Fgrid.Row = 25
    Fgrid.Col = 3
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    Fgrid = Format(PreviousLimit, "#,#.00;(#,#.00)")
    Fgrid.Col = 4
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    Fgrid = Format(TotalActual, "#,#.00;(#,#.00)")
    TotalFActual = Fgrid
    Fgrid.Col = 5
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    
    If ANNInd = False Then
        PreviousLimit = totalApprove
    End If
    
    If VISInd = False Then
        lblVisitLimit = totalApprove
    End If
            
    If CDbl(lblVisitLimit) > CDbl(PreviousLimit) Then
        If CDbl(totalApprove) > CDbl(PreviousLimit) Then
            Fgrid = Format(PreviousLimit, "#,#.00;(#,#.00)")
        Else
            If totalApprove <> 0 Then
                Fgrid = Format(CDbl(totalApprove) - CDbl(CoPayApp), "#,#.00;(#,#.00)")
            Else
                Fgrid = "-"
            End If
        End If
    ElseIf CDbl(lblVisitLimit) <= CDbl(PreviousLimit) Then
        If CDbl(totalApprove) >= CDbl(lblVisitLimit) Then
            Fgrid = Format(lblVisitLimit, "#,#.00;(#,#.00)")
        ElseIf CDbl(totalApprove) < CDbl(lblVisitLimit) Then
            Fgrid = Format(totalApprove, "#,#.00;(#,#.00)")
        End If
    End If
    
    'If totalApprove > CDbl(PreviousLimit) Then
    '    Fgrid = Format(PreviousLimit, "#,#.00;(#,#.00)")
    'Else
    '    If totalApprove <> 0 Then
    '        Fgrid = Format(CDbl(totalApprove) - CDbl(CoPayApp), "#,#.00;(#,#.00)")
    '    Else
    '        Fgrid = "-"
    '    End If
    'End If
    FinalApprove = IIf(IsNumeric(Fgrid) = True, Fgrid, 0)
    totalFApprove = FinalApprove
    
    
    ' Clinic
    Fgrid.Col = 6
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    
        
    If Excess <> 0 Then
        If Excess < 0 Then
            Excess = Excess * -1
        End If
        If totalCLN = Excess * -1 Then
            totalFCLN = Format(IIf(IsNumeric(totalCLN), totalCLN, 0), "#,#.00;(#,#.00)")
        ElseIf IsNumeric(totalCLN) = True And totalCLN > 0 Then
            totalFCLN = Format(IIf(IsNumeric(totalCLN), totalCLN, 0) - CDbl(Excess), "#,#.00;(#,#.00)")
        Else
            totalFCLN = 0
        End If
    Else
        If TotalExcess < 0 Then
            TotalExcess = TotalExcess * -1
        End If
        If IsNumeric(totalCLN) = True And totalCLN > 0 Then
            totalFCLN = Format(IIf(IsNumeric(totalCLN), totalCLN, 0) - CDbl(Excess), "#,#.00;(#,#.00)")
        Else
            totalFCLN = 0
        End If
        
    End If
    If totalFCLN <> 0 Then
        Fgrid = totalFCLN
    Else
        Fgrid = "-"
    End If
    
        ' Total Exceeded
    Fgrid.Row = 23
    'Fgrid.Row = 28
    Fgrid.Col = 5
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    
    If CoPayApp = 0 Then
        totalExceed = CDbl(FinalApprove) - totalApprove
        If totalExceed < 0 Then
            Fgrid = Format(totalExceed, "#,#.00;(#,#.00)")
            'Fgrid.Col = 7
            'Fgrid.CellFontBold = True
            'Fgrid.CellForeColor = vbBlue
            'Fgrid = Format(totalExceed, "#,#.00;(#,#.00)")
        Else
            Fgrid = "-"
            'Fgrid.Col = 7
            'Fgrid = "-"
        End If

        Fgrid.Col = 7
        Fgrid.CellFontBold = True
        Fgrid.CellForeColor = vbBlue
                
        If totalExceed < 0 Then
          Fgrid = Format(totalExceed, "#,#.00;(#,#.00)")
        Else
          Fgrid = "-"
        End If
    End If
    
    'TOTAL EXCESS
    Fgrid.Col = 7
    Fgrid.Row = 21
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    If IsNumeric(totalMem) = False Then
        totalMem = 0
    End If
    If IsNumeric(CoPayApp) = False Then
        CoPayApp = 0
    End If
    Fgrid = Format(CDbl(CoPayApp * -1) + CDbl(totalMem), "#,#.00;(#,#.00)")

    TotalExcess = CDbl(totalMem) + CDbl(CoPayMBM) + totalExceed
    
    
    Fgrid.Col = 7
    Fgrid.Row = 25
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue

    If TotalExcess > Excess Then
        If TotalExcess < 0 Then
            totalFMem = IIf(((TotalExcess * -1) = Excess), "0", Format(TotalExcess * -1, "#,#.00;(#,#.00)"))
        Else
            totalFMem = IIf(((TotalExcess * -1) = Excess), "0", Format(TotalExcess, "#,#.00;(#,#.00)"))
        End If
            TotalExcess = TotalExcess * -1
    Else
        If ClaimStatus = True Then
            If ClickTExcess = False Then
                totalFMem = IIf(((TotalExcess) = Excess), "0", Format(TotalExcess - (Excess * -1), "#,#.00;(#,#.00)"))
            Else
                totalFMem = Format(TotalExcess + Excess, "#,#.00;(#,#.00)")
            End If
        Else
            totalFMem = IIf(((TotalExcess * -1) = Excess), "0", Format(TotalExcess, "#,#.00;(#,#.00)"))
        End If
    End If
    
    If IsNumeric(totalMem) = True Then
        totalMem = CDbl(totalMem)
    Else
        totalMem = 0
    End If

    
    If totalFMem = 0 Then
        Fgrid = "-"
    Else
        Fgrid = Format(totalFMem, "#,#.00;(#,#.00)")
        If bChangesMade = False Then
            totalFMemOriginal = totalFMem
        End If
    End If
    
    ' totalFMem  = Total Payable 2 Member  by Medix
    ' ready for debit or credit note
    'If IsNumeric(cboVersion) = True And bChangesMade = False Then
        If totalFMem < 0 Then
            cmdDebit.Enabled = True
            cmdCredit.Enabled = False
        ElseIf totalFMem > 0 Then
            cmdDebit.Enabled = False
            cmdCredit.Enabled = True
        Else
            cmdDebit.Enabled = False
            cmdCredit.Enabled = False
       End If
    'End If

    TotalReimburse = 0
    totalExcessMemNeed2Pay = 0
    Fgrid.Col = 7
    
    For curRow = 1 To 17
        Fgrid.Row = curRow
        If IsNumeric(Fgrid.Text) = True Then
            If CDbl(Fgrid.Text) >= 0 Then
                TotalReimburse = TotalReimburse + CDbl(Fgrid.Text)
            Else
                
                    totalExcessMemNeed2Pay = totalExcessMemNeed2Pay + CDbl(Fgrid.Text)
            End If
        End If
    Next
   Call GetPaymentCollection
End Sub

' Calculation
Private Function GetMin(X As String, Y As String)
    If IsNumeric(Y) = True Then
        If CDbl(X) > CDbl(Y) Then
            GetMin = IIf(CDbl(Y) = 0, "-", Format(Y, "#,#.00"))
        Else
            GetMin = IIf(CDbl(X) = 0, "-", Format(X, "#,#.00"))
        End If
    Else
        If ChkStr(Y) = "FULL" Then
            GetMin = Format(X, "#,#.00")
        Else
            GetMin = 0
        End If
        
    End If
    If GetMin = 0 Then
        GetMin = "-"
    End If
End Function
Private Function GetMax(X As String, Y As String)
    If IsNumeric(Y) = True Then
        If CDbl(X) > CDbl(Y) Then
            GetMax = IIf(CDbl(X) = 0, "-", Format(Y, "#,#.00"))
        Else
            GetMax = IIf(CDbl(Y) = 0, "-", Format(Y, "#,#.00"))
        End If
    Else
        If ChkStr(Y) = "FULL" Then
            GetMax = Format(X, "#,#.00")
        Else
            GetMax = 0
        End If
        
    End If
    If GetMax = 0 Then
        GetMax = "-"
    End If
End Function
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

Private Sub Form_Load()
    Call ActionLogUpdate(12, "frmNewWSAjd")
    BNFcombolisting
    display_benefit
End Sub

Private Sub displaySameLmt()
Dim strsqlBNF As String
Dim BNF As New ADODB.Recordset
    Fgrid.Col = 3
        'If Trim(lblClaimability) = "Active" Then
    'If Len(NewPLNCode) Then
    '    If CDate(AdmissionDate) <= CDate(PolEffDate) Then
    '        strsqlBNF = "Select BNFCode , BNFNoOFDays, BNFClaimableAmount , BNFRMPerDis  , BNFEffDate from BNF  " & _
    '                " where PLnCode = '" & txtPLNCode & "' and INSCode = '" & Trim(txtINScode) & "' and HLTCode = '" & Trim(txtHLTcode) & "' And BNFEffDate <= '" & Format(PolEffDate, "mm/dd/yyyy") & "'  "
    '        strsqlBNF = strsqlBNF & " order by BNFCOde "
    '        BNF.Open strsqlBNF, medixmasterconn, adOpenKeyset, adLockOptimistic
    '    Else
    '        strsqlBNF = "Select BNFCode , BNFNoOFDays, BNFClaimableAmount , BNFRMPerDis  , BNFEffDate from BNF  " & _
    '                " where PLnCode = '" & NewPLNCode & "' and INSCode = '" & Trim(txtINScode) & "' and HLTCode = '" & Trim(txtHLTcode) & "' And BNFEffDate <= '" & Format(NewPolEffDate, "mm/dd/yyyy") & "' "
    '        strsqlBNF = strsqlBNF & " order by BNFCOde "
    '        BNF.Open strsqlBNF, medixmasterconn, adOpenKeyset, adLockOptimistic
    '    End If
    'Else
    '    strsqlBNF = "Select BNFCode , BNFNoOFDays, BNFClaimableAmount , BNFRMPerDis  , BNFEffDate from BNF  " & _
    '    " where PLnCode = '" & txtPLNCode & "' and INSCode = '" & Trim(txtINScode) & "' and HLTCode = '" & Trim(txtHLTcode) & "'"
    '    strsqlBNF = strsqlBNF & " order by BNFCOde "
    '    BNF.Open strsqlBNF, medixmasterconn, adOpenKeyset, adLockOptimistic
    'End If
    
    
    'strsqlBNF = "Select BNFCode , BNFNoOFDays, BNFClaimableAmount , BNFRMPerDis  , BNFEffDate from BNF  " & _
    '        " where PLnCode = '" & txtPLNCode & "' and INSCode = '" & Trim(txtINSCode) & "' and HLTCode = '" & Trim(txtHLTCode) & "' "
    'strsqlBNF = strsqlBNF & " order by BNFCOde "
    'BNF.Open strsqlBNF, medixmasterconn, adOpenKeyset, adLockOptimistic
    'Do Until BNF.EOF
    '     With BNF
    '        If !BNFCode = "101" Then
    '            RBlimit = Format(!BNFClaimAbleAmount, "#,#.00")
    '            If IsNumeric(!BNFNoOFDays) = False Then
    '                !BNFNoOFDays = 0
    '            End If
    '            RBDays = !BNFNoOFDays
    '            'Fgrid.Text = " @ " & Format(RBlimit, "#,#.00")
    '        ElseIf !BNFCode = "102" Then
    '            ICUlimit = !BNFClaimAbleAmount
    '            If IsNumeric(!BNFNoOFDays) = False Then
    '                !BNFNoOFDays = 0
     '           End If
    '            ICUDays = !BNFNoOFDays
    '            'Fgrid.Text = " @ " & Format(ICUlimit, "#,#.00")
    '        ElseIf !BNFCode = "601" Then
    '            CashLimit = Format(!BNFClaimAbleAmount, "#,#.00")
    '            CashDays = !BNFNoOFDays
    '        ElseIf !BNFCode = "701" Then
    '            GuardianLimit = Format(!BNFClaimAbleAmount, "#,#.00")
    '            GuardianDays = !BNFNoOFDays
    '            'Fgrid.Text = " @ " & Format(GuardianLimit, "#,#.00")
    '        ElseIf !BNFCode = "302" Then
    '            InHosPhycisianLimit = Format(!BNFRMperDis, "#,#.00")
    '            InHosPhycisianDays = !BNFNoOFDays
    '        ElseIf !BNFCode = "404" Then
     '           If !BNFRMperDis = "-" Then
     '               !BNFRMperDis = 0
     '           End If
     '           If IsNumeric(!BNFRMperDis) = True Then
      '              KidneyDCLimit = Format(!BNFRMperDis, "#,#.00")
      '          Else
      '              'If Len(!BNFRMperDis) Then
      '              KidneyDCLimit = 0
      '              'End If
      '          End If
      '          If IsNumeric(!BNFNoOFDays) = False Then
      '              !BNFNoOFDays = 0
      '          Else
      '              KidneyDCDays = !BNFNoOFDays
      '          End If
      '      End If
      '      'If !BNFCode = "602" Then
      '      '    Fgrid.Text = Format(!BNFRMperDis, "0.00%")
            
            'If !BNFCode = "602" Then
            '    Fgrid.Text = Format(!BNFRMperDis, "0.00%")
            'ElseIf !BNFNoOFDays <> "" And !BNFNoOFDays <> 0 Then
            '    If !BNFCode = "302" Then
            '        Fgrid.Text = Fgrid.Text & "@" & Format(!BNFRMperDis, "#,#.00") & "(" & !BNFNoOFDays & ")"
            '    ElseIf !BNFCode = "404" Then
            '        Fgrid.Text = Fgrid.Text & "@" & Format(!BNFRMperDis, "#,#.00") & "(" & !BNFNoOFDays & ")"
            '        KidneyLimit = !BNFRMperDis
            '    Else
            '.MoveNext
         'End With
    'Loop
    
   ' BNF.Close
   ' Set BNF = Nothing
   ' 'End If
    
    'strsqlSameLmt = "select  CASNUmber , * from ClaimSameLimit "
    'strsqlSameLmt = strsqlSameLmt & " where CASNumber = '" & Trim(txtCASNumber) & "' "
    'SameLmt.Open strsqlSameLmt, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    If Not SameLmt.EOF Then
        With SameLmt
            Fgrid.Row = 1
            Fgrid = Format(!CLMGPConsAmt, "#,#0.00")
            Fgrid.Row = 2
            Fgrid = Format(!CLMSPConsAmt, "#,#0.00")
            Fgrid.Row = 3
            Fgrid = Format(!CLMMedicine, "#,#0.00")
            Fgrid.Row = 4
            Fgrid = Format(!CLMProcedures, "#,#0.00")
            Fgrid.Row = 5
            Fgrid = Format(!CLMLab, "#,#0.00")
            Fgrid.Row = 6
            Fgrid = Format(!CLMXray, "#,#0.00")
            'Fgrid.Row = 7
            'Fgrid = Format(!CLMMRI, "#,#0.00")
            Fgrid.Row = 7
            Fgrid = Format(!CLMDental, "#,#0.00")
            Fgrid.Row = 8
            Fgrid = Format(!CLMChineseMed, "#,#0.00")
            Fgrid.Row = 9
            Fgrid = Format(!CLMTraditionalMEd, "#,#0.00")
            Fgrid.Row = 10
            Fgrid = Format(!CLMphysioTheraphy, "#,#0.00")
            Fgrid.Row = 11
            Fgrid = Format(!CLMAccident, "#,#0.00")
            Fgrid.Row = 12
            Fgrid = Format(!CLMPRVCheckup, "#,#0.00")
            Fgrid.Row = 13
            Fgrid = Format(!CLMVision, "#,#0.00")
            Fgrid.Row = 14
            Fgrid = Format(!CLMID, "#,#0.00")
            Fgrid.Row = 15
            Fgrid = Format(!CLMTSSTO, "#,#0.00")
            Fgrid.Row = 16
            Fgrid = Format(!CLMInjection, "#,#0.00")
            Fgrid.Row = 17
            Fgrid = Format(!CLMOTHERS1, "#,#0.00")
        
        End With
    End If
    SameLmt.Close
    Set SameLmt = Nothing
    
    
    If Trim(Mid(txtPLNCode, 1, 1)) = "1" Then
        If txtPLNCode = "11" Then
            ComPLNCode = "01"
        ElseIf txtPLNCode = "12" Then
            ComPLNCode = "02"
        ElseIf txtPLNCode = "13" Then
            ComPLNCode = "03"
        ElseIf txtPLNCode = "14" Then
            ComPLNCode = "04"
        ElseIf txtPLNCode = "15" Then
            ComPLNCode = "05"
        ElseIf txtPLNCode = "16" Then
            ComPLNCode = "06"
        End If
        
    End If
End Sub
Private Sub displaySameCLNLmt()
    If Not SameCLNLmt.EOF Then
        With SameCLNLmt
            CB1 = IIf(IsNumeric(!CLMGPConsAmt = True), Format(!CLMGPConsAmt, "#,#0.00"), 0)
            CB2 = IIf(IsNumeric(!CLMSPConsAmt = True), Format(!CLMSPConsAmt, "#,#0.00"), 0)
            CB3 = IIf(IsNumeric(!CLMMedicine = True), Format(!CLMMedicine, "#,#0.00"), 0)
            CB4 = IIf(IsNumeric(!CLMProcedures = True), Format(!CLMProcedures, "#,#0.00"), 0)
            CB5 = IIf(IsNumeric(!CLMLab = True), Format(!CLMLab, "#,#0.00"), 0)
            CB6 = IIf(IsNumeric(!CLMXray = True), Format(!CLMXray, "#,#0.00"), 0)
            CB7 = IIf(IsNumeric(!CLMDental = True), Format(!CLMDental, "#,#0.00"), 0)
            CB8 = IIf(IsNumeric(!CLMChineseMed = True), Format(!CLMChineseMed, "#,#0.00"), 0)
            CB9 = IIf(IsNumeric(!CLMTraditionalMEd = True), Format(!CLMTraditionalMEd, "#,#0.00"), 0)
            CB10 = IIf(IsNumeric(!CLMphysioTheraphy = True), Format(!CLMphysioTheraphy, "#,#0.00"), 0)
            CB11 = IIf(IsNumeric(!CLMAccident = True), Format(!CLMAccident, "#,#0.00"), 0)
            CB12 = IIf(IsNumeric(!CLMPRVCheckup = True), Format(!CLMPRVCheckup, "#,#0.00"), 0)
            CB13 = IIf(IsNumeric(!CLMVision = True), Format(!CLMVision, "#,#0.00"), 0)
            CB14 = IIf(IsNumeric(!CLMID = True), Format(!CLMID, "#,#0.00"), 0)
            CB15 = IIf(IsNumeric(!CLMTSSTO = True), Format(!CLMTSSTO, "#,#0.00"), 0)
            CB16 = IIf(IsNumeric(!CLMInjection = True), Format(!CLMInjection, "#,#0.00"), 0)
            CB17 = IIf(IsNumeric(!CLMOTHERS1 = True), Format(!CLMOTHERS1, "#,#0.00"), 0)
        End With
    End If
    SameCLNLmt.Close
    Set SameCLNLmt = Nothing
End Sub
Private Sub ClearEntries()
    Dim ctl As Control
    Dim FCol, Frow As Integer
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Name <> "txtCASNumber" Then
                ctl.Text = ""
            End If
        'ElseIf TypeOf ctl Is ComboBox Then
        '   If ctl.Name <> "cboVersion" Then
         '       ctl.Text = ""
          '  End If
        End If
    Next
    'TotalReimburse = 0
    'totalExcessMemNeed2Pay = 0
        
    'lstFinalDiag.ListItems.Clear
    lstInvoice.ListItems.Clear
    txtPolStatus = ""
    lblCASStatus = ""
    'lblWSNo = ""
    For FCol = 3 To 7
        Fgrid.Col = FCol
        'For Frow = 1 To 31
        For Frow = 1 To 24
            Fgrid.Row = Frow
            Fgrid = "-"
        Next
    Next
        For FCol = 2 To 6
        Fgrid.Col = FCol
        
        For Frow = 19 To 24
            Fgrid.CellFontBold = True
            Fgrid.CellForeColor = vbBlue
            Fgrid.Row = Frow
            Fgrid = "-"
        Next
    Next

    'lblHCDAmt.Visible = False
    'lblHCDAmtW.Visible = False
    'lblHCDNoW.Visible = False
    'lblHCDNo.Visible = False
    lblCDNo.Visible = False
    lblCDNoteNo.Visible = False
    cmdDebit.Enabled = False
    cmdCredit.Enabled = False
End Sub

Public Sub Display_Diagnosis(DIAStatus As String, CASNumber As String)
    Dim DiagnosisSql, DiagnosisFirst, DiagnosisFinal As String
    Dim lstDiag As ListItem
    Dim rst3 As New ADODB.Recordset
    Dim rstFinal As New ADODB.Recordset
    
    DiagnosisSql = "select  CASNumber , CASTreatReason1 , CASTreatReason2 ," & _
                        " CASTreatReason3,  ICDDescription1, ICDDescription2, ICDDescription3"
    DiagnosisSql = DiagnosisSql & " from VIEW_DiagnosisDesc where CASNUMBer ='" & Trim(CASNumber) & "'"
    
    rst3.Open DiagnosisSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    ' For Final
    lstFinalDiag.ListItems.Clear

     If rst3.RecordCount Then
        Set lstDiag = lstFinalDiag.ListItems.Add(, , rst3!CASNumber)
        If Trim(rst3!CASTreatReason1) <> "" Then
            With lstDiag
                If Trim(rst3!ICDDescription1) <> "" Then
                    .SubItems(1) = rst3!ICDDescription1
                End If
                If Trim(rst3!CASTreatReason1) <> "" Then
                    .SubItems(2) = Trim(rst3!CASTreatReason1)
                End If
                txtICDCode = rst3!CASTreatReason1
            End With
        End If
        If Trim(rst3!CASTreatReason2) <> "" Then
            Set lstDiag = lstFinalDiag.ListItems.Add(, , rst3!CASNumber)
            With lstDiag
                If Trim(rst3!ICDDescription2) <> "" Then
                    .SubItems(1) = Trim(rst3!ICDDescription2)
                End If
                If Trim(rst3!CASTreatReason2) <> "" Then
                    .SubItems(2) = Trim(rst3!CASTreatReason2)
                End If
                 txtICDCode = txtICDCode & ", " & rst3!CASTreatReason1
            End With
        End If
        If Trim(rst3!CASTreatReason3) <> "" Then
         Set lstDiag = lstFinalDiag.ListItems.Add(, , rst3!CASNumber)
         With lstDiag
         
             If Trim(rst3!ICDDescription3) <> "" Then
              .SubItems(1) = Trim(rst3!ICDDescription3)
             End If
         
             If Trim(rst3!CASTreatReason3) <> "" Then
              .SubItems(2) = Trim(rst3!CASTreatReason3)
             End If
             txtICDCode = txtICDCode & ", " & rst3!CASTreatReason1
         End With
        End If
    End If
    rst3.Close
    Set rst3 = Nothing
End Sub

Public Sub display_benefit()
    Dim BNF As New ADODB.Recordset
    Dim strsql As String
    Dim counter As Integer
    Dim i  As Integer
    Dim strsqlCOV As String
    Dim MBMCov As New ADODB.Recordset
    Dim PLNCode As String
    Dim INV As Integer
    Dim LMT As Integer
    Fgrid.Col = 3
    For i = 2 To 17
        Fgrid.Row = i
        Fgrid.Text = "-"
    Next
    
    'If Trim(lblClaimability) = "Accepted" Then
        'If Len(NewPLNCode) Then
        '    If CDate(AdmissionDate) <= CDate(PolEffDate) Then
        '        strsql = "Select BNFCode , BNFNoOFDays, BNFClaimableAmount , BNFRMPerDis  , BNFEffDate from BNF  " & _
        '                " where PLnCode = '" & txtPLNCode & "' and INSCode = '" & Trim(txtINSCode) & "' and HLTCode = '" & Trim(txtHLTCode) & "' And BNFEffDate <= '" & Format(PolEffDate, "mm/dd/yyyy") & "'  "
        '        strsql = strsql & " order by BNFCOde "
        '        BNF.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        '    Else
        '        strsql = "Select BNFCode , BNFNoOFDays, BNFClaimableAmount , BNFRMPerDis  , BNFEffDate from BNF  " & _
        '                " where PLnCode = '" & PLNCode & "' and INSCode = '" & Trim(txtINSCode) & "' and HLTCode = '" & Trim(txtHLTCode) & "' And BNFEffDate <= '" & Format(NewPolEffDate, "mm/dd/yyyy") & "' "
        '        strsql = strsql & " order by BNFCOde "
        '        BNF.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        '    End If
        'Else
            strsql = "Select BNFCode , BNFNoOFDays, BNFNoOfAnn, BNFRMPerVisit , BNFPercentage, BNFEffDate,BNFCalculation from BNF  " & _
            " where PLnCode = '" & txtPLNCode & "' and INSCode = '" & Trim(txtINSType) & "' and HLTCode = '" & Trim(lblHLTcode) & "'"
            strsql = strsql & " order by BNFCOde "
            BNF.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        'End If
        
        Do Until BNF.EOF
             With BNF
                For counter = 1 To 17 Step 1
                 Fgrid.Row = counter
                 Fgrid.Col = 0
                    If Fgrid.Text = !BNFCode Then
                        'Fgrid.Col = 1
                        'If !BNFClaimAbleAmount <> "" Then
                        '    Fgrid.Text = " @ " & Format(!BNFClaimAbleAmount, "#,#.00")
                        '    If !BNFCode = "101" Then
                        '        RBlimit = !BNFClaimAbleAmount
                        '        Fgrid.Text = " @ " & Format(RBlimit, "#,#.00")
                            'ElseIf !BNFCode = "102" Then
                            '    ICUlimit = !BNFClaimAbleAmount
                            '    Fgrid.Text = " @ " & Format(ICUlimit, "#,#.00")
                            'ElseIf !BNFCode = "601" Then
                            '    CashLimit = Format(!BNFClaimAbleAmount, "#,#.00")
                            '    CashDays = !BNFNoOfDays
                            'ElseIf !BNFCode = "701" Then
                                'GuardianLimit = !BNFClaimAbleAmount
                                'GuardianDays = !BNFNoOfDays
                                'Fgrid.Text = " @ " & Format(GuardianLimit, "#,#.00")
                            
                            'End If
                        'End If
    
                        'If !BNFCode = "602" Then
                        '    Fgrid.Text = Format(!BNFRMperDis, "0.00%")
                        'ElseIf !BNFNoOfDays <> "" And !BNFNoOfDays <> 0 Then
                        '    If !BNFCode = "302" Then
                        '        Fgrid.Text = Fgrid.Text & "@" & Format(!BNFRMperDis, "#,#.00") & "(" & !BNFNoOfDays & "" & "/" & "" & !BNFNoOfDays & " )"
                        '        InHosPhycisianDays = !BNFNoOfDays
                        '    ElseIf !BNFCode = "404" Then
                        '        Fgrid.Text = Fgrid.Text & "@" & Format(!BNFRMperDis, "#,#.00") & "(" & !BNFNoOfDays & "" & "/" & "" & !BNFNoOfDays & " )"
                        '        KidneyLimit = !BNFRMperDis
                        '        KidneyDCDays = !BNFNoOfDays
                        '    Else
                        '        If !BNFCode = "101" Then
                        '            RBDays = !BNFNoOfDays
                        '        ElseIf !BNFCode = "102" Then
                        '            ICUDays = !BNFNoOfDays
                        '        End If
                        '            Fgrid.Text = Fgrid.Text & "(" & !BNFNoOfDays & "" & "/" & "" & !BNFNoOfDays & ")"
                        '    End If
                        'End If
                        ' LImit col
                        Fgrid.Col = 3
                        'If !BNFRMPerVisit <> "" And IsNull(!BNFRMPerVisit) = False Then
                            'If !BNFCalculation = "1" Or !BNFCalculation = "2" Then
                                If !BNFCode = "001" Then
                                    If I1 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I1
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I1, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                ElseIf !BNFCode = "002" Then
                                    If I2 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I2
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I2, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                ElseIf !BNFCode = "003" Then
                                    If I3 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I3
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I3, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If

                                ElseIf !BNFCode = "004" Then
                                    If I4 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I14
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I4, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                
                                ElseIf !BNFCode = "005" Then
                                    If I5 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I5
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I5, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                        
                                ElseIf !BNFCode = "006" Then
                                    If I6 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I6
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I6, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                        
                                'ElseIf !BNFCode = "007" Then
                                '    If I7 > 0 Then
                                '        If IsNumeric(!BNFPercentage) = False Then
                                '            INV = 0
                                '        Else
                                '            INV = (I1 / !BNFPercentage) * 100
                                '        End If
                                '        If IsNumeric(!BNFRMpervisit) = False Then
                                '            LMT = 0
                                '        Else
                                '            LMT = !BNFRMpervisit
                                '        End If
                                '        If !BNFCalculation = "1" Then
                                '            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                '        ElseIf !BNFCalculation = "2" Then
                                '            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                '        ElseIf !BNFCalculation = "3" Then
                                '            Fgrid.Text = Format(I7, "#,##0.00")
                                '        ElseIf !BNFCalculation = "4" Then
                                '            Fgrid.Text = Format(LMT, "#,##0.00")
                                '        End If
                                '    Else
                                '        Fgrid.Text = "-"
                                '    End If
                                        
                                ElseIf !BNFCode = "007" Then
                                    If I7 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I7
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I7, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                    
                                ElseIf !BNFCode = "008" Then
                                    If I8 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I8
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I8, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If

                                ElseIf !BNFCode = "009" Then
                                    If I9 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I9
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I9, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                    
                                ElseIf !BNFCode = "010" Then
                                    If I10 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I10
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I10, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                        
                                ElseIf !BNFCode = "011" Then
                                    If I11 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I11
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I11, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                        
                                ElseIf !BNFCode = "012" Then
                                    If I12 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I12
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I12, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                    
                                ElseIf !BNFCode = "013" Then
                                    If I13 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I13
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If

                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I13, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                        
                                ElseIf !BNFCode = "014" Then
                                    If I14 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I14
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I14, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                        
                                ElseIf !BNFCode = "015" Then
                                    If I15 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I15
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I15, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                    
                                ElseIf !BNFCode = "016" Then
                                    If I16 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I16
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I16, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                        
                                ElseIf !BNFCode = "017" Then
                                    If I17 > 0 Then
                                        If IsNumeric(!BNFPercentage) = False Then
                                            INV = 0
                                        Else
                                            INV = (!BNFPercentage / 100) * I17
                                        End If
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            Fgrid.Text = Format(GetMin((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            Fgrid.Text = Format(GetMax((INV), (LMT)), "#,##0.00")
                                        ElseIf !BNFCalculation = "3" Then
                                            Fgrid.Text = Format(I17, "#,##0.00")
                                        ElseIf !BNFCalculation = "4" Then
                                            Fgrid.Text = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        Fgrid.Text = "-"
                                    End If
                                        

                                'Fgrid.Text = Format(!BNFRMPerVisit, "#,#.00")
                        ElseIf !BNFRMPerVisit = 0 Then
                            Fgrid.Text = "-"
                        End If
                        Exit For
                    'End If
                    'End If
                    End If
                Next
                .MoveNext
             End With
        Loop
        
        If BNF.RecordCount = 0 Then
            BNFStatus = False
        Else
            BNFStatus = True
        End If
        BNF.Close
        Set BNF = Nothing
        
End Sub
Private Sub showInvoice()
    Dim Invoice As New ADODB.Recordset
    Dim INVlist As ListItem
    Dim strsql As String
    I1 = 0
    I2 = 0
    I3 = 0
    I4 = 0
    I5 = 0
    I6 = 0
    I7 = 0
    I8 = 0
    I9 = 0
    I10 = 0
    I11 = 0
    I12 = 0
    I13 = 0
    I14 = 0
    I15 = 0
    I16 = 0
    I17 = 0
    I18 = 0

    strsql = "SELECt INVIndex, InvNo, " & _
         " InvDate,  INVAmount,  INVFolUpDate, " & _
        " INVPayee, INVPayto , CLMSPConsAmt, CLMMedicine, " & _
        " INVPaidByMem , CLMGPConsAmt, " & _
        "  CLMProcedures, CLMLab, CLMMedicine, CLMXRay, CLMMRI,INVBordxed," & _
        "  CLMDental, CLMChineseMed, CLMTraditionalMed, " & _
        " CLMPhysioTheraphy, CLMAccident, CLMPRVCheckup, CLMVision," & _
        "  CLMID, CLMTSSTO,CLMOthers1, CLMOthers1Desc,CLMOthers2, CLMOtherS2Desc,CASNUmber, " & _
        "  SPfee, Medicine,Dental, Physio,Maternity, PRVCheckup, Vision,INVWSStatus ,CLMInjection" & _
        " FROM View_Invoice where CASNUmber ='" & RemStr(txtCASNumber) & "'"
    
     'If IsNumeric(cboVersion) = True Then
     '   strsql = strsql & " AND CLMVersion =" & Trim(cboVersion)
    'Else
        'strsql = strsql & " and claimStatus = 'C' "
    'End If
    strsql = strsql & " and INVWSStatus = '1' "
    
    Invoice.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    If Invoice.RecordCount Then
        lstInvoice.ListItems.Clear
        Do Until Invoice.EOF
            Set INVlist = lstInvoice.ListItems.Add(, , Invoice!InvINdex)
            With INVlist
                If Trim(Invoice!INVBordxed) <> "" Then
                    .SubItems(1) = Invoice!INVBordxed
                End If
                
                If Trim(Invoice!CASNumber) <> "" Then
                    .SubItems(2) = Invoice!CASNumber
                End If
                
                'If Trim(Invoice!clmversion) <> "" Then
                '    .SubItems(3) = Invoice!clmversion
                'End If
                
                If Trim(Invoice!InvNo) <> "" Then
                    .SubItems(4) = Invoice!InvNo
                End If
                
                If Trim(Invoice!InvDate) <> "" Then
                    .SubItems(5) = Invoice!InvDate
                End If
                If Trim(Invoice!INVAmount) <> "" Then
                    .SubItems(6) = Invoice!INVAmount
                End If
                If Trim(Invoice!INVPayTo) <> "" Then
                    .SubItems(7) = Invoice!INVPayTo
                End If
                
                If Trim(Invoice!INVPayee) <> "" Then
                    .SubItems(8) = Invoice!INVPayee
                End If
                
                'If Trim(Invoice!ClaimType) <> "" Then
                '    .SubItems(9) = Invoice!ClaimType
                'End If
                
                If Trim(Invoice!INVFolUpDate) <> "" Then
                    .SubItems(10) = Invoice!INVFolUpDate
                End If
                
                If Trim(Invoice!CLMGPConsAmt) <> "" Then
                    .SubItems(11) = Invoice!CLMGPConsAmt
                    I1 = I1 + Invoice!CLMGPConsAmt
                End If
                
                If Trim(Invoice!CLMSPConsAmt) <> "" Then
                    .SubItems(12) = Invoice!CLMSPConsAmt
                    I2 = I2 + Invoice!CLMSPConsAmt
                End If
                
                If Trim(Invoice!CLMMedicine) <> "" Then
                    .SubItems(13) = Invoice!CLMMedicine
                    I3 = I3 + Invoice!CLMMedicine
                End If

                If Trim(Invoice!CLMProcedures) <> "" Then
                    .SubItems(14) = Invoice!CLMProcedures
                    I4 = I4 + Invoice!CLMProcedures
                End If
                If Trim(Invoice!CLMLab) <> "" Then
                    .SubItems(15) = Invoice!CLMLab
                    I5 = I5 + Invoice!CLMLab
                End If
                If Trim(Invoice!CLMXray) <> "" Then
                    .SubItems(16) = Invoice!CLMXray
                    I6 = I6 + Invoice!CLMXray
                End If
                'If Trim(Invoice!CLMMRI) <> "" Then
                '    .SubItems(17) = Invoice!CLMMRI
                '    I7 = I7 + Invoice!CLMMRI
                'End If
                If Trim(Invoice!CLMDental) <> "" Then
                    .SubItems(17) = Invoice!CLMDental
                    I7 = I7 + Invoice!CLMDental
                End If
                If Trim(Invoice!CLMChineseMed) <> "" Then
                    .SubItems(18) = Invoice!CLMChineseMed
                    I8 = I8 + Invoice!CLMChineseMed
                End If
                If Trim(Invoice!CLMTraditionalMEd) <> "" Then
                    .SubItems(19) = Invoice!CLMTraditionalMEd
                    I9 = I9 + Invoice!CLMTraditionalMEd
                End If
                If Trim(Invoice!CLMphysioTheraphy) <> "" Then
                    .SubItems(20) = Invoice!CLMphysioTheraphy
                    I10 = I10 + Invoice!CLMphysioTheraphy
                End If
                If Trim(Invoice!CLMAccident) <> "" Then
                    .SubItems(21) = Invoice!CLMAccident
                    I11 = I11 + Invoice!CLMAccident
                End If
                If Trim(Invoice!CLMPRVCheckup) <> "" Then
                    .SubItems(22) = Invoice!CLMPRVCheckup
                    I12 = I12 + Invoice!CLMPRVCheckup
                End If
                If Trim(Invoice!CLMVision) <> "" Then
                    .SubItems(23) = Invoice!CLMVision
                    I13 = I13 + Invoice!CLMVision
                End If
                If Trim(Invoice!CLMID) <> "" Then
                    .SubItems(24) = Invoice!CLMID
                    I14 = I14 + Invoice!CLMID
                End If
                If Trim(Invoice!CLMTSSTO) <> "" Then
                    .SubItems(25) = Invoice!CLMTSSTO
                    I15 = I15 + Invoice!CLMTSSTO
                End If
                If Trim(Invoice!CLMInjection) <> "" Then
                    .SubItems(26) = Invoice!CLMInjection
                    I16 = I16 + Invoice!CLMInjection
                End If
                'If Trim(Invoice!CLMOthers1Desc) <> "" Then
                '    .SubItems(27) = Invoice!CLMOthers1Desc
                'End If
                If Trim(Invoice!CLMOTHERS1) <> "" Then
                    .SubItems(28) = Invoice!CLMOTHERS1
                    I17 = I17 + Invoice!CLMOTHERS1
                End If
                If Trim(Invoice!CLMOthers1Desc) <> "" Then
                    .SubItems(29) = Invoice!CLMOthers1Desc
                End If
                
                'Items not covered
                If Trim(Invoice!SPFee) <> "" Then
                    .SubItems(30) = Invoice!SPFee
                End If
                If Trim(Invoice!Medicine) <> "" Then
                    .SubItems(31) = Invoice!Medicine
                End If
                If Trim(Invoice!Dental) <> "" Then
                    .SubItems(32) = Invoice!Dental
                End If
                If Trim(Invoice!Physio) <> "" Then
                    .SubItems(33) = Invoice!Physio
                End If
                If Trim(Invoice!INVPaidByMem) <> "" Then
                    .SubItems(34) = Invoice!INVPaidByMem
                End If
                If Trim(Invoice!Maternity) <> "" Then
                    .SubItems(35) = Invoice!Maternity
                End If
                If Trim(Invoice!PRVCheckup) <> "" Then
                    .SubItems(36) = Invoice!PRVCheckup
                End If
                If Trim(Invoice!Vision) <> "" Then
                    .SubItems(37) = Invoice!Vision
                End If
                
            End With
            Invoice.MoveNext
        Loop
    End If
    Invoice.Close
    Set Invoice = Nothing
End Sub

Private Sub CalculateCashless()
    Dim ListCount As Integer
    Dim ForLoop As Integer
    Dim curColReim As String
    Dim curColCashless As String
    
    CGPConsAmt = 0
    CSPConsAmt = 0
    CMedicine = 0
    CProcedures = 0
    CLab = 0
    CXRay = 0
    CMRI = 0
    CDental = 0
    CChineseMed = 0
    CTraditionalMed = 0
    CPhysioTheraphy = 0
    CAccident = 0
    CPRVCheckup = 0
    CVision = 0
    CID = 0
    CTSSTO = 0
    COthers1 = 0
    COthers2 = 0

    ListCount = lstInvoice.ListItems.Count
    For ForLoop = 1 To ListCount
        With lstInvoice.ListItems(ForLoop)
        If .SubItems(7) = "C" Then
            CGPConsAmt = CDbl(CGPConsAmt) + CDbl(IIf(IsNumeric(.SubItems(11)), .SubItems(11), 0))
            CSPConsAmt = CDbl(CSPConsAmt) + (IIf(IsNumeric(.SubItems(12)), .SubItems(12), 0))
            CMedicine = CDbl(CMedicine) + CDbl(IIf(IsNumeric(.SubItems(13)), .SubItems(13), 0))
            CProcedures = CDbl(CProcedures) + CDbl(IIf(IsNumeric(.SubItems(14)), .SubItems(14), 0))
            CLab = CDbl(CLab) + CDbl(IIf(IsNumeric(.SubItems(15)), .SubItems(15), 0))
            CXRay = CDbl(CXRay) + CDbl(IIf(IsNumeric(.SubItems(16)), .SubItems(16), 0))
            'CMRI = CDbl(CMRI) + CDbl(IIf(IsNumeric(.SubItems(17)), .SubItems(17), 0))
            CDental = CDbl(CDental) + CDbl(IIf(IsNumeric(.SubItems(17)), .SubItems(17), 0))
            CChineseMed = CDbl(CChineseMed) + CDbl(IIf(IsNumeric(.SubItems(18)), .SubItems(18), 0))
            CTraditionalMed = CDbl(CTraditionalMed) + CDbl(IIf(IsNumeric(.SubItems(19)), .SubItems(19), 0))
            CPhysioTheraphy = CDbl(CPhysioTheraphy) + CDbl(IIf(IsNumeric(.SubItems(20)), .SubItems(20), 0))
            CAccident = CDbl(CAccident) + CDbl(IIf(IsNumeric(.SubItems(21)), .SubItems(21), 0))
            CPRVCheckup = CDbl(CPRVCheckup) + CDbl(IIf(IsNumeric(.SubItems(22)), .SubItems(22), 0))
            CVision = CDbl(CVision) + CDbl(IIf(IsNumeric(.SubItems(23)), .SubItems(23), 0))
            CID = CDbl(CID) + CDbl(IIf(IsNumeric(.SubItems(24)), .SubItems(24), 0))
            CTSSTO = CDbl(CTSSTO) + CDbl(IIf(IsNumeric(.SubItems(25)), .SubItems(25), 0))
            COthers1 = CDbl(COthers1) + CDbl(IIf(IsNumeric(.SubItems(26)), .SubItems(26), 0))
            COthers2 = CDbl(COthers2) + CDbl(IIf(IsNumeric(.SubItems(28)), .SubItems(28), 0))
        End If
       End With
    Next
End Sub

Private Sub CalculateReim()
    Dim ListCount As Integer
    Dim ForLoop As Integer
    Dim curColReim As String
    Dim curColCashless As String
    
    RGPConsAmt = 0
    RSPConsAmt = 0
    RMedicine = 0
    RProcedures = 0
    RLab = 0
    RXRay = 0
    RMRI = 0
    RDental = 0
    RChineseMed = 0
    RTraditionalMed = 0
    RPhysioTheraphy = 0
    RAccident = 0
    RPRVCheckup = 0
    RVision = 0
    RID = 0
    RTSSTO = 0
    ROthers1 = 0
    ROthers2 = 0
    
    ListCount = lstInvoice.ListItems.Count
    For ForLoop = 1 To ListCount
        With lstInvoice.ListItems(ForLoop)
        If .SubItems(7) = "M" Then
            RGPConsAmt = CDbl(RGPConsAmt) + CDbl(IIf(IsNumeric(.SubItems(11)), .SubItems(11), 0))
            RSPConsAmt = CDbl(RSPConsAmt) + (IIf(IsNumeric(.SubItems(12)), .SubItems(12), 0))
            RMedicine = CDbl(RMedicine) + CDbl(IIf(IsNumeric(.SubItems(13)), .SubItems(13), 0))
            RProcedures = CDbl(RProcedures) + CDbl(IIf(IsNumeric(.SubItems(14)), .SubItems(14), 0))
            RLab = CDbl(RLab) + CDbl(IIf(IsNumeric(.SubItems(15)), .SubItems(15), 0))
            RXRay = CDbl(RXRay) + CDbl(IIf(IsNumeric(.SubItems(16)), .SubItems(16), 0))
            'RMRI = CDbl(CMRI) + CDbl(IIf(IsNumeric(.SubItems(17)), .SubItems(17), 0))
            RDental = CDbl(RDental) + CDbl(IIf(IsNumeric(.SubItems(17)), .SubItems(17), 0))
            RChineseMed = CDbl(RChineseMed) + CDbl(IIf(IsNumeric(.SubItems(18)), .SubItems(18), 0))
            RTraditionalMed = CDbl(RTraditionalMed) + CDbl(IIf(IsNumeric(.SubItems(19)), .SubItems(19), 0))
            RPhysioTheraphy = CDbl(RPhysioTheraphy) + CDbl(IIf(IsNumeric(.SubItems(20)), .SubItems(20), 0))
            RAccident = CDbl(RAccident) + CDbl(IIf(IsNumeric(.SubItems(21)), .SubItems(21), 0))
            RPRVCheckup = CDbl(RPRVCheckup) + CDbl(IIf(IsNumeric(.SubItems(22)), .SubItems(22), 0))
            RVision = CDbl(RVision) + CDbl(IIf(IsNumeric(.SubItems(23)), .SubItems(23), 0))
            RID = CDbl(RID) + CDbl(IIf(IsNumeric(.SubItems(24)), .SubItems(24), 0))
            RTSSTO = CDbl(RTSSTO) + CDbl(IIf(IsNumeric(.SubItems(25)), .SubItems(25), 0))
            ROthers1 = CDbl(ROthers1) + CDbl(IIf(IsNumeric(.SubItems(26)), .SubItems(26), 0))
            ROthers2 = CDbl(ROthers2) + CDbl(IIf(IsNumeric(.SubItems(28)), .SubItems(28), 0))
        End If
       End With
    Next
End Sub

Private Sub getDatatoWS()
    Dim strsql As String
    Dim CLM As New ADODB.Recordset
    Dim ForLoop, ListCount As Integer
    
    Dim A1, A2, A3, A4, A5 As Double
    Dim A6, A7, A8, A9, A10 As Double
    Dim A11, A12, A13, A14, A15 As Double
    Dim A16, A17, A18 As Double
    
    ' Payto Hospital
    Dim C1, C2, C3, C4, C5 As Double
    Dim C6, C7, C8, C9, C10 As Double
    Dim C11, C12, C13, C14, C15 As Double
    Dim C16, C17, C18  As Double
    Dim Tax As Double
    
    ' LIMIT
    Dim L1, L2, L3, L4, L5 As String
    Dim L6, L7, L8, L9, L10 As String
    Dim L11, L12, L13, L14, L15 As String
    Dim L16, L17, L18 As String
    
    Dim INVAmount, Access As Double
    Dim SPFee, Medicine, Dental, Physio, Maternity, PRVCheckup, Vision As Double
    
    Dim strsqlSameClm As String
    Dim SameClm As New ADODB.Recordset
    
    ListCount = lstInvoice.ListItems.Count
    txtINVNo = ""
    txtFolUpDate = ""
    txtINvDate = ""
    'OthersDesc1 = ""
    'OthersDesc2 = ""
    totalMem = 0
    'RoomAmount = "-"
    'ICUAmount = "-"
    'LodgerAmt = "-"
    'RMCash = "-"
    INVAmount = 0
    Access = 0
    ' Hospital
    DB1 = 0
    A1 = 0
    ICU1 = 0
    A2 = 0
    A3 = 0
    A4 = 0
    ' Surgical
    'Xray = 0
    'PreLab = 0
    'PreOthers = 0
    A5 = 0
    
    'SurgicalFee = 0
    'Physician = 0
    'PostTreat = 0
    A6 = 0
    A7 = 0
    
    ' Non Surgical
    A8 = 0
    A9 = 0
    A10 = 0
    PhyVisit = 0
    ' Out- pAtient
    A11 = 0
    A12 = 0
    A13 = 0
    KidneyMonths = 0
    A14 = 0
    
    ' Others
    A15 = 0
    A16 = 0
    lodgerday = 0
    A17 = 0
    A18 = 0
    
    ' Clinic
    C1 = 0
    C2 = 0
    C3 = 0
    C4 = 0
    C5 = 0
    C6 = 0
    C7 = 0
    C8 = 0
    C9 = 0
    C10 = 0
    C11 = 0
    C12 = 0
    C13 = 0
    C14 = 0
    C15 = 0
    C16 = 0
    C17 = 0
    C18 = 0
    
    ' Not COver
     SPFee = 0
     Medicine = 0
     Dental = 0
     Physio = 0
     Maternity = 0
     PRVCheckup = 0
     Vision = 0

     ClaimStatus = False
     ClickTExcess = False
    For ForLoop = 1 To ListCount
        With lstInvoice.ListItems(ForLoop)
          
            If Trim(.SubItems(4)) <> "" Then
                If Trim(txtINVNo) <> "" Then
                    txtINVNo = txtINVNo & ", " & .SubItems(4)
                Else
                    txtINVNo = .SubItems(4)
                End If
            End If
            If Trim(.SubItems(11)) <> "" Then
                If Trim(txtFolUpDate) <> "" Then
                    txtFolUpDate = txtFolUpDate & ", " & .SubItems(11)
                Else
                    txtFolUpDate = .SubItems(11)
                End If
            End If
            
            If Trim(.SubItems(5)) <> "" Then
                If Trim(txtINvDate) <> "" Then
                    txtINvDate = txtINvDate & "," & .SubItems(5)
                Else
                    txtINvDate = .SubItems(5)
                End If
            End If
            
            INVAmount = CDbl(INVAmount) + CDbl(IIf(IsNumeric(.SubItems(6)), .SubItems(6), 0))
            Access = IIf(IsNumeric(.SubItems(34)), .SubItems(34), 0)
            
            ' SPFee
            A1 = CDbl(A1) + CDbl(IIf(IsNumeric(.SubItems(11)), .SubItems(11), 0))
            ' GPFee
            A2 = CDbl(A2) + CDbl(IIf(IsNumeric(.SubItems(12)), .SubItems(12), 0))
            'Medication
            A3 = CDbl(A3) + CDbl(IIf(IsNumeric(.SubItems(13)), .SubItems(13), 0))
            ' Procedures
            A4 = CDbl(A4) + CDbl(IIf(IsNumeric(.SubItems(14)), .SubItems(14), 0))
            ' Lab
            A5 = CDbl(A5) + CDbl(IIf(IsNumeric(.SubItems(15)), .SubItems(15), 0))
            ' XRay/MRI
            A6 = CDbl(A6) + CDbl(IIf(IsNumeric(.SubItems(16)), .SubItems(16), 0))
            ' MRI
            'A6 = CDbl(A6) + CDbl(IIf(IsNumeric(.SubItems(16)), .SubItems(16), 0))
            ' Dental
            A7 = CDbl(A7) + CDbl(IIf(IsNumeric(.SubItems(17)), .SubItems(17), 0))
            ' Chinese Medicine
            A8 = CDbl(A8) + CDbl(IIf(IsNumeric(.SubItems(18)), .SubItems(18), 0))
            ' Traditonal Medicine
            A9 = CDbl(A9) + CDbl(IIf(IsNumeric(.SubItems(19)), .SubItems(19), 0))
            ' Phy Theraphy
            A10 = CDbl(A10) + CDbl(IIf(IsNumeric(.SubItems(20)), .SubItems(20), 0))
            ' Accident
            A11 = CDbl(A11) + CDbl(IIf(IsNumeric(.SubItems(21)), .SubItems(21), 0))
            ' PRV Checkup
            A12 = CDbl(A12) + CDbl(IIf(IsNumeric(.SubItems(22)), .SubItems(22), 0))
            ' Vision
            A13 = CDbl(A13) + CDbl(IIf(IsNumeric(.SubItems(23)), .SubItems(23), 0))
            ' ID
            A14 = CDbl(A14) + CDbl(IIf(IsNumeric(.SubItems(24)), .SubItems(24), 0))
            ' TSSTO
            A15 = CDbl(A15) + CDbl(IIf(IsNumeric(.SubItems(25)), .SubItems(25), 0))
            ' Other1
            A16 = CDbl(A16) + CDbl(IIf(IsNumeric(.SubItems(26)), .SubItems(26), 0))
            ' Other2
            A17 = CDbl(A17) + CDbl(IIf(IsNumeric(.SubItems(28)), .SubItems(28), 0))
             
            'A18 = CDbl(A18) + CDbl(IIf(IsNumeric(.SubItems(29)), .SubItems(29), 0))
             
            
            ' For paying to Clinic
           If .SubItems(7) = "C" Then
            ' SPFee
            C1 = CDbl(C1) + CDbl(IIf(IsNumeric(.SubItems(11)), .SubItems(11), 0))
            ' GPFee
            C2 = CDbl(C2) + CDbl(IIf(IsNumeric(.SubItems(12)), .SubItems(12), 0))
            ' Medication
            C3 = CDbl(C3) + CDbl(IIf(IsNumeric(.SubItems(13)), .SubItems(13), 0))
            ' Procedures
            C4 = CDbl(C4) + CDbl(IIf(IsNumeric(.SubItems(14)), .SubItems(14), 0))
            ' Lab
            C5 = CDbl(C5) + CDbl(IIf(IsNumeric(.SubItems(15)), .SubItems(15), 0))
            ' XRay
            C6 = CDbl(C6) + CDbl(IIf(IsNumeric(.SubItems(16)), .SubItems(16), 0))
                ' MRI
                'C6 = CDbl(C6) + CDbl(IIf(IsNumeric(.SubItems(16)), .SubItems(16), 0))
            ' Dental
            C7 = CDbl(C7) + CDbl(IIf(IsNumeric(.SubItems(17)), .SubItems(17), 0))
            ' Chinese Medicine
            C8 = CDbl(C8) + CDbl(IIf(IsNumeric(.SubItems(18)), .SubItems(18), 0))
            ' Traditonal Medicine
            C9 = CDbl(C9) + CDbl(IIf(IsNumeric(.SubItems(19)), .SubItems(19), 0))
            ' Phy Theraphy
            C10 = CDbl(C10) + CDbl(IIf(IsNumeric(.SubItems(20)), .SubItems(20), 0))
            ' Accident
            C11 = CDbl(C11) + CDbl(IIf(IsNumeric(.SubItems(21)), .SubItems(21), 0))
            ' PRV Checkup
            C12 = CDbl(C12) + CDbl(IIf(IsNumeric(.SubItems(22)), .SubItems(22), 0))
            ' Vision
            C13 = CDbl(C13) + CDbl(IIf(IsNumeric(.SubItems(23)), .SubItems(23), 0))
            ' ID
            C14 = CDbl(C14) + CDbl(IIf(IsNumeric(.SubItems(24)), .SubItems(24), 0))
            ' TSSTO
            C15 = CDbl(C15) + CDbl(IIf(IsNumeric(.SubItems(25)), .SubItems(25), 0))
            ' Other1
            C16 = CDbl(C16) + CDbl(IIf(IsNumeric(.SubItems(26)), .SubItems(26), 0))
            ' Other2
            C17 = CDbl(C17) + CDbl(IIf(IsNumeric(.SubItems(28)), .SubItems(28), 0))
             
            'C18 = CDbl(C18) + CDbl(IIf(IsNumeric(.SubItems(29)), .SubItems(29), 0))
        End If
            ' Not COver
            SPFee = CDbl(SPFee) + CDbl(IIf(IsNumeric(.SubItems(30)), .SubItems(30), 0))
            Medicine = CDbl(Medicine) + CDbl(IIf(IsNumeric(.SubItems(31)), .SubItems(31), 0))
            Dental = CDbl(Dental) + CDbl(IIf(IsNumeric(.SubItems(32)), .SubItems(32), 0))
            Physio = CDbl(Physio) + CDbl(IIf(IsNumeric(.SubItems(33)), .SubItems(33), 0))
            Maternity = CDbl(Maternity) + CDbl(IIf(IsNumeric(.SubItems(35)), .SubItems(35), 0))
            PRVCheckup = CDbl(PRVCheckup) + CDbl(IIf(IsNumeric(.SubItems(36)), .SubItems(36), 0))
            Vision = CDbl(Vision) + CDbl(IIf(IsNumeric(.SubItems(37)), .SubItems(37), 0))
            
        End With
    Next
    
    
    ' Limit col. for Days/Months
    Fgrid.Col = 3
    Fgrid.Row = 1
                        
    'strsqlSameClm = "Select * from ClaimSameLimit where CASNumber = '" & Trim(txtCASNumber) & "'"
    'SameClm.Open strsqlSameClm, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    'If SameClm.RecordCount Then
    '    If CDbl(SameClm!CLMBoardDays) > 0 Then
    '        If CDbl(SameClm!CLMBoardDays) > CDbl(DB1) Then
    '            Fgrid.Text = DB1 & Fgrid.Text
    '            If IsNumeric(RBlimit) = False Then
    '                RBlimit = 0
    '            End If
    '            TotalRB = DB1 * RBlimit
    '            DBDaysLmt = DB1
    '        ElseIf CDbl(SameClm!CLMBoardDays) <= CDbl(DB1) Then
            
    '            Fgrid.Text = SameClm!CLMBoardDays & Fgrid.Text
    '            TotalRB = SameClm!CLMBoardDays * RBlimit
    '            DBDaysLmt = SameClm!CLMBoardDays
    '        End If
    '    Else
    '        Fgrid.Text = "0" & Fgrid.Text
    '        If IsNumeric(RBlimit) = False Then
    '            RBlimit = 0
    '        End If
    '        TotalRB = "0" * RBlimit
    '        DBDaysLmt = "0"
    '    End If
    
    'Fgrid.Row = 2
    '    If CDbl(SameClm!CLMICUDays) > 0 Then
    '        If CDbl(SameClm!CLMICUDays) > CDbl(ICU1) Then
    '            Fgrid.Text = ICU1 & Fgrid.Text
    '            If IsNumeric(ICUlimit) Then
    '                TotalICU = ICU1 * ICUlimit
    '                ICUDaysLmt = ICU1
    '            Else
    '                TotalICU = 0
    '            End If
    '        ElseIf CDbl(SameClm!CLMICUDays) <= CDbl(ICU1) Then
    '            Fgrid.Text = SameClm!CLMICUDays & Fgrid.Text
    '            TotalICU = SameClm!CLMICUDays * ICUlimit
    '            ICUDaysLmt = SameClm!CLMICUDays
    '        End If
    '    Else
    '        Fgrid.Text = "0" & Fgrid.Text
    '        TotalICU = "0" * ICUlimit
    '        ICUDaysLmt = "0"
    '    End If
    
    'If Mid(PLNCode, 1, 1) <> "2" Then

    '    Fgrid.Row = 10
    '        If CDbl(SameClm!CLMINHOSPhyDays) > 0 Then
    '            If CDbl(SameClm!CLMINHOSPhyDays) > CDbl(PhyVisit) Then
    '                Fgrid.Text = PhyVisit & Fgrid.Text
    '            ElseIf CDbl(SameClm!CLMINHOSPhyDays) <= CDbl(PhyVisit) Then
    '                Fgrid.Text = SameClm!CLMINHOSPhyDays & Fgrid.Text
    '            End If
    '        Else
    '                Fgrid.Text = "0" & Fgrid.Text
    '        End If
    'End If
    
    'Fgrid.Row = 17
    '    If CDbl(SameClm!CLMKIDDCDays) > 0 Then
    '        If CDbl(SameClm!CLMKIDDCDays) > CDbl(KidneyMonths) Then
    '            If KidneyMonths <> 0 Then
    '                Fgrid.Text = KidneyMonths & Fgrid.Text
    '                KidneyMonthsLmt = KidneyMonths
    '                totalKidney = KidneyDCLimit * KidneyMonths
    '            Else
    '                KidneyMonthsLmt = KidneyMonths
    '            End If
    
    '        ElseIf CDbl(SameClm!CLMKIDDCDays) <= CDbl(KidneyMonths) Then
    '                Fgrid.Text = SameClm!CLMKIDDCDays & Fgrid.Text
    '                totalKidney = KidneyDCLimit * SameClm!CLMKIDDCDays
    '                KidneyMonthsLmt = SameClm!CLMKIDDCDays
    '        End If
    '    Else
    '        Fgrid.Text = "0" & Fgrid.Text
    '        totalKidney = "0" * KidneyLimit
    '        KidneyMonthsLmt = "0"
    '    End If
    
   ' Fgrid.Row = 18
   '     If CDbl(SameClm!CLMCashDays) > 0 Then
   '         If CDbl(SameClm!CLMCashDays) > CDbl(daycash) Then
   '             Fgrid.Text = daycash & Fgrid.Text
   '             If IsNumeric(CashLimit) = False Then
   '                 CashLimit = 0
   '             End If
   '             TotalCash = daycash * CashLimit
   '             daycashLmt = daycash
   '         ElseIf CDbl(SameClm!CLMCashDays) <= CDbl(daycash) Then
   '             Fgrid.Text = SameClm!CLMCashDays & Fgrid.Text
   '             TotalCash = SameClm!CLMCashDays * CashLimit
   '             daycashLmt = SameClm!CLMCashDays
   '         End If
   '     Else
   '         Fgrid.Text = "0" & Fgrid.Text
   '         If IsNumeric(CashLimit) = False Then
   '             CashLimit = 0
   '         End If
   '         TotalCash = "0" * CashLimit
   '         daycashLmt = "0"
   '     End If
    
   ' Fgrid.Row = 19
   '     If CDbl(SameClm!CLMCashDays) > 0 Then
   '         If CDbl(SameClm!CLMGurdianDays) > CDbl(daycash) Then
   '             Fgrid.Text = lodgerday & Fgrid.Text
   '             If IsNumeric(GuardianLimit) = False Then
   '                 GuardianLimit = 0
   '             End If
   '             TotalGuardian = lodgerday * GuardianLimit
   '
   '             lodgerdayLmt = lodgerday
   '         ElseIf CDbl(SameClm!CLMGurdianDays) <= CDbl(daycash) Then
   '             Fgrid.Text = SameClm!CLMGurdianDays & Fgrid.Text
   '             If GuardianLimit = "" Then
   '                 GuardianLimit = 0
   '             End If
   '                 TotalGuardian = SameClm!CLMGurdianDays * GuardianLimit
   '                 lodgerdayLmt = SameClm!CLMGurdianDays
   '         End If
   '     Else
   '         Fgrid.Text = "0" & Fgrid.Text
   '         If IsNumeric(GuardianLimit) = False Then
   '             GuardianLimit = 0
   '         End If
   '         TotalGuardian = "0" * GuardianLimit
   '         lodgerdayLmt = "0"
   '     End If
    'Else
    '    If BNFStatus = True Then
    '        If IsNumeric(RBDays) = False Then
    '            RBDays = 0
    '        End If
    '            If CDbl(RBDays) > CDbl(DB1) Then
    '                FGrid.Text = DB1 & FGrid.Text
    '                TotalRB = DB1 * RBlimit
    '                DBDaysLmt = DB1
    '            ElseIf CDbl(RBDays) <= CDbl(DB1) Then
    '                FGrid.Text = RBDays & FGrid.Text
    '                TotalRB = RBDays * RBlimit
    '                DBDaysLmt = RBDays
    '            End If
            
     '   FGrid.Row = 2
            'If ICUDays = "" Then
            '    ICUDays = 0
            'End If
            
            'If CDbl(ICUDays) > CDbl(ICU1) Then
            '    Fgrid.Text = ICU1 & Fgrid.Text
            '    ICUDaysLmt = ICU1
            '    If IsNumeric(ICUlimit) Then
            '        TotalICU = ICU1 * ICUlimit
            '    Else
            '        TotalICU = 0
            '    End If
            'ElseIf CDbl(ICUDays) <= CDbl(ICU1) Then
            '    Fgrid.Text = ICUDays & Fgrid.Text
            '    ICUDaysLmt = ICUDays
            '    If IsNumeric(ICUlimit) Then
            '        TotalICU = ICUDays * ICUlimit
            '    Else
            '        TotalICU = 0
            '    End If
            'End If
    'If Mid(PLNCode, 1, 1) <> "2" Then
    
     '   Fgrid.Row = 10
     '       If CDbl(InHosPhycisianDays) > CDbl(PhyVisit) Then
     '           Fgrid.Text = PhyVisit & Fgrid.Text
     '           PhyVisitLmt = PhyVisit
     '       ElseIf CDbl(InHosPhycisianDays) <= CDbl(PhyVisit) Then
     '           Fgrid.Text = InHosPhycisianDays & Fgrid.Text
     '           PhyVisitLmt = InHosPhycisianDays
     '       End If
    'End If
        'Fgrid.Row = 16
        '    If KidneyMonths <> 0 Then
        '        If CDbl(KidneyDCDays) > CDbl(KidneyMonths) Then
        '            Fgrid.Text = KidneyMonths & Fgrid.Text
        '            totalKidney = KidneyLimit * KidneyMonths
        '            KidneyMonthsLmt = KidneyMonths
        '        ElseIf CDbl(KidneyDCDays) <= CDbl(KidneyMonths) Then
        '            Fgrid.Text = KidneyDCDays & Fgrid.Text
        '            totalKidney = KidneyDCDays * KidneyLimit
        '            KidneyMonthsLmt = KidneyDCDays
        '        End If
        '    Else
        '        totalKidney = 0 * KidneyLimit
        '        KidneyMonthsLmt = 0
        '    End If
                    
        'Fgrid.Row = 18
        '    If CDbl(CashDays) > CDbl(daycash) Then
        '        Fgrid.Text = daycash & Fgrid.Text
        '        TotalCash = daycash * CashLimit
        '        daycashLmt = daycash
        '    ElseIf CDbl(CashDays) <= CDbl(daycash) Then
        '        Fgrid.Text = CashDays & Fgrid.Text
        '        TotalCash = CashDays * CashLimit
        '        daycashLmt = CashDays
        '    End If
        'Fgrid.Row = 19
        '    If CDbl(GuardianDays) > CDbl(lodgerday) Then
        '        Fgrid.Text = lodgerday & Fgrid.Text
        '        TotalGuardian = lodgerday * GuardianLimit
        '        lodgerdayLmt = lodgerday
        '    ElseIf CDbl(GuardianDays) <= CDbl(lodgerday) Then
        '        Fgrid.Text = GuardianDays & Fgrid.Text
        '        TotalGuardian = GuardianDays * GuardianLimit
        '        lodgerdayLmt = GuardianDays
        '    End If
        'End If
    
    'Limit col
    'Fgrid.Col = 3
    'Fgrid.Row = 1
    'Fgrid.Text = IIf(TotalRB <> 0, Format(TotalRB, "#,#.00"), "-")
    'Fgrid.Row = 2
    'If Trim(Fgrid.Text) = "-" Then
    '    Fgrid.Text = IIf(TotalICU <> 0, Format(TotalICU, "#,#.00"), "-")
    'End If
    'Fgrid.Row = 16
    '    Fgrid.Text = IIf(totalKidney <> 0, Format(totalKidney, "#,#.00"), "-")
    'Fgrid.Row = 18
     '   Fgrid.Text = IIf(TotalCash <> 0, Format(TotalCash, "#,#.00"), "-")
    'Fgrid.Row = 19
    '    Fgrid.Text = IIf((TotalGuardian) <> 0, Format(TotalGuardian, "#,#.00"), "-")
    'If Mid(PLNCode, 1, 1) <> "2" Then
    '    Fgrid.Row = 24
    '        Fgrid.Text = IIf((TotalRB) <> 0, Format(TotalRB * 0.05, "#,#.00"), "-")
    '
    '    Fgrid.Row = 25
    '        Fgrid.Text = IIf((TotalICU) <> 0, Format(TotalICU * 0.05, "#,#.00"), "-")
    '
    '    Fgrid.Row = 26
    '        Fgrid.Text = IIf((TotalGuardian) <> 0, Format(TotalGuardian * 0.05, "#,#.00"), "-")
    'End If
    'SameClm.Close
    'Set SameClm = Nothing
    
    Fgrid.Row = 1
    L1 = Fgrid.Text
    Fgrid.Row = 2
    L2 = Fgrid.Text
    Fgrid.Row = 3
    L3 = Fgrid.Text
    Fgrid.Row = 4
    L4 = Fgrid.Text
    Fgrid.Row = 5
    L5 = Fgrid.Text
    Fgrid.Row = 6
    L6 = Fgrid.Text
    Fgrid.Row = 7
    L7 = Fgrid.Text
    Fgrid.Row = 8
    L8 = Fgrid.Text
    Fgrid.Row = 9
    L9 = Fgrid.Text
    Fgrid.Row = 10
    L10 = Fgrid.Text
    Fgrid.Row = 11
    L11 = Fgrid.Text
    Fgrid.Row = 12
    L12 = Fgrid.Text
    Fgrid.Row = 13
    L13 = Fgrid.Text
    Fgrid.Row = 14
    L14 = Fgrid.Text
    Fgrid.Row = 15
    L15 = Fgrid.Text
    Fgrid.Row = 16
    L16 = Fgrid.Text
    Fgrid.Row = 17
    L17 = Fgrid.Text
    'Fgrid.Row = 18
    'L18 = Fgrid.Text
    
    'Actual Col.
    Fgrid.Col = 4
    Fgrid.Row = 1
    Fgrid.Text = IIf(A1 = 0, "-", Format(A1, "#,#.00"))
    Fgrid.Row = 2
    Fgrid.Text = IIf(A2 = 0, "-", Format(A2, "#,#.00"))
    Fgrid.Row = 3
    Fgrid.Text = IIf(A3 = 0, "-", Format(A3, "#,#.00"))
    Fgrid.Row = 4
    Fgrid.Text = IIf(A4 = 0, "-", Format(A4, " #,#.00"))
    Fgrid.Row = 5
    Fgrid.Text = IIf(A5 = 0, "-", Format(A5, "#,#.00"))
    Fgrid.Row = 6
    Fgrid.Text = IIf(A6 = 0, "-", Format(A6, "#,#.00"))
    Fgrid.Row = 7
    Fgrid.Text = IIf(A7 = 0, "-", Format(A7, "#,#.00"))
    Fgrid.Row = 8
    Fgrid.Text = IIf(A8 = 0, "-", Format(A8, "#,#.00"))
    Fgrid.Row = 9
    Fgrid.Text = IIf(A9 = 0, "-", Format(A9, "#,#.00"))
    Fgrid.Row = 10
    Fgrid.Text = IIf(A10 = 0, "-", Format(A10, "#,#.00"))
    Fgrid.Row = 11
    Fgrid.Text = IIf(A11 = 0, "-", Format(A11, "#,#.00"))
    Fgrid.Row = 12
    Fgrid.Text = IIf(A12 = 0, "-", Format(A12, "#,#.00"))
    Fgrid.Row = 13
    Fgrid.Text = IIf(A13 = 0, "-", Format(A13, "#,#.00"))
    Fgrid.Row = 14
    Fgrid.Text = IIf(A14 = 0, "-", Format(A14, "#,#.00"))
    Fgrid.Row = 15
    Fgrid.Text = IIf(A15 = 0, "-", Format(A15, "#,#.00"))
    Fgrid.Row = 16
    Fgrid.Text = IIf(A16 = 0, "-", Format(A16, "#,#.00"))
    Fgrid.Row = 17
    Fgrid.Text = IIf(A17 = 0, "-", Format(A17, "#,#.00"))
    'Fgrid.Row = 18
    'Fgrid.Text = IIf(A18 = 0, "-", Format(A18, "#,#.00"))
    
    ' approved Col'
    Fgrid.Col = 5
    Fgrid.Row = 1
    Fgrid.Text = GetMin((A1), (L1))
    App1 = Fgrid.Text
    Fgrid.Row = 2
    Fgrid.Text = GetMin((A2), (L2))
    App2 = Fgrid.Text
    Fgrid.Row = 3
    Fgrid.Text = GetMin((A3), (L3))
    App3 = Fgrid.Text
    Fgrid.Row = 4
    Fgrid.Text = GetMin((A4), (L4))
    App4 = Fgrid.Text
    Fgrid.Row = 5
    Fgrid.Text = GetMin((A5), (L5))
    App5 = Fgrid.Text
    Fgrid.Row = 6
    Fgrid.Text = GetMin((A6), (L6))
    App6 = Fgrid.Text
    Fgrid.Row = 7
    Fgrid.Text = GetMin((A7), (L7))
    App7 = Fgrid.Text
    Fgrid.Row = 8
    Fgrid.Text = GetMin((A8), (L8))
    App8 = Fgrid.Text
    Fgrid.Row = 9
    Fgrid.Text = GetMin((A9), (L9))
    App9 = Fgrid.Text
    Fgrid.Row = 10
    Fgrid.Text = GetMin((A10), (L10))
    App10 = Fgrid.Text
    Fgrid.Row = 11
    Fgrid.Text = GetMin((A11), (L11))
    App11 = Fgrid.Text
    Fgrid.Row = 12
    Fgrid.Text = GetMin((A12), (L12))
    App12 = Fgrid.Text
    Fgrid.Row = 13
    Fgrid.Text = GetMin((A13), (L13))
    App13 = Fgrid.Text
    Fgrid.Row = 14
    Fgrid.Text = GetMin((A14), (L14))
    App14 = Fgrid.Text
    Fgrid.Row = 15
    Fgrid.Text = GetMin((A15), (L15))
    App15 = Fgrid.Text
    Fgrid.Row = 16
    Fgrid.Text = GetMin((A16), (L16))
    App16 = Fgrid.Text
    Fgrid.Row = 17
    Fgrid.Text = GetMin((A17), (L17))
    App17 = Fgrid.Text
    'Fgrid.Row = 18
    'Fgrid.Text = GetMin((A18), (L18))
    'App18 = Fgrid.Text
    
    ' Clinic Col
    Fgrid.Col = 6
    If GenCB = True Then
        'Fgrid.Col = 6
        Fgrid.Row = 1
        Fgrid.Text = Format(C1, "#,#.00")
        Fgrid.Row = 2
        Fgrid.Text = Format(C2, "#,#.00")
        Fgrid.Row = 3
        Fgrid.Text = Format(C3, "#,#.00")
        Fgrid.Row = 4
        Fgrid.Text = Format(C4, " #,#.00")
        Fgrid.Row = 5
        Fgrid.Text = Format(C5, "#,#.00")
        Fgrid.Row = 6
        Fgrid.Text = Format(C6, "#,#.00")
        Fgrid.Row = 7
        Fgrid.Text = Format(C7, "#,#.00")
        Fgrid.Row = 8
        Fgrid.Text = Format(C8, "#,#.00")
        Fgrid.Row = 9
        Fgrid.Text = Format(C9, "#,#.00")
        Fgrid.Row = 10
        Fgrid.Text = Format(C10, "#,#.00")
        Fgrid.Row = 11
        Fgrid.Text = Format(C11, "#,#.00")
        Fgrid.Row = 12
        Fgrid.Text = Format(C12, "#,#.00")
        Fgrid.Row = 13
        Fgrid.Text = Format(C13, "#,#.00")
        Fgrid.Row = 14
        Fgrid.Text = Format(C14, "#,#.00")
        Fgrid.Row = 15
        Fgrid.Text = Format(C15, "#,#.00")
        Fgrid.Row = 16
        Fgrid.Text = Format(C16, "#,#.00")
        Fgrid.Row = 17
        Fgrid.Text = Format(C17, "#,#.00")
        'Fgrid.Row = 18
        'Fgrid.Text = IIf(C18 = 0, "-", Format(C18, "#,#.00"))
    Else
    
        Fgrid.Row = 1
        Fgrid.Text = GetMin((C1), (CB1)) 'IIf(C1 = 0, "-", Format(CB1, "#,#.00"))
        Fgrid.Row = 2
        Fgrid.Text = GetMin((C2), (CB2)) 'IIf(C2 = 0, "-", Format(CB2, "#,#.00"))
        Fgrid.Row = 3
        Fgrid.Text = GetMin((C3), (CB3)) 'IIf(C3 = 0, "-", Format(CB3, "#,#.00"))
        Fgrid.Row = 4
        Fgrid.Text = GetMin((C4), (CB4)) 'IIf(C4 = 0, "-", Format(CB4, " #,#.00"))
        Fgrid.Row = 5
        Fgrid.Text = GetMin((C5), (CB5)) 'IIf(C5 = 0, "-", Format(CB5, "#,#.00"))
        Fgrid.Row = 6
        Fgrid.Text = GetMin((C6), (CB6)) 'IIf(C6 = 0, "-", Format(CB6, "#,#.00"))
        Fgrid.Row = 7
        Fgrid.Text = GetMin((C7), (CB7)) 'IIf(C7 = 0, "-", Format(CB7, "#,#.00"))
        Fgrid.Row = 8
        Fgrid.Text = GetMin((C8), (CB8)) 'IIf(C8 = 0, "-", Format(CB8, "#,#.00"))
        Fgrid.Row = 9
        Fgrid.Text = GetMin((C9), (CB9)) 'IIf(C9 = 0, "-", Format(CB9, "#,#.00"))
        Fgrid.Row = 10
        Fgrid.Text = GetMin((C10), (CB10)) 'IIf(C10 = 0, "-", Format(CB10, "#,#.00"))
        Fgrid.Row = 11
        Fgrid.Text = GetMin((C11), (CB11)) 'IIf(C11 = 0, "-", Format(CB11, "#,#.00"))
        Fgrid.Row = 12
        Fgrid.Text = GetMin((C12), (CB12)) 'IIf(C12 = 0, "-", Format(CB12, "#,#.00"))
        Fgrid.Row = 13
        Fgrid.Text = GetMin((C13), (CB13)) 'IIf(C13 = 0, "-", Format(CB13, "#,#.00"))
        Fgrid.Row = 14
        Fgrid.Text = GetMin((C14), (CB14)) 'IIf(C14 = 0, "-", Format(CB14, "#,#.00"))
        Fgrid.Row = 15
        Fgrid.Text = GetMin((C15), (CB15)) 'IIf(C15 = 0, "-", Format(CB15, "#,#.00"))
        Fgrid.Row = 16
        Fgrid.Text = GetMin((C16), (CB16)) 'IIf(C16 = 0, "-", Format(CB16, "#,#.00"))
        Fgrid.Row = 17
        Fgrid.Text = GetMin((C17), (CB17)) 'IIf(C17 = 0, "-", Format(CB17, "#,#.00"))
        'Fgrid.Row = 18
        'Fgrid.Text = IIf(C18 = 0, "-", Format(C18, "#,#.00"))
    End If
    ' approved Due to From MEM
    Fgrid.Col = 7
    Fgrid.Row = 1
    'If C1 = 0 Then
    '    Fgrid.Text = "-"
    If lstInvoice.SelectedItem.SubItems(7) = "M" Or GenCB = True Then
           
        Fgrid.Row = 1
        Fgrid.Text = DueToFromMBM((App1), (C1))
             
        Fgrid.Row = 2
        Fgrid.Text = DueToFromMBM((App2), (C2))
        
        Fgrid.Row = 3
        Fgrid.Text = DueToFromMBM((App3), (C3))
        
        Fgrid.Row = 4
        Fgrid.Text = DueToFromMBM((App4), (C4))
        
        Fgrid.Row = 5
        Fgrid.Text = DueToFromMBM((App5), (C5))
        
        Fgrid.Row = 6
        Fgrid.Text = DueToFromMBM((App6), (C6))
        
        Fgrid.Row = 7
        Fgrid.Text = DueToFromMBM((App7), (C7))
        
        Fgrid.Row = 8
        Fgrid.Text = DueToFromMBM((App8), (C8))
        
        Fgrid.Row = 9
        Fgrid.Text = DueToFromMBM((App9), (C9))
        
        Fgrid.Row = 10
        Fgrid.Text = DueToFromMBM((App10), (C10))
        
        Fgrid.Row = 11
        Fgrid.Text = DueToFromMBM((App11), (C11))
        
        Fgrid.Row = 12
        Fgrid.Text = DueToFromMBM((App12), (C12))
        
        Fgrid.Row = 13
        Fgrid.Text = DueToFromMBM((App13), (C13))
        
        Fgrid.Row = 14
        Fgrid.Text = DueToFromMBM((App14), (C14))
        
        Fgrid.Row = 15
        Fgrid.Text = DueToFromMBM((App15), (C15))
        
        Fgrid.Row = 16
        Fgrid.Text = DueToFromMBM((App16), (C16))
        
        Fgrid.Row = 17
        Fgrid.Text = DueToFromMBM((App17), (C17))
    Else
        If CDbl(C1) < CDbl(CB1) Then
            Fgrid.Text = DueToFromMBM((App1), (C1))
        ElseIf CDbl(CB1) < CDbl(C1) Then
            Fgrid.Text = DueToFromMBM((App1), (CB1))
        End If
             
        Fgrid.Row = 2
        'If C2 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C2) < CDbl(CB2) Then
            Fgrid.Text = DueToFromMBM((App2), (C2))
        ElseIf CDbl(CB2) < CDbl(C2) Then
            Fgrid.Text = DueToFromMBM((App2), (CB2))
        End If
        
        Fgrid.Row = 3
        'If C3 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C3) < CDbl(CB3) Then
            Fgrid.Text = DueToFromMBM((App3), (C3))
        ElseIf CDbl(CB3) < CDbl(C3) Then
            Fgrid.Text = DueToFromMBM((App3), (CB3))
        End If
        
        Fgrid.Row = 4
        'If C4 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C4) < CDbl(CB4) Then
            Fgrid.Text = DueToFromMBM((App4), (C4))
        ElseIf CDbl(CB4) < CDbl(C4) Then
            Fgrid.Text = DueToFromMBM((App4), (CB4))
        End If
        
        Fgrid.Row = 5
        'If C5 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C5) < CDbl(CB5) Then
            Fgrid.Text = DueToFromMBM((App5), (C5))
        ElseIf CDbl(CB5) < CDbl(C5) Then
            Fgrid.Text = DueToFromMBM((App5), (CB5))
        End If
        
        Fgrid.Row = 6
        'If C6 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C6) < CDbl(CB6) Then
            Fgrid.Text = DueToFromMBM((App6), (C6))
        ElseIf CDbl(CB6) < CDbl(C6) Then
            Fgrid.Text = DueToFromMBM((App6), (CB6))
        End If
        
        Fgrid.Row = 7
        'If C7 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C7) < CDbl(CB7) Then
            Fgrid.Text = DueToFromMBM((App7), (C7))
        ElseIf CDbl(CB7) < CDbl(C7) Then
            Fgrid.Text = DueToFromMBM((App7), (CB7))
        End If
        
        Fgrid.Row = 8
        'If C8 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C8) < CDbl(CB8) Then
            Fgrid.Text = DueToFromMBM((App8), (C8))
        ElseIf CDbl(CB8) < CDbl(C8) Then
            Fgrid.Text = DueToFromMBM((App8), (CB8))
        End If
        
        Fgrid.Row = 9
        'If C9 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C9) < CDbl(CB9) Then
            Fgrid.Text = DueToFromMBM((App9), (C9))
        ElseIf CDbl(CB9) < CDbl(C9) Then
            Fgrid.Text = DueToFromMBM((App9), (CB9))
        End If
        
        Fgrid.Row = 10
        'If C10 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C10) < CDbl(CB10) Then
            Fgrid.Text = DueToFromMBM((App10), (C10))
        ElseIf CDbl(CB10) < CDbl(C10) Then
            Fgrid.Text = DueToFromMBM((App10), (CB10))
        End If
        
        Fgrid.Row = 11
        'If C11 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C11) < CDbl(CB11) Then
            Fgrid.Text = DueToFromMBM((App11), (C11))
        ElseIf CDbl(CB11) < CDbl(C11) Then
            Fgrid.Text = DueToFromMBM((App11), (CB11))
        End If
        
        Fgrid.Row = 12
        'If C12 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C12) < CDbl(CB12) Then
            Fgrid.Text = DueToFromMBM((App12), (C12))
        ElseIf CDbl(CB12) < CDbl(C12) Then
            Fgrid.Text = DueToFromMBM((App12), (CB12))
        End If
        
        Fgrid.Row = 13
        'If C13 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C13) < CDbl(CB13) Then
            Fgrid.Text = DueToFromMBM((App13), (C13))
        ElseIf CDbl(CB13) < CDbl(C13) Then
            Fgrid.Text = DueToFromMBM((App13), (CB13))
        End If
        
        Fgrid.Row = 14
        'If C14 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C14) < CDbl(CB14) Then
            Fgrid.Text = DueToFromMBM((App14), (C14))
        ElseIf CDbl(CB14) < CDbl(C14) Then
            Fgrid.Text = DueToFromMBM((App14), (CB14))
        End If
        
        Fgrid.Row = 15
        'If C15 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C15) < CDbl(CB15) Then
            Fgrid.Text = DueToFromMBM((App15), (C15))
        ElseIf CDbl(CB15) < CDbl(C15) Then
            Fgrid.Text = DueToFromMBM((App15), (CB15))
        End If
        
        Fgrid.Row = 16
        'If C16 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C16) < CDbl(CB16) Then
            Fgrid.Text = DueToFromMBM((App16), (C16))
        ElseIf CDbl(CB16) < CDbl(C16) Then
            Fgrid.Text = DueToFromMBM((App16), (CB16))
        End If
        
        Fgrid.Row = 17
        'If C17 = 0 Then
        '    Fgrid.Text = "-"
        If CDbl(C17) < CDbl(CB17) Then
            Fgrid.Text = DueToFromMBM((App17), (C17))
        ElseIf CDbl(CB17) < CDbl(C17) Then
            Fgrid.Text = DueToFromMBM((App17), (CB17))
        End If
    
    End If
    ' &&& For Others
    Fgrid.Col = 2
    Fgrid.Row = 22

    
    'If OthersDesc1 <> "" Then
    '    Fgrid = OthersDesc1
    'End If
    
    Fgrid.Row = 19
    Fgrid.Col = 3
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    
    Fgrid = Format(PreviousLimit, "#,#.00;(#,#.00)")
    
    ' Total for Actual
    Fgrid.Col = 4
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    
    Fgrid = getTotal(4)
    
    TotalActual = Fgrid
    ' Total for Approved
    Fgrid.Col = 5
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    
    Fgrid = getTotal(5)
    totalApprove = Fgrid
    
    ' Total for Clinic
    Fgrid.Col = 6
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    
    Fgrid = getTotal(6)
    totalCLN = Fgrid
    
    'If Mid(PLNCode, 1, 1) = "2" Then
    'If PLNCode = "21" Or PLNCode = "22" Or PLNCode = "23" Or PLNCode = "24" Or PLNCode = "25" Or PLNCode = "26" Then
        'Co-Payment
        'If IsNumeric(L1) = False Then
        '    L1 = 0
        'End If
        'If CDbl(L1) < CDbl(A1) Then
            Fgrid.Col = 5
            Fgrid.Row = 20
            Fgrid.CellFontBold = True
            Fgrid.CellForeColor = vbBlue
        '        If IsNumeric(App1) = False Then
        '            App1 = 0
        '        End If
                    If IsNumeric(totalApprove) = False Then
                        totalApprove = 0
                    End If
                    'Fgrid = Format((CDbl((totalApprove) - CDbl(App1) - (CDbl(A25) * 0.8)) * 0.2), "#,#.00;(#,#.00)")
                    'CoPayApp = Fgrid
                    Fgrid = Format(CoPayApp * -1, "#,#.00;(#,#.00)")
        
            Fgrid.Col = 7
            Fgrid.Row = 20
            Fgrid.CellFontBold = True
            Fgrid.CellForeColor = vbBlue
            Fgrid = Format(CoPayApp * -1, "#,#.00;(#,#.00)")
            CoPayMBM = Fgrid
        'End If
    
    ' Total for member
    Fgrid.Row = 19
    Fgrid.Col = 7
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    'Fgrid = 1
    totalMem = getTotal(7)
    If IsNumeric(totalMem) = True And totalMem <> 0 Then
        Fgrid = totalMem
        Excess = (CDbl(CoPayMBM) + CDbl(totalMem)) * -1
        TotalExcess = CDbl(totalMem) + CDbl(CoPayMBM) + totalExceed
    Else
        Fgrid = "-"
        Excess = "0"
        TotalExcess = 0
    End If
       
    Fgrid.Col = 7
    Fgrid.Row = 22
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    
    ' Total Excess  paid by MBMer
    strsql = "select CasNumber from Claim "
    strsql = strsql & " where  CASNumber = '" & Trim(txtCASNumber) & "'"
    CLM.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    If CLM.RecordCount Then
        ClaimStatus = True
        If IsNumeric(Access) = True And Access <> "0" Then
            Fgrid = Format(Access, "#,#.00;(#,#.00)")
            Excess = Access
        Else
            Fgrid = "-"
            Excess = 0
        End If
    End If
    CLM.Close
    Set CLM = Nothing
    
    Fgrid.Col = 7
    Fgrid.Row = 21
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue

    If IsNumeric(TotalExcess) = True And TotalExcess < "0" Then
        Fgrid = Format(TotalExcess, "#,#.00;(#,#.00)")
    ElseIf IsNumeric(TotalExcess) = True And TotalExcess > "0" Then
        Fgrid = "-"
    '    Excess = Access
    Else
       Fgrid = "-"
       TotalExcess = 0
    End If
    
    'Total Excess
        Fgrid.Col = 6
        Fgrid.Row = 22
        Fgrid.CellFontBold = True
        Fgrid.CellForeColor = vbBlue
    If ClaimStatus = False Then
        If IsNumeric(totalMem) = False Then
            totalMem = 0
        End If
        Fgrid = Format(CDbl(totalMem), "#,#.00;(#,#.00)")
    Else
        Fgrid = Format(Excess, "#,#.00;(#,#.00)")
        Fgrid = Format(Fgrid * -1, "#,#.00;(#,#.00)")
    End If

     ' Paid by MEMBEr
    
    Fgrid.Col = 7
    ' Total Access paid by MBMer
    If ClaimStatus = False Then
        Fgrid.CellFontBold = True
        Fgrid.CellForeColor = vbBlue
        If IsNumeric(TotalExcess) = True And TotalExcess <> "0" And totalMem < 0 Then
            Fgrid = Format(TotalExcess * -1, "#,#.00;(#,#.00)")
        Else
            Fgrid = "-"
        End If
    Else
        Fgrid = Format(Excess, "#,#.00;(#,#.00)")
    End If
    Call getFinalTotal
    'Call ComRBLmt(A1, L1)
    ' For Not covered Items
    FGridNCover.Col = 1
    FGridNCover.Row = 1
    FGridNCover = IIf(SPFee <> 0, Format(SPFee, "#,#.00"), "-")
    FGridNCover.Row = 2
    FGridNCover = IIf(Medicine <> 0, Format(Medicine, "#,#.00"), "-")
    FGridNCover.Row = 3
    FGridNCover = IIf(Dental <> 0, Format(Dental, "#,#.00"), "-")
    FGridNCover.Row = 4
    FGridNCover = IIf(Physio <> 0, Format(Physio, "#,#.00"), "-")
    FGridNCover.Row = 5
    FGridNCover = IIf(Maternity <> 0, Format(Maternity, "#,#.00"), "-")
    FGridNCover.Row = 6
    FGridNCover = IIf(PRVCheckup <> 0, Format(PRVCheckup, "#,#.00"), "-")
    FGridNCover.Row = 7
    FGridNCover = IIf(Vision <> 0, Format(Vision, "#,#.00"), "-")
    
    'Fgrid.Col = 2
    'Fgrid.Row = 23
    'If OthersDesc1 <> "" Then
    '    Fgrid = OthersDesc1
    'End If
    'End If
End Sub

Private Function DueToFromMBM(Approved As String, PaytoHospital As String)
    If IsNumeric(Approved) = False Then
        Approved = 0
    End If
    If IsNumeric(PaytoHospital) = False Then
        PaytoHospital = 0
    End If
    If lstInvoice.SelectedItem.SubItems(7) = "C" Then
        If CDbl(PaytoHospital) > CDbl(Approved) Then
            DueToFromMBM = Format(CDbl(Approved) - CDbl(PaytoHospital), "#,#0.00;(#,#0.00)")
        End If
    Else
        DueToFromMBM = Format(CDbl(Approved) - CDbl(PaytoHospital), "#,#0.00;(#,#0.00)")
    End If

    If DueToFromMBM = 0# Then
        DueToFromMBM = "-"
    End If
    
End Function
Private Sub GetPaymentCollection()
    Dim collectByPayor, collectByHos, TotalBordxAmt, totalMemCollect As Double
    If Trim(txtCASNumber) <> "" Then
      
      
      'Excess
        'Fgrid.Row = 33
      'If optExcessPayor = True Then
      '    Fgrid.Col = 7
      '    Fgrid = ""
      '    Fgrid.Row = 32
          
      '    Fgrid.Col = 6
      '    Fgrid = chkNumber(-1 * CDbl(IIf(CDbl(TotalExcess) + CDbl(Excess) > 0, 0, CDbl(TotalExcess) + CDbl(Excess))))
      '    collectByPayor = IIf(IsNumeric(Fgrid), Fgrid, 0)
      '    collectByHos = 0
      '    Fgrid.Col = 7
      '    Fgrid.Row = 34
      '    Fgrid = chkNumber(CStr(totalFCLN))

      'ElseIf optExcessCLN = True Then
      '    Fgrid.Col = 6
      '    Fgrid = ""
      '    collectByPayor = 0
      '    Fgrid.Col = 6
      '    Fgrid = chkNumber(IIf(CDbl(TotalExcess) + CDbl(Excess) > 0, 0, CDbl(TotalExcess) + CDbl(Excess)))
      '    collectByHos = IIf(IsNumeric(Fgrid), Fgrid, 0)
      '    'Fgrid.Row = 30
      '    Fgrid.Row = 34
      '    'Fgrid = chkNumber(CStr(Fgrid + CDbl(collectByHos)))
      '    Fgrid = chkNumber(CStr(totalFCLN + Excess + CoPayApp))
      
      'Else
      '    Fgrid.Col = 6
      '    Fgrid = ""
      '    collectByHos = 0
      '    Fgrid.Col = 7
      '    Fgrid = ""
      '    collectByPayor = 0
      '    'Fgrid.Row = 30
      '    Fgrid.Row = 34
      '    Fgrid = chkNumber(CStr(totalFCLN))
      'End If
      
    
      'If totalExceed < 0 Then
      '    'Limit Exceeded
      '    Fgrid.Row = 23
      '    If optLimitCLN = True Then
      '        Fgrid.Col = 5
      '        Fgrid = chkNumber(CStr(collectByPayor))
      '        Fgrid.Col = 6
      '        If totalExceed < 0 Then
      '            Fgrid = chkNumber(IIf(IsNumeric(Fgrid), Fgrid, 0) + (totalExceed))
      '            collectByHos = CDbl(Fgrid)
      '        End If
             
      '        'Total payable to Hospital
      '        Fgrid.Row = 23
      '        Fgrid = chkNumber(CStr(totalFCLN + CDbl(collectByHos)))
              
      '        'Hospital/Payor to collect for hospital Column
      '        Fgrid.Col = 7
      '        Fgrid = chkNumber(IIf(IsNumeric(collectByPayor), collectByPayor, 0) - CDbl(IIf(IsNumeric(collectByHos), collectByHos, 0)))
      '        totalMemCollect = IIf(IsNumeric(Fgrid), Fgrid, 0)
              
      '    ElseIf optLimitPayor = True Then
      '        Fgrid.Col = 6
      '        Fgrid = chkNumber(CStr(collectByHos))
      '        Fgrid.Col = 7
      '        If totalExceed < 0 Then
      '            Fgrid = chkNumber(IIf(IsNumeric(Fgrid), Fgrid, 0) - CDbl(totalExceed))
      '            collectByPayor = CDbl(Fgrid)
      '        End If
      '    Else
      '        Fgrid.Col = 5
      '        Fgrid = chkNumber(CStr(collectByPayor))
      '        Fgrid.Col = 6
      '        Fgrid = chkNumber(CStr(collectByHos))
      '    End If
       'End If
      
      'Total Bordx Amount
      Fgrid.Col = 5
      Fgrid.Row = 26
      Fgrid.CellFontBold = True
      Fgrid.CellForeColor = vbBlue
      
      Fgrid = chkNumber(collectByPayor + CDbl(totalFApprove))
      TotalBordxAmt = IIf(IsNumeric(Fgrid), Fgrid, 0)
      BordxAmount = IIf(IsNumeric(Fgrid), Fgrid, 0)
      
      'Hospital/Payor to collect for member  Column
      'Fgrid.Row = 33
      'Fgrid.Col = 8
      'Fgrid = chkNumber(IIf(IsNumeric(collectByPayor), collectByPayor, 0) - CDbl(IIf(IsNumeric(collectByHos), collectByHos, 0)))
      'totalMemCollect = IIf(IsNumeric(Fgrid), Fgrid, 0)
      
      'Final Total for Member
      'Fgrid.Col = 8
      'Fgrid.Row = 34
      '  If ClickTExcess = True Then
      '      If Excess = 0 Then
      '         TotalExcess = TotalExcess * -1
      '     End If
      '      Fgrid = IIf(((TotalExcess * -1) = Excess), "0", Format(TotalExcess, "#,#.00;(#,#.00)"))
      '  Else
      '      Fgrid = chkNumber(totalMemCollect + CoPayApp * -1 + totalExceed + totalMem + Excess)
      '  End If

    End If
End Sub

Private Function chkNumber(str As String) As String
    chkNumber = "-"
    If IsNumeric(str) = True Then
        If str <> "0" Then
            chkNumber = Format(str, "#,#.00;(#,#.00)")
        End If
    ElseIf ChkStr(str) = "FULL" Then
        chkNumber = "FULL"
    End If
End Function

Private Sub SaveClaim(Optional Approve As String)
    'On Error Resume Next
    Dim Claim As New ADODB.Recordset
    Dim strsqlCAS As String
    Dim CAS As New ADODB.Recordset
    
    strsqlCAS = "select CasNumber , * from Claim "
    strsqlCAS = strsqlCAS & " where  CASNumber = '" & Trim(txtCASNumber) & "'"
    Claim.Open strsqlCAS, MediConnCISWS, adOpenKeyset, adLockOptimistic

    'Claim.Open "select CAsnumber , * from Claim", MediConnCISWS, adOpenKeyset, adLockOptimistic
    'Claim.Close
    If Claim.RecordCount = 0 Then
        Claim.AddNew
    End If
    With Claim
         !CASNumber = ChkStr(txtCASNumber)
         '!clmversion = cboVersion
         !ClaimOpendate = Date
         !ClaimOpenBy = ChkStr(Logon)
         !ClaimLastUpdateUser = ChkStr(Logon)
         !ClaimLastUpdateDate = Date
         !ClaimBordxStatus = 0
'         If ChkStr(Approve) = "A" Then
            !ClaimStatus = "C"
            !ClaimClosedBy = ChkStr(Logon)
            !ClaimCloseDate = Date
 '        Else
  '          !ClaimStatus = "O"
   '      End If
        ' Total for Actual
        'Fgrid.Row = 25
        'Fgrid.Col = 5
        !CLMTotalActual = IIf(IsNumeric(TotalFActual) = True, TotalFActual, 0)
        ' Total for Approved
        'Fgrid.Col = 6
        !CLMTotalApproved = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
         
        ' Total for Hospital
        'Fgrid.Col = 7
        !CLMPayableByMedix2C = IIf(IsNumeric(totalFCLN) = True, totalFCLN, 0)
        
        ' Total for member
        'Fgrid.Col = 8
        ' "+" ==> Credit note   "-" == > Debit note
        'If Len(Claim!CLMPayablebyMedix2m) And Claim!CLMPayablebyMedix2m <> totalFMem Then
        '    !DebitCreditPreviousAmt = !CLMPayablebyMedix2m
        '    !DbCrListPrintStatus = False
        'End If

        !CLMCompExcess = IIf(IsNumeric(RBCOMExcess) = True, RBCOMExcess, 0)
        
        !CLMTotalCompApp = IIf(IsNumeric(RBComAmt) = True, RBComAmt, 0)
        !CLMBordxWSStatus = "N"
        !CLMBordxWSASCII = "N"
        
        ' Excess and Reimbursement
        !CLMExcess = CDbl(totalExcessMemNeed2Pay) + (-44)
        !CLMReimburse = TotalReimburse
        '!CLMBalanceLimit = txtBalLimit - totalFApprove
        !CLMBalanceLimit = PreviousLimit - totalFApprove
        '!CLMPreviousBalance = txtBalLimit
         
        !CLMPaidBYMedix2M = 0
        !CLMPaidByMedix2C = 0
        !CLMLimitExceeded = IIf(IsNumeric(totalExceed) = True, totalExceed, 0)
        !CLMDiscount = 0
        !CLMPayableByMedix2M = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
        If optExcessPayor = True Then
           !ExcessCollectedBy = "P"
        ElseIf optExcessCLN = True Then
           !ExcessCollectedBy = "H"
        Else
           !ExcessCollectedBy = "M"
        End If
         
        If optLimitPayor = True Then
           !LimitCollectedBy = "P"
        ElseIf optLimitCLN = True Then
           !LimitCollectedBy = "H"
        Else
           !LimitCollectedBy = "M"
        End If
        
        
        'If Len(txtWSRemarks) Then
            !CLMWSRemarks = txtWSRemarks
        'End If
        'If Len(txtCDRemarks) Then
            !CLMDCNoteRemarks = txtCDRemarks
        'End If
            '!CLMCoPayRB = CoPayApp
        
        
    End With
    Claim.update
    
    strsqlCAS = " select CASnumber,CASPending from CAS "
    strsqlCAS = strsqlCAS & " where  CASNumber = '" & Trim(txtCASNumber) & "'"
    CAS.Open strsqlCAS, medixmasterconn, adOpenKeyset, adLockOptimistic
                
    If CAS.RecordCount Then
        CAS!CASPending = "N"
        CAS.update
    End If
    CAS.Close
    Set CAS = Nothing

    
    Claim.Close
    Set Claim = Nothing
End Sub

Private Sub saveLimit()
    Dim ClaimLimit  As New ADODB.Recordset
    Dim strsqlLMT As String
    
    strsqlLMT = "select CasNumber , * from CLMLimit "
    strsqlLMT = strsqlLMT & " where  CASNumber = '" & Trim(txtCASNumber) & "'"
    ClaimLimit.Open strsqlLMT, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    If ClaimLimit.RecordCount = 0 Then
        ClaimLimit.AddNew
    End If

    With ClaimLimit
         Fgrid.Col = 3
         !CASNumber = ChkStr(txtCASNumber)
         '!clmversion = cboVersion
         Fgrid.Row = 1
            !CLMGPConsAmt = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 2
            !CLMSPConsAmt = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 3
            !CLMMedicine = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 4
            !CLMProcedures = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
        Fgrid.Row = 5
            !CLMLab = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
        Fgrid.Row = 6
            !CLMXray = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         'Fgrid.Row = 7
         '   !CLMMRI = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 7
            !CLMDental = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 8
            !CLMChineseMed = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 9
            !CLMTraditionalMEd = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 10
            !CLMphysioTheraphy = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
        
         Fgrid.Row = 11
            !CLMAccident = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 12
            !CLMPRVCheckup = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 13
            !CLMVision = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 14
            !CLMID = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 15
            !CLMTSSTO = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
        
         Fgrid.Row = 16
            !CLMInjection = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 17
            !CLMOTHERS1 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
    End With
    ClaimLimit.update
    ClaimLimit.Close
    Set ClaimLimit = Nothing
End Sub
Private Sub saveActual()
    Dim ClaimActual  As New ADODB.Recordset
    Dim strsqlACT As String
    
    strsqlACT = "select CasNumber , * from CLMActual "
    strsqlACT = strsqlACT & " where  CASNumber = '" & Trim(txtCASNumber) & "'"
    ClaimActual.Open strsqlACT, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    If ClaimActual.RecordCount = 0 Then
        ClaimActual.AddNew
    End If

    With ClaimActual
         Fgrid.Col = 4
         !CASNumber = ChkStr(txtCASNumber)
         '!clmversion = cboVersion
         Fgrid.Row = 1
            !CLMGPConsAmt = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            
         Fgrid.Row = 2
            !CLMSPConsAmt = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 3
            !CLMMedicine = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 4
            !CLMProcedures = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
        Fgrid.Row = 5
            !CLMLab = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
        Fgrid.Row = 6
            !CLMXray = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         'Fgrid.Row = 7
         '   !CLMMRI = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 7
            !CLMDental = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 8
            !CLMChineseMed = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 9
            !CLMTraditionalMEd = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 10
            !CLMphysioTheraphy = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 11
            !CLMAccident = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 12
            !CLMPRVCheckup = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 13
            !CLMVision = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 14
            !CLMID = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 15
            !CLMTSSTO = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
        
         Fgrid.Row = 16
            !CLMInjection = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 17
            !CLMOTHERS1 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
    End With
    ClaimActual.update
    ClaimActual.Close
    Set ClaimActual = Nothing
End Sub

Private Sub saveApprove()
    Dim ClaimAPP  As New ADODB.Recordset
    Dim strsqlAPP As String
    
    strsqlAPP = "select CasNumber , * from CLMApp "
    strsqlAPP = strsqlAPP & " where  CASNumber = '" & Trim(txtCASNumber) & "'"
    ClaimAPP.Open strsqlAPP, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    If ClaimAPP.RecordCount = 0 Then
        ClaimAPP.AddNew
    End If

    With ClaimAPP
         Fgrid.Col = 5
         !CASNumber = ChkStr(txtCASNumber)
         '!clmversion = cboVersion
         Fgrid.Row = 1
            !CLMGPConsAmt = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 2
            !CLMSPConsAmt = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 3
            !CLMMedicine = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 4
            !CLMProcedures = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
        Fgrid.Row = 5
            !CLMLab = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
        Fgrid.Row = 6
            !CLMXray = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         'Fgrid.Row = 7
         '   !CLMMRI = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 7
            !CLMDental = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 8
            !CLMChineseMed = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 9
            !CLMTraditionalMEd = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 10
            !CLMphysioTheraphy = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 11
            !CLMAccident = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 12
            !CLMPRVCheckup = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 13
            !CLMVision = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 14
            !CLMID = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 15
            !CLMTSSTO = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
        
         Fgrid.Row = 16
            !CLMInjection = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 17
            !CLMOTHERS1 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
    End With
    ClaimAPP.update
    ClaimAPP.Close
    Set ClaimAPP = Nothing
End Sub

Private Sub saveCLN()
    Dim ClaimCLN  As New ADODB.Recordset
    Dim strsqlCLN As String
    
    strsqlCLN = "select CasNumber , * from CLMClinic "
    strsqlCLN = strsqlCLN & " where  CASNumber = '" & Trim(txtCASNumber) & "'"
    ClaimCLN.Open strsqlCLN, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    If ClaimCLN.RecordCount = 0 Then
        ClaimCLN.AddNew
    End If

    With ClaimCLN
         Fgrid.Col = 6
         !CASNumber = ChkStr(txtCASNumber)
         '!clmversion = cboVersion
         Fgrid.Row = 1
            !CLMGPConsAmt = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 2
            !CLMSPConsAmt = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 3
            !CLMMedicine = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 4
            !CLMProcedures = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
        Fgrid.Row = 5
            !CLMLab = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
        Fgrid.Row = 6
            !CLMXray = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         'Fgrid.Row = 7
         '   !CLMMRI = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 7
            !CLMDental = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 8
            !CLMChineseMed = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 9
            !CLMTraditionalMEd = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 10
            !CLMphysioTheraphy = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 11
            !CLMAccident = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 12
            !CLMPRVCheckup = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 13
            !CLMVision = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 14
            !CLMID = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 15
            !CLMTSSTO = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
        
         Fgrid.Row = 16
            !CLMInjection = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 17
            !CLMOTHERS1 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
    End With
    ClaimCLN.update
    ClaimCLN.Close
    Set ClaimCLN = Nothing
End Sub

Private Sub saveMBM()
    Dim ClaimMBM  As New ADODB.Recordset
    Dim strsqlMBM As String
    
    strsqlMBM = "select CasNumber , * from CLMMember "
    strsqlMBM = strsqlMBM & " where  CASNumber = '" & Trim(txtCASNumber) & "'"
    ClaimMBM.Open strsqlMBM, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    If ClaimMBM.RecordCount = 0 Then
        ClaimMBM.AddNew
    End If

    With ClaimMBM
         Fgrid.Col = 7
         !CASNumber = ChkStr(txtCASNumber)
         '!clmversion = cboVersion
         Fgrid.Row = 1
            !CLMGPConsAmt = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            
         Fgrid.Row = 2
            !CLMSPConsAmt = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 3
            !CLMMedicine = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 4
            !CLMProcedures = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
        Fgrid.Row = 5
            !CLMLab = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
        Fgrid.Row = 6
            !CLMXray = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         'Fgrid.Row = 7
         '   !CLMMRI = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 7
            !CLMDental = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 8
            !CLMChineseMed = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 9
            !CLMTraditionalMEd = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 10
            !CLMphysioTheraphy = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 11
            !CLMAccident = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 12
            !CLMPRVCheckup = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 13
            !CLMVision = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 14
            !CLMID = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 15
            !CLMTSSTO = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            
         Fgrid.Row = 16
            !CLMInjection = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 17
            !CLMOTHERS1 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
    End With
    ClaimMBM.update
    ClaimMBM.Close
    Set ClaimMBM = Nothing
End Sub

Private Sub UpdateInvoice()
    Dim strsql As String
    If totalFMem <= 0 Then
        strsql = "update ClaimInvoice set INVWSStatus = '1', INVPaidByMEM =  Convert(Money, '" & Excess & "') " & _
            " where CASNumber ='" & ChkStr(txtCASNumber) & "'"
    Else
        strsql = "update ClaimInvoice set INVWSStatus = '1', INVPaidByMEM =  Convert(Money, '" & 0 & "') " & _
            " where CASNumber ='" & ChkStr(txtCASNumber) & "'"
    End If

    MediConnCISWS.Execute strsql

End Sub

Private Sub UpdateCreditDebit(str As String)
    Dim Claim As New ADODB.Recordset
    Dim strsql As String
    strsql = "Select CasNUmber , DebitCreditRefNo , debitCreditDate, CLMPayablebyMedix2m  from Claim " & _
        " where CasNUmber ='" & txtCASNumber & "'"
    Claim.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    If Claim.RecordCount > 0 Then
        Claim!DebitCreditRefNo = Trim(ChkStr(str))
        Claim!DebitCreditDate = Date
    End If
    Claim.update
    Claim.Close
    Set Claim = Nothing

End Sub

Private Sub getVersion()
    'Dim Version As New ADODB.Recordset
    'Dim strsql As String
    'Dim currIndex As Integer
    'currIndex = cboVersion.ListIndex
    'cboVersion.Clear
    'strsql = "select casNumber, clmversion , CLMLimitExceeded from claim where casNUmber ='" & Trim(txtCASNumber) & "'"
    'Version.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    ''currIndex = cboVersion.ListIndex
    ''cboVersion.Clear

    ''If cboVersion = "" Then
    ''    currIndex = "0"
    ''Else
    ''    currIndex = cboVersion.ListIndex
    ''End If

    'If Not Version.EOF Then
    '    Do While Not Version.EOF
    '        totalExceed = 0
    '        cboVersion.AddItem Version!clmversion
    '        cboVersion.AddItem "M" & Version!clmversion
    '        'For MAster only
    '        totalExceed = CDbl(totalExceed) + Version!CLMLimitExceeded
    '        Version.MoveNext
    '    Loop
    'Else
     '   cboVersion.AddItem "1"
    'End If

    'Version.Close
    'Set Version = Nothing
    'cboVersion.AddItem "Master"
    ''If currIndex  = "" Then
    'cboVersion.ListIndex = currIndex
    ''End If
End Sub

Private Sub UpdateMBMBalance()
    Dim MBM As New ADODB.Recordset
    Dim strsql As String
    Dim FinalBalance  As Double
    Dim strsqlannlmt As String
    Dim AnnLmt As New ADODB.Recordset
    
    If lblClaimability <> "Repudiated" Then
        strsqlannlmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(txtPLNCode) & "'"
        AnnLmt.Open strsqlannlmt, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        If ChkStr(txtMemType) = "P" Or Trim(txtPatientID) = "00" Then
            strsql = "Select MBMNUmber , MBMCLNBalLmt from MBM " & _
                    " where MBMNUmber   ='" & Trim(txtMBMNumber) & "'"
        Else
            If AnnLmt!SuppLimitStatus = "B" Then
                strsql = "Select MBMCNUmber , MBMCAvailableLimit as MBMAvailableLimit from MBMCoveredPersons " & _
                " where MBMCNUmber   ='" & Trim(txtMBMNumber) & "'"
            Else
                strsql = "Select MBMCNUmber , MBMCAvailableLimit as MBMAvailableLimit from MBMCoveredPersons " & _
                        " where MBMCNUmber   ='" & Trim(txtMBMNumber) & "' And MBMCCoverID   = '" & Trim(txtPatientID) & "'"
            End If
        End If
        
        MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If IsNumeric(totalApprove) = False Then
            totalApprove = 0
        End If
        With MBM
            FinalBalance = PreviousLimit - totalFApprove
            If FinalBalance < 0 Then
                !MBMCLNBalLmt = 0
            Else
                !MBMCLNBalLmt = PreviousLimit - totalFApprove
            End If
        End With
        MBM.update
        MBM.Close
        Set MBM = Nothing
        AnnLmt.Close
        Set AnnLmt = Nothing
    End If
End Sub

Private Sub InputLimitToExcel(objWorksheet As Excel.Worksheet)
    Dim tempOthers  As Double
    
    Fgrid.Col = 3
    Fgrid.Row = 1
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(14, "F") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 2
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(15, "F") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 3
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(16, "F") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(17, "F") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(18, "F") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 6
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(19, "F") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(20, "F") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(21, "F") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 9
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(22, "F") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 10
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(23, "F") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 11
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(24, "F") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 12
    If Trim(Fgrid) <> "-" Then
         objWorksheet.Cells(25, "F") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 13
    If Trim(Fgrid) <> "-" Then
         objWorksheet.Cells(26, "F") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 14
    If Trim(Fgrid) <> "-" Then
     objWorksheet.Cells(27, "F") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 15
    If Trim(Fgrid) <> "-" Then
      objWorksheet.Cells(28, "F") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 16
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(29, "F") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 17
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(30, "F") = chkNumber(Fgrid)
    End If
    'Fgrid.Row = 18
    'If Trim(Fgrid) <> "-" Then
    '   objWorksheet.Cells(31, "F") = chkNumber(Fgrid)
    'End If

    'Balance Limit
    Fgrid.Row = 19
    If ANNInd = True Then
        If Trim(Fgrid) <> "-" And Trim(Fgrid) <> "" Then
            objWorksheet.Cells(31, "F") = CDbl(chkNumber(Fgrid))
        End If
    Else
        objWorksheet.Cells(31, "F") = "FULL"
    End If
End Sub
Private Sub InputActualToExcel(objWorksheet As Excel.Worksheet)
    Dim tempOthers  As Double
    
    Fgrid.Col = 4
    Fgrid.Row = 1
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(14, "G") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 2
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(15, "G") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 3
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(16, "G") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(17, "G") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(18, "G") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 6
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(19, "G") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(20, "G") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(21, "G") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 9
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(22, "G") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 10
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(23, "G") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 11
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(24, "G") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 12
    If Trim(Fgrid) <> "-" Then
         objWorksheet.Cells(25, "G") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 13
    If Trim(Fgrid) <> "-" Then
         objWorksheet.Cells(26, "G") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 14
    If Trim(Fgrid) <> "-" Then
     objWorksheet.Cells(27, "G") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 15
    If Trim(Fgrid) <> "-" Then
      objWorksheet.Cells(28, "G") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 16
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(29, "G") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 17
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(30, "G") = chkNumber(Fgrid)
    End If
    'Fgrid.Row = 18
    'If Trim(Fgrid) <> "-" Then
    '   objWorksheet.Cells(31, "G") = chkNumber(Fgrid)
    'End If
    
End Sub
Private Sub InputAppToExcel(objWorksheet As Excel.Worksheet)
    Dim tempOthers  As Double
    
    Fgrid.Col = 5
    Fgrid.Row = 1
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(14, "H") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 2
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(15, "H") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 3
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(16, "H") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(17, "H") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(18, "H") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 6
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(19, "H") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(20, "H") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(21, "H") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 9
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(22, "H") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 10
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(23, "H") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 11
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(24, "H") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 12
    If Trim(Fgrid) <> "-" Then
         objWorksheet.Cells(25, "H") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 13
    If Trim(Fgrid) <> "-" Then
         objWorksheet.Cells(26, "H") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 14
    If Trim(Fgrid) <> "-" Then
     objWorksheet.Cells(27, "H") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 15
    If Trim(Fgrid) <> "-" Then
      objWorksheet.Cells(28, "H") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 16
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(29, "H") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 17
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(30, "H") = chkNumber(Fgrid)
    End If
    'Fgrid.Row = 18
    'If Trim(Fgrid) <> "-" Then
    '   objWorksheet.Cells(31, "H") = chkNumber(Fgrid)
    'End If
    
End Sub

Private Sub InputCLNToExcel(objWorksheet As Excel.Worksheet)
    Dim tempOthers  As Double
    
    Fgrid.Col = 6
    Fgrid.Row = 1
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(14, "I") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 2
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(15, "I") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 3
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(16, "I") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(17, "I") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(18, "I") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 6
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(19, "I") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(20, "I") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(21, "I") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 9
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(22, "I") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 10
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(23, "I") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 11
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(24, "I") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 12
    If Trim(Fgrid) <> "-" Then
         objWorksheet.Cells(25, "I") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 13
    If Trim(Fgrid) <> "-" Then
         objWorksheet.Cells(26, "I") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 14
    If Trim(Fgrid) <> "-" Then
     objWorksheet.Cells(27, "I") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 15
    If Trim(Fgrid) <> "-" Then
      objWorksheet.Cells(28, "I") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 16
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(29, "I") = chkNumber(Fgrid)
    End If
    Fgrid.Row = 17
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(30, "I") = chkNumber(Fgrid)
    End If
    'Fgrid.Row = 18
    'If Trim(Fgrid) <> "-" Then
    '   objWorksheet.Cells(31, "I") = chkNumber(Fgrid)
    'End If
    
    Fgrid.Row = 22
    If Trim(Fgrid) <> "-" Then
       objWorksheet.Cells(34, "I") = chkNumber(Fgrid)
    End If

End Sub

Private Sub InputCActualToExcel(objWorksheet As Excel.Worksheet)
    Dim tempOthers  As Double
    objWorksheet.Cells(14, "G") = CGPConsAmt
    objWorksheet.Cells(15, "G") = CSPConsAmt
    objWorksheet.Cells(16, "G") = CMedicine
    objWorksheet.Cells(17, "G") = CProcedures
    objWorksheet.Cells(18, "G") = CLab
    objWorksheet.Cells(19, "G") = CXRay
    'objWorksheet.Cells(20, "G") = CMRI
    objWorksheet.Cells(20, "G") = CDental
    objWorksheet.Cells(21, "G") = CChineseMed
    objWorksheet.Cells(22, "G") = CTraditionalMed
    objWorksheet.Cells(23, "G") = CPhysioTheraphy
    objWorksheet.Cells(24, "G") = CAccident
    objWorksheet.Cells(25, "G") = CPRVCheckup
    objWorksheet.Cells(26, "G") = CVision
    objWorksheet.Cells(27, "G") = CID
    objWorksheet.Cells(28, "G") = CTSSTO
    objWorksheet.Cells(29, "G") = COthers1
    objWorksheet.Cells(30, "G") = COthers2
End Sub

Private Sub InputReimToExcel(objWorksheet As Excel.Worksheet)
    Dim tempOthers  As Double
     FGridNCover.Col = 1
    
    FGridNCover.Row = 1
    If Trim(FGridNCover) <> "-" Then
        objWorksheet.Cells(32, "G") = chkNumber(FGridNCover)
    End If
    
    FGridNCover.Row = 2
    If Trim(FGridNCover) <> "-" Then
        objWorksheet.Cells(33, "G") = chkNumber(FGridNCover)
    End If
    
    FGridNCover.Row = 3
    If Trim(FGridNCover) <> "-" Then
        objWorksheet.Cells(34, "G") = chkNumber(FGridNCover)
    End If
    
    FGridNCover.Row = 4
    If Trim(FGridNCover) <> "-" Then
        objWorksheet.Cells(35, "G") = chkNumber(FGridNCover)
    End If
    
    FGridNCover.Row = 5
    If Trim(FGridNCover) <> "-" Then
        objWorksheet.Cells(36, "G") = chkNumber(FGridNCover)
    End If
    
    FGridNCover.Row = 6
    If Trim(FGridNCover) <> "-" Then
        objWorksheet.Cells(37, "G") = chkNumber(FGridNCover)
    End If
    
    FGridNCover.Row = 7
    If Trim(FGridNCover) <> "-" Then
        objWorksheet.Cells(38, "G") = chkNumber(FGridNCover)
    End If
End Sub
Private Sub InputReimActualToExcel(objWorksheet As Excel.Worksheet)
    Dim tempOthers  As Double
    
    objWorksheet.Cells(13, "G") = RGPConsAmt
    objWorksheet.Cells(14, "G") = RSPConsAmt
    objWorksheet.Cells(15, "G") = RMedicine
    objWorksheet.Cells(16, "G") = RProcedures
    objWorksheet.Cells(17, "G") = RLab
    objWorksheet.Cells(18, "G") = RXRay
    'objWorksheet.Cells(19, "G") = RMRI
    objWorksheet.Cells(19, "G") = RDental
    objWorksheet.Cells(20, "G") = RChineseMed
    objWorksheet.Cells(21, "G") = RTraditionalMed
    objWorksheet.Cells(22, "G") = RPhysioTheraphy
    objWorksheet.Cells(23, "G") = RAccident
    objWorksheet.Cells(24, "G") = RPRVCheckup
    objWorksheet.Cells(25, "G") = RVision
    objWorksheet.Cells(26, "G") = RID
    objWorksheet.Cells(27, "G") = RTSSTO
    objWorksheet.Cells(28, "G") = ROthers1
    objWorksheet.Cells(29, "G") = ROthers2
    objWorksheet.Cells(42, "H") = totalExceed

End Sub

Private Sub GetInvoiceToExcel(obj As Excel.Worksheet)
    Dim ListCount As Integer
    Dim ForLoop As Integer
    Dim curColReim As String
    Dim curColCashless As String
        Call InputLimitToExcel(obj)
        Call InputActualToExcel(obj)
        Call InputAppToExcel(obj)
        Call InputCLNToExcel(obj)
End Sub

Public Sub display_CLNbenefit()
    Dim strsqlCOPay As String
    Dim COPAY As New ADODB.Recordset
    Dim CLNBNF As New ADODB.Recordset
    Dim strsql As String
    Dim counter As Integer
    Dim LMT As Integer
    Dim strsqlCLN As String
    Dim CLN As New ADODB.Recordset
    
    CB1 = 0
    CB2 = 0
    CB3 = 0
    CB4 = 0
    CB5 = 0
    CB6 = 0
    CB7 = 0
    CB8 = 0
    CB9 = 0
    CB10 = 0
    CB11 = 0
    CB12 = 0
    CB13 = 0
    CB14 = 0
    CB15 = 0
    CB16 = 0
    CB17 = 0
    CB18 = 0

    
        strsqlCLN = "Select CLNCode ,CLNName, CLNBNFType from CLN  " & _
        " where CLNCode = '" & Trim(CLNCode) & "'"
        CLN.Open strsqlCLN, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If CLN.RecordCount Then
            'If CLN!CLNBNFType = "1" Then
            '    strsql = "Select BNFCode ,CLNBNFRMPerVisit, CLNBNFEffDate,CLNBNFCalculation from STDClinicBNF  " & _
            '    " where CLNCode = '" & Trim(CLNCode) & "'"
            '    strsql = strsql & " order by BNFCOde "
            '    CLNBNF.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            If CLN!CLNBNFType = "2" Then
        
                strsql = "Select BNFCode ,BNFRMPerVisit, BNFEffDate,BNFCalculation from ClinicBNF  " & _
                " where PLnCode = '" & txtPLNCode & "' and INSCode = '" & Trim(txtINSType) & "' and HLTCode = '" & Trim(lblHLTcode) & "'"
                strsql = strsql & " order by BNFCOde "
                CLNBNF.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic

        
            Do Until CLNBNF.EOF
                 With CLNBNF
                    For counter = 1 To 17 Step 1
                     Fgrid.Row = counter
                     Fgrid.Col = 0
                        If Fgrid.Text = !BNFCode Then
                            ' LImit col
                            Fgrid.Col = 3
                                If !BNFCode = "001" Then
                                    If I1 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB1 = Format((I1), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB1 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB1 = 0
                                    End If
                                ElseIf !BNFCode = "002" Then
                                    If I2 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB2 = Format((I2), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB2 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB2 = 0
                                    End If
                                ElseIf !BNFCode = "003" Then
                                    If I3 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB3 = Format((I3), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB3 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB3 = 0
                                    End If
    
                                ElseIf !BNFCode = "004" Then
                                    If I4 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB4 = Format((I4), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB4 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB4 = 0
                                    End If
                                
                                ElseIf !BNFCode = "005" Then
                                    If I5 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB5 = Format((I5), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB5 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB5 = 0
                                    End If
                                        
                                ElseIf !BNFCode = "006" Then
                                    If I6 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB6 = Format((I6), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB6 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB6 = 0
                                    End If
                                        
                                ElseIf !BNFCode = "007" Then
                                    If I7 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB7 = Format((I7), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB7 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB7 = 0
                                    End If
                                    
                                ElseIf !BNFCode = "008" Then
                                    If I8 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB8 = Format((I8), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB8 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB8 = 0
                                    End If
    
                                ElseIf !BNFCode = "009" Then
                                    If I9 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB9 = Format((I9), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB9 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB9 = 0
                                    End If
                                    
                                ElseIf !BNFCode = "010" Then
                                    If I10 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB10 = Format((I10), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB10 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB10 = 0
                                    End If
                                        
                                ElseIf !BNFCode = "011" Then
                                    If I11 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB11 = Format((I11), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB11 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB11 = 0
                                    End If
                                        
                                ElseIf !BNFCode = "012" Then
                                    If I12 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB12 = Format((I12), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB12 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB12 = 0
                                    End If
                                    
                                ElseIf !BNFCode = "013" Then
                                    If I13 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
    
                                        If !BNFCalculation = "1" Then
                                            CB13 = Format((I13), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB13 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB13 = 0
                                    End If
                                        
                                ElseIf !BNFCode = "014" Then
                                    If I14 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB14 = Format((I14), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB14 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB14 = 0
                                    End If
                                        
                                ElseIf !BNFCode = "015" Then
                                    If I15 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB15 = Format((I15), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB15 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB15 = 0
                                    End If
                                    
                                ElseIf !BNFCode = "016" Then
                                    If I16 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB16 = Format((I16), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB16 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB16 = 0
                                    End If
                                        
                                ElseIf !BNFCode = "017" Then
                                    If I17 > 0 Then
                                        If IsNumeric(!BNFRMPerVisit) = False Then
                                            LMT = 0
                                        Else
                                            LMT = !BNFRMPerVisit
                                        End If
                                        If !BNFCalculation = "1" Then
                                            CB17 = Format((I17), "#,##0.00")
                                        ElseIf !BNFCalculation = "2" Then
                                            CB17 = Format(LMT, "#,##0.00")
                                        End If
                                    Else
                                        CB17 = 0
                                    End If
                        'ElseIf !BNFRMPerVisit = 0 Then
                        '    Fgrid.Text = 0
                        End If
                        
                        'to Display Calculation Type
                        'If Len(!BNFCalculation) Then
                        '    strsqlCOPay = "Select COPCode , COPDesc from COPayment  " & _
                        '    " where COPCode = '" & !BNFCalculation & "'"
                        '    COPAY.Open strsqlCOPay, MediConnCISDB, adOpenKeyset, adLockOptimistic
                            
                        '    If COPAY.RecordCount Then
                        '        Fgrid.Col = 2
                        '        Fgrid.Text = !BNFCalculation & "-" & COPAY!COPDesc
                        '    End If
                         '   COPAY.Close
                         '   Set COPAY = Nothing
                        'End If
    
                        Exit For
                    'End If
                    'End If
                    End If
    
                Next
                .MoveNext
             End With
        Loop
        
        'If CLNBNF.RecordCount = 0 Then
        '    BNFStatus = False
        'Else
        '    BNFStatus = True
        'End If
        CLNBNF.Close
        Set CLNBNF = Nothing
    Else
        GenCB = True
    End If
    End If
    CLN.Close
    Set CLN = Nothing
        
End Sub

Private Sub saveCLNLimit()
    Dim ClaimCLNLimit  As New ADODB.Recordset
    Dim strsqlCLNLMT As String
    
    strsqlCLNLMT = "select CasNumber , * from CLMCLNLimit "
    strsqlCLNLMT = strsqlCLNLMT & " where  CASNumber = '" & Trim(txtCASNumber) & "'"
    ClaimCLNLimit.Open strsqlCLNLMT, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    If ClaimCLNLimit.RecordCount = 0 Then
        ClaimCLNLimit.AddNew
    End If

    With ClaimCLNLimit
         Fgrid.Col = 6
         !CASNumber = ChkStr(txtCASNumber)
         '!clmversion = CLMVersionNo
         Fgrid.Row = 1
            !CLMGPConsAmt = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 2
            !CLMSPConsAmt = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 3
            !CLMMedicine = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 4
            !CLMProcedures = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
        Fgrid.Row = 5
            !CLMLab = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
        Fgrid.Row = 6
            !CLMXray = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         'Fgrid.Row = 7
         '   !CLMMRI = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 7
            !CLMDental = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 8
            !CLMChineseMed = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 9
            !CLMTraditionalMEd = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 10
            !CLMphysioTheraphy = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 11
            !CLMAccident = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 12
            !CLMPRVCheckup = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 13
            !CLMVision = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 14
            !CLMID = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 15
            !CLMTSSTO = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
        
         Fgrid.Row = 16
            !CLMInjection = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
         Fgrid.Row = 17
            !CLMOTHERS1 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
         
    End With
    ClaimCLNLimit.update
    ClaimCLNLimit.Close
    Set ClaimCLNLimit = Nothing
End Sub

Private Sub InputMBMToExcel(objWorksheet As Excel.Worksheet)
    Dim tempOthers  As Double
    Fgrid.Col = 7
    
    Fgrid.Row = 1
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(14, "J") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 2
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(15, "J") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 3
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(16, "J") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 4
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(17, "J") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 5
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(18, "J") = chkNumber(Fgrid)
    End If

    Fgrid.Row = 6
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(19, "J") = chkNumber(Fgrid)
    End If

    Fgrid.Row = 7
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(20, "J") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 8
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(21, "J") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 9
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(22, "J") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 10
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(23, "J") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 11
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(24, "J") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 12
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(25, "J") = chkNumber(Fgrid)
    End If

    Fgrid.Row = 13
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(26, "J") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 14
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(27, "J") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 15
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(28, "J") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 16
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(29, "J") = chkNumber(Fgrid)
    End If
    
    Fgrid.Row = 17
    If Trim(Fgrid) <> "-" Then
        objWorksheet.Cells(30, "J") = chkNumber(Fgrid)
    End If

    If CDbl(totalExceed) <> 0 Then
        objWorksheet.Cells(35, "H") = totalExceed
    End If

    
End Sub

Public Function FindCLNName(CLNCode As String) As String
    Dim CLNSql As String
    Dim rst3 As New ADODB.Recordset
    
    CLNSql = " select CLNCode, CLNNAme  from CLN where CLNCOde='" & Trim(CLNCode) & "'"
    
    rst3.Open CLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If rst3.RecordCount Then
        FindCLNName = rst3!CLNName
    Else
        MsgBox "Clinic Name not found!", vbCritical + vbOKOnly
    End If
    rst3.Close
    Set rst3 = Nothing
End Function

