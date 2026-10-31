VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmWorksheetAdj 
   Caption         =   "WorkSheet Adjustment"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtState 
      Height          =   285
      Left            =   2160
      TabIndex        =   17
      Text            =   "Text4"
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   1560
      TabIndex        =   16
      Text            =   "Text3"
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAdd2 
      Height          =   285
      Left            =   840
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   8160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAdd1 
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   15
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Worksheet Adjustment"
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
         Left            =   180
         TabIndex        =   13
         Top             =   60
         Width           =   6615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   11775
      Begin VB.TextBox txtCASNumber 
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
         Left            =   960
         TabIndex        =   3
         Top             =   150
         Width           =   1380
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdSearchCase 
         Caption         =   "&Search"
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lblClaimability 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10440
         TabIndex        =   10
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Claimability"
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
         Left            =   9600
         TabIndex        =   9
         Top             =   150
         Width           =   900
      End
      Begin VB.Label txtPolStatus 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6360
         TabIndex        =   8
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label48 
         Caption         =   "Pol Status"
         Height          =   255
         Left            =   5520
         TabIndex        =   7
         Top             =   150
         Width           =   795
      End
      Begin VB.Label Label9 
         Caption         =   "Case Status"
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
         Left            =   7560
         TabIndex        =   6
         Top             =   150
         Width           =   900
      End
      Begin VB.Label lblCASStatus 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8520
         TabIndex        =   5
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label46 
         Caption         =   "Case Num"
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
         TabIndex        =   4
         Top             =   150
         Width           =   1275
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   18
      Top             =   960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Invoices && Diagnosis"
      TabPicture(0)   =   "frmWorksheetAdj.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblHLTcode"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtMemType"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "GeneralFrame"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdExit2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&WorkSheet"
      TabPicture(1)   =   "frmWorksheetAdj.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "txtINVNo"
      Tab(1).Control(2)=   "txtBalLimit"
      Tab(1).Control(3)=   "txtINvDate"
      Tab(1).Control(4)=   "Fgrid"
      Tab(1).Control(5)=   "lblVisitLimit"
      Tab(1).Control(6)=   "Label12"
      Tab(1).Control(7)=   "Label10"
      Tab(1).Control(8)=   "Label11"
      Tab(1).Control(9)=   "Label1"
      Tab(1).Control(10)=   "lblCDNo"
      Tab(1).Control(11)=   "lblCDNoteNo"
      Tab(1).ControlCount=   12
      Begin VB.Frame Frame4 
         Height          =   495
         Left            =   -68760
         TabIndex        =   62
         Top             =   6480
         Width           =   5415
         Begin VB.CommandButton cmdExport 
            Caption         =   "Export"
            Height          =   255
            Left            =   1200
            TabIndex        =   68
            Top             =   165
            Width           =   1095
         End
         Begin VB.CommandButton cmdApprove 
            Caption         =   "&Save WS"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   165
            Width           =   1095
         End
         Begin VB.CommandButton cmdExit 
            Cancel          =   -1  'True
            Caption         =   "Exit"
            Height          =   255
            Left            =   4320
            TabIndex        =   66
            Top             =   165
            Width           =   975
         End
         Begin VB.CommandButton cmdCredit 
            Caption         =   "Credit Note"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2280
            TabIndex        =   65
            Top             =   165
            Width           =   975
         End
         Begin VB.CommandButton cmdDebit 
            Caption         =   "Debit Note"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3240
            TabIndex        =   64
            Top             =   165
            Width           =   1095
         End
         Begin VB.CommandButton cmdExcLetter 
            Caption         =   "Exc Letter"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4080
            TabIndex        =   63
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.TextBox txtINVNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73680
         TabIndex        =   61
         Top             =   420
         Width           =   6615
      End
      Begin VB.TextBox txtBalLimit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -66120
         TabIndex        =   60
         Top             =   420
         Width           =   1695
      End
      Begin VB.CommandButton cmdExit2 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10680
         TabIndex        =   59
         Top             =   6600
         Width           =   855
      End
      Begin VB.Frame GeneralFrame 
         Height          =   6135
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   11535
         Begin VB.Frame Frame1 
            Caption         =   "Invoices"
            Height          =   2295
            Left            =   120
            TabIndex        =   41
            Top             =   3720
            Width           =   11335
            Begin MSComctlLib.ListView lstInvoice 
               Height          =   1935
               Left            =   120
               TabIndex        =   42
               Top             =   240
               Width           =   11175
               _ExtentX        =   19711
               _ExtentY        =   3413
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1680
            Left            =   120
            TabIndex        =   39
            Top             =   2040
            Width           =   11355
            Begin MSComctlLib.ListView lstFinalDiag 
               Height          =   1335
               Left            =   120
               TabIndex        =   40
               Top             =   240
               Width           =   11145
               _ExtentX        =   19659
               _ExtentY        =   2355
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
                  Size            =   8.25
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
            TabIndex        =   38
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtMBMName 
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
            Height          =   285
            Left            =   1200
            TabIndex        =   37
            Top             =   240
            Width           =   3495
         End
         Begin VB.TextBox txtINSType 
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
            Height          =   285
            Left            =   5880
            TabIndex        =   36
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtPatientID 
            Enabled         =   0   'False
            Height          =   285
            Left            =   10900
            TabIndex        =   35
            Top             =   240
            Width           =   585
         End
         Begin VB.TextBox txtPatientName 
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
            Height          =   285
            Left            =   1200
            TabIndex        =   34
            Top             =   480
            Width           =   3495
         End
         Begin VB.TextBox txtClinic 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   33
            Top             =   720
            Width           =   3495
         End
         Begin VB.TextBox txtDisability 
            Enabled         =   0   'False
            Height          =   765
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   32
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txtCASRemark 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   31
            Top             =   1680
            Width           =   3495
         End
         Begin VB.TextBox txtPLNCode 
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
            TabIndex        =   29
            Top             =   720
            Width           =   1065
         End
         Begin VB.TextBox txtPolEffDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5880
            TabIndex        =   28
            Top             =   960
            Width           =   1065
         End
         Begin VB.TextBox txtExpDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5880
            TabIndex        =   27
            Top             =   1200
            Width           =   1065
         End
         Begin VB.TextBox txtICDCode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5880
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   1440
            Width           =   1065
         End
         Begin VB.TextBox txtVisitDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5880
            TabIndex        =   25
            Top             =   1680
            Width           =   1065
         End
         Begin VB.TextBox txtMBMPolNo 
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
            Height          =   285
            Left            =   8040
            TabIndex        =   24
            Top             =   480
            Width           =   3445
         End
         Begin VB.TextBox txtType 
            Enabled         =   0   'False
            Height          =   285
            Left            =   8040
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   720
            Width           =   3445
         End
         Begin VB.TextBox txtFolUpDate 
            Enabled         =   0   'False
            Height          =   765
            Left            =   8040
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   22
            Top             =   960
            Width           =   3445
         End
         Begin VB.Label Label2 
            Caption         =   "Mem Name"
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
            TabIndex        =   58
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label8 
            Caption         =   "Patient Name"
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
            TabIndex        =   57
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label Label5 
            Caption         =   "Plan Code"
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
            Left            =   4920
            TabIndex        =   56
            Top             =   480
            Width           =   795
         End
         Begin VB.Label Label47 
            Caption         =   "Eff Date"
            Height          =   255
            Left            =   4920
            TabIndex        =   55
            Top             =   975
            Width           =   1335
         End
         Begin VB.Label Label54 
            Caption         =   "Exp Date"
            Height          =   255
            Left            =   4920
            TabIndex        =   54
            Top             =   1215
            Width           =   1335
         End
         Begin VB.Label Label55 
            Caption         =   "Payor"
            Height          =   255
            Left            =   4920
            TabIndex        =   53
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label44 
            Caption         =   "Claim Types"
            Height          =   255
            Left            =   7080
            TabIndex        =   52
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label27 
            Caption         =   "Pol Num"
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
            Left            =   7080
            TabIndex        =   51
            Top             =   480
            Width           =   795
         End
         Begin VB.Label Label26 
            Caption         =   "Mem Num"
            Height          =   255
            Left            =   7080
            TabIndex        =   50
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "ICD Code"
            Height          =   255
            Left            =   4920
            TabIndex        =   49
            Top             =   1455
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Disability"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Fol-up Dates"
            Height          =   255
            Left            =   7080
            TabIndex        =   47
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label51 
            Caption         =   "Clinic"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label49 
            Caption         =   "DOV"
            Height          =   255
            Left            =   4920
            TabIndex        =   45
            Top             =   1695
            Width           =   1335
         End
         Begin VB.Label Label60 
            Caption         =   "Remark"
            Height          =   375
            Left            =   120
            TabIndex        =   44
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "INS Type"
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
            Left            =   4920
            TabIndex        =   43
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.TextBox txtMemType 
         Height          =   285
         Left            =   600
         TabIndex        =   20
         Top             =   5940
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtINvDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73680
         TabIndex        =   19
         Top             =   660
         Width           =   6615
      End
      Begin MSFlexGridLib.MSFlexGrid Fgrid 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   77
         Top             =   960
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   9763
         _Version        =   393216
         Rows            =   36
         Cols            =   9
         FixedCols       =   2
         BackColorFixed  =   -2147483638
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblVisitLimit 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -66120
         TabIndex        =   76
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Visit Limit"
         Height          =   255
         Left            =   -66960
         TabIndex        =   75
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Invoice Date"
         Height          =   255
         Left            =   -74880
         TabIndex        =   74
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Bal Limit"
         Height          =   255
         Left            =   -66960
         TabIndex        =   73
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Invoice Num"
         Height          =   255
         Left            =   -74880
         TabIndex        =   72
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label lblCDNo 
         BackColor       =   &H0080FFFF&
         Height          =   255
         Left            =   -64320
         TabIndex        =   71
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lblHLTcode 
         BackColor       =   &H0080FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   5940
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblCDNoteNo 
         Caption         =   "Credit No :"
         Height          =   255
         Left            =   -64320
         TabIndex        =   69
         Top             =   420
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmWorksheetAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset
Dim strsqlDiag As String
Dim Diag As New ADODB.Recordset
Dim DiagInd As String
Dim DiagREC As Integer
    BNFcombolisting
    strsql = "SELECT CASNumber, MBMStatus,  MBMNumber, PAycode, " & _
        " CoverID, MBMPolicyNo, MBMName, MBMCName , HLTCode, " & _
        " PLNCode, MBMPayorExpDate , CASStatus,Claimability," & _
        " MBMPayorEffDate , HLTCode ,ClinicCode," & _
        " MBMAvailableLimit,MBMCAvailableLimit, Claimability, CASStatus, " & _
        "  INSCOde ,  MBMAddress1, MBMAddress2, MBMMemberType, " & _
        " MBMAddress3, MBMPostCode, CTYCode, STACode,DateVisit, ANNLmtInd, VISLmtInd, " & _
        " MBMTelnoH, MBMTelNoM, MBMTelNoO, MBMCLNVisitLmt, MBMCLNBalLmt,MBMClinical "
    
    strsql = "SELECT CASNumber,  ANNLmtInd, VISLmtInd, " & _
        " MBMCLNVisitLmt, MBMCLNBalLmt,MBMClinical "
    strsql = strsql & " from VIEW_MBM_CAS_Claim where 0=0"
    
    If Trim(txtCASNumber) = "" Then
        MsgBox "Please select a Case Number ! ", vbCritical
        txtCASNumber.SetFocus
    'Else

    'ScanCode = Trim(txtCASNumber)
    'If Mid(ScanCode, 1, 2) = "LO" Then
    '    txtCASNumber = ""
    '    Call SearchForm(ScanCode)
    Else
        
        Screen.MousePointer = vbHourglass
        Call ClearEntries
        
        strsql = strsql & " And CASNumber = '" & Trim(txtCASNumber) & "'"
        CAS.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            If CAS.RecordCount Then
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
                
                    'txtPolStatus = PolStatus!pstatusDesc
                    lblHLTcode = CAS!HLTCode
                    txtPLNCode = CAS!PLNCode
                    txtINSType = CAS!INSCode
                
                If Len(CAS!MBMPayorEffDate) Then
                    txtPolEffDate = CAS!MBMPayorEffDate
                End If
                If Len(CAS!MBMPayorExpDate) Then
                    txtExpDate = CAS!MBMPayorExpDate
                End If
                If Len(CAS!PAYCode) Then
                    TxtPayor = CAS!PAYCode
                End If
                
                ' 'p= separated  c = combined
                If Len(CAS!MBMMemberType) Then
                    txtMemType = CAS!MBMMemberType
                End If
                
                If ANNInd = True Then
                    strsqlannlmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(txtPLNCode) & "'"
                    AnnLmt.Open strsqlannlmt, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
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
                End If
                
                If VISInd = True Then
                    lblVisitLimit = Format(CAS!MBMCLNVisitLmt, "#,#.00")
                End If
                If Len(CAS!DateVisit) Then
                    txtVisitDate = CAS!DateVisit
                End If
                If Len(Trim(CAS!ClinicCode)) <> 0 Then
                    CLNCode = CAS!ClinicCode
                    txtClinic = FindCLNName(CAS!ClinicCode)
                End If
            'If frmInvoice.lstInvoice.ListItems.count > 0 Then
            '    PLNCode = frmInvoice.lstInvoice.SelectedItem.SubItems(67)
            '    INSCode = frmInvoice.lstInvoice.SelectedItem.SubItems(68)
            '    HLTCode = frmInvoice.lstInvoice.SelectedItem.SubItems(69)
            '    CLNCode = frmInvoice.lstInvoice.SelectedItem.SubItems(12)
            '    CLNBnfCard = frmInvoice.lstInvoice.SelectedItem.SubItems(72)
            '    BalLimit = frmInvoice.lstInvoice.SelectedItem.SubItems(70)
            '    VisitLimit = frmInvoice.lstInvoice.SelectedItem.SubItems(71)
             '   ANNInd = frmInvoice.lstInvoice.SelectedItem.SubItems(74)
             '   VISInd = frmInvoice.lstInvoice.SelectedItem.SubItems(73)
            'End If
            
            Call showInvoice
            
            'strsqlDiag = " select PLNCode  from DiagNotPayable "
            'strsqlDiag = strsqlDiag & " where  PLNCode = '" & Trim(txtPLNCode) & "'"
            'Diag.Open strsqlDiag, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
            'If Diag.RecordCount = 0 Then
            Call display_benefit
            'End If
            'Diag.Close
            'Set Diag = Nothing
            
            Call display_CLNbenefit
            Call CalculateReim
            Call CalculateCashless
            Call getDatatoWS
        
            
        Else
            MsgBox "Record not found! ", vbCritical
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrorShow:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description

End Sub

Private Sub ClearEntries()
    Dim ctl As Control
    Dim FCol, Frow As Integer
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Name <> "txtCASNumber" Then
                ctl.Text = ""
            End If
        End If
    Next
    'lstInvoice.ListItems.Clear
    For FCol = 3 To 8
        Fgrid.Col = FCol
        For Frow = 1 To 29
            Fgrid.Row = Frow
            Fgrid = "-"
        Next
    Next
        For FCol = 2 To 8
        Fgrid.Col = FCol
        
        For Frow = 31 To 35
            Fgrid.CellFontBold = True
            Fgrid.CellForeColor = vbBlue
            Fgrid.Row = Frow
            Fgrid = "-"
        Next
    Next
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
                
                If Trim(Invoice!INVNo) <> "" Then
                    .SubItems(4) = Invoice!INVNo
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
                If Trim(Invoice!CLMXRay) <> "" Then
                    .SubItems(16) = Invoice!CLMXRay
                    I6 = I6 + Invoice!CLMXRay
                End If
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
            End With
            Invoice.MoveNext
        Loop
    End If
    Invoice.Close
    Set Invoice = Nothing
End Sub

