VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPayMBMPost 
   Caption         =   "Member Payment Posting"
   ClientHeight    =   8592
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8592
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   810
      Left            =   0
      TabIndex        =   13
      Top             =   240
      Width           =   11895
      Begin VB.ComboBox cboPAYCode 
         Height          =   315
         Left            =   1080
         TabIndex        =   17
         Top             =   180
         Width           =   3495
      End
      Begin VB.TextBox TxtINVNoFr 
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   470
         Width           =   1455
      End
      Begin VB.TextBox TxtINVNoTo 
         Height          =   285
         Left            =   3000
         TabIndex        =   15
         Top             =   470
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtINVRegDateTo 
         Height          =   285
         Left            =   7320
         TabIndex        =   18
         Top             =   180
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtINVRegDateFr 
         Height          =   285
         Left            =   5880
         TabIndex        =   19
         Top             =   180
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboPostStatus 
         Height          =   315
         Left            =   5880
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   440
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Posting Status"
         Height          =   255
         Left            =   4800
         TabIndex        =   25
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Inv No from"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   500
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "INV Reg Date"
         Height          =   255
         Left            =   4800
         TabIndex        =   21
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   7080
         TabIndex        =   20
         Top             =   195
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   7320
      Width           =   11775
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   9480
         TabIndex        =   6
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   8160
         TabIndex        =   5
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   7440
         TabIndex        =   4
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10440
         TabIndex        =   3
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   11040
         TabIndex        =   2
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "&Post"
         Height          =   375
         Left            =   8760
         TabIndex        =   1
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lblCreditAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   4800
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblDebitAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Credit Amt"
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Debit Amt"
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstMBMPosting 
      Height          =   6015
      Left            =   0
      TabIndex        =   26
      Top             =   1320
      Width           =   11895
      _ExtentX        =   20976
      _ExtentY        =   10605
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "A/C Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Based Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Debit/Credit"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Trans Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Journal Type"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Trans Ref"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Desc"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "T1 Clinic"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "T2 Plan No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "T3 Pay Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "T4  InvDORMth"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "T5 Grp Com"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "T6 Bordx Ref"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Status"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Date"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Member Payment Posting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   11925
   End
   Begin VB.Label Label38 
      Caption         =   "Member Posting List"
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
      TabIndex        =   27
      Top             =   1080
      Width           =   2295
   End
End
Attribute VB_Name = "frmPayMBMPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call ActionLogUpdate(12, "frmPayMBMPost")

End Sub
