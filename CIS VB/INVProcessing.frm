VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmINVProcessing 
   Caption         =   "MedixCIS - (Invoice Processing)"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView LstInvProcess 
      Height          =   2415
      Left            =   0
      TabIndex        =   36
      Top             =   2520
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Case No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Medicine Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Procedure Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Consult Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "X-Ray Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Lab Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Co Payment"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Final Bill"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Bill Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Annual MC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "MC Days"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Balance MC"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   0
      TabIndex        =   21
      Top             =   360
      Width           =   6375
      Begin VB.TextBox txtCASNo 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtfinalBill 
         Height          =   315
         Left            =   4320
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtMedFee 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtproFee 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtconFee 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtxrayFee 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtlabFee 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtCoPayment 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   1680
         Width           =   1815
      End
      Begin MSMask.MaskEdBox MskBillDate 
         Height          =   315
         Left            =   4320
         TabIndex        =   9
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txttotal 
         Height          =   315
         Left            =   4320
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtAnnMC 
         Height          =   315
         Left            =   4320
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtMCday 
         Height          =   315
         Left            =   4320
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtBalMC 
         Height          =   315
         Left            =   4320
         TabIndex        =   13
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Bill Date"
         Height          =   255
         Left            =   3240
         TabIndex        =   37
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Case Number"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Fee"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Consult Fee"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Procedure Fee"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "X-Ray Fee"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Lab Fee"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Co Payment"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Final Bill"
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Total"
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Annual MC"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "MC Days"
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Balance MC"
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   4920
      Width           =   6375
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   375
         Left            =   2280
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   4800
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5520
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   4080
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   11895
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Invoice Processing"
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
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   0
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmINVProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command6_Click()

End Sub

Private Sub CmdExit_Click()
End
End Sub

