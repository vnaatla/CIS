VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPayment 
   Caption         =   "Payment Tracking"
   ClientHeight    =   8592
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8592
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2505
      Left            =   120
      TabIndex        =   42
      Top             =   720
      Width           =   11655
      Begin VB.CheckBox ChkClinic 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11115
         TabIndex        =   44
         Top             =   255
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpComp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   43
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox TxtClaimNo1 
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   2
         Top             =   240
         Width           =   1600
      End
      Begin VB.TextBox TxtClaimNo2 
         Height          =   285
         Left            =   3700
         MaxLength       =   15
         TabIndex        =   3
         Top             =   240
         Width           =   1600
      End
      Begin VB.ComboBox CboGrpCompany 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   480
         Width           =   4000
      End
      Begin VB.TextBox TxtMBMNo 
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   5
         Top             =   720
         Width           =   4000
      End
      Begin VB.TextBox txtPVNoTo 
         Height          =   285
         Left            =   3700
         MaxLength       =   15
         TabIndex        =   7
         Top             =   960
         Width           =   1600
      End
      Begin VB.TextBox txtPVNoFR 
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   6
         Top             =   960
         Width           =   1600
      End
      Begin VB.TextBox TxtInvoiceNo1 
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1215
         Width           =   1600
      End
      Begin VB.TextBox TxtInvoiceNo2 
         Height          =   285
         Left            =   3700
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1215
         Width           =   1600
      End
      Begin VB.TextBox TxtBordxNo1 
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1470
         Width           =   1600
      End
      Begin VB.TextBox TxtBordxNo2 
         Height          =   285
         Left            =   3700
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1470
         Width           =   1600
      End
      Begin VB.TextBox txtOPBatchNo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1725
         Width           =   1600
      End
      Begin VB.TextBox txtOPBatchNo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3700
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1725
         Width           =   1600
      End
      Begin VB.ComboBox CboPayment 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1980
         Width           =   4000
      End
      Begin VB.ComboBox CboClinic 
         Height          =   315
         Left            =   7365
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   3700
      End
      Begin MSMask.MaskEdBox txtPVDate2 
         Height          =   315
         Left            =   9465
         TabIndex        =   17
         Top             =   525
         Width           =   1605
         _ExtentX        =   2836
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPVDate1 
         Height          =   315
         Left            =   7365
         TabIndex        =   16
         Top             =   525
         Width           =   1605
         _ExtentX        =   2836
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDate1 
         Height          =   315
         Left            =   7365
         TabIndex        =   18
         Top             =   810
         Width           =   1605
         _ExtentX        =   2836
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDate2 
         Height          =   315
         Left            =   9465
         TabIndex        =   19
         Top             =   810
         Width           =   1605
         _ExtentX        =   2836
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtAmtTo 
         Height          =   285
         Left            =   9465
         MaxLength       =   8
         TabIndex        =   21
         Top             =   1080
         Width           =   1600
      End
      Begin VB.TextBox txtAmtFr 
         Height          =   285
         Left            =   7365
         MaxLength       =   8
         TabIndex        =   20
         Top             =   1080
         Width           =   1600
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         Height          =   315
         ItemData        =   "frmPayment.frx":0000
         Left            =   9465
         List            =   "frmPayment.frx":0002
         TabIndex        =   23
         Top             =   1320
         Width           =   1600
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         Height          =   315
         ItemData        =   "frmPayment.frx":0004
         Left            =   7365
         List            =   "frmPayment.frx":0006
         TabIndex        =   22
         Top             =   1320
         Width           =   1600
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         Height          =   315
         ItemData        =   "frmPayment.frx":0008
         Left            =   7365
         List            =   "frmPayment.frx":000A
         TabIndex        =   24
         Top             =   1605
         Width           =   1600
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         Height          =   315
         ItemData        =   "frmPayment.frx":000C
         Left            =   9465
         List            =   "frmPayment.frx":000E
         TabIndex        =   25
         Top             =   1605
         Width           =   1600
      End
      Begin VB.Label Label19 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   11160
         TabIndex        =   70
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label36 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5560
         TabIndex        =   69
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "Pymt Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   2085
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "A/C Period Mth"
         Height          =   255
         Left            =   6060
         TabIndex        =   67
         Top             =   1395
         Width           =   1200
      End
      Begin VB.Label Label22 
         Caption         =   "To"
         Height          =   255
         Left            =   9105
         TabIndex        =   66
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "Inv. Date"
         Height          =   255
         Left            =   6060
         TabIndex        =   65
         Top             =   855
         Width           =   1575
      End
      Begin VB.Label Label35 
         Caption         =   "OP Batch No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "To"
         Height          =   255
         Left            =   3200
         TabIndex        =   63
         Top             =   1760
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Mem No"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   810
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   9105
         TabIndex        =   61
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "PV Date"
         Height          =   255
         Left            =   6060
         TabIndex        =   60
         Top             =   585
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   3200
         TabIndex        =   59
         Top             =   1250
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Invoice No"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   3200
         TabIndex        =   57
         Top             =   1000
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "PV No"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         Height          =   255
         Left            =   9105
         TabIndex        =   55
         Top             =   1110
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Inv. Amt"
         Height          =   255
         Left            =   6060
         TabIndex        =   54
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Clinic"
         Height          =   255
         Left            =   6060
         TabIndex        =   53
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Group Company"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   555
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "A/C Period Yr"
         Height          =   255
         Left            =   6060
         TabIndex        =   51
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label Label24 
         Caption         =   "To"
         Height          =   255
         Left            =   9105
         TabIndex        =   50
         Top             =   1410
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "To"
         Height          =   255
         Left            =   9105
         TabIndex        =   49
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "Claim No"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   3200
         TabIndex        =   47
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "To"
         Height          =   255
         Left            =   3200
         TabIndex        =   45
         Top             =   1520
         Width           =   375
      End
   End
   Begin VB.ComboBox CboReportType 
      Height          =   315
      ItemData        =   "frmPayment.frx":0010
      Left            =   1440
      List            =   "frmPayment.frx":001A
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.ComboBox CboRptSelection 
      Height          =   315
      ItemData        =   "frmPayment.frx":003C
      Left            =   7320
      List            =   "frmPayment.frx":003E
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   27
      Top             =   7320
      Width           =   11655
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         MaxLength       =   15
         TabIndex        =   72
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   71
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   375
         Left            =   7200
         TabIndex        =   34
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   375
         Left            =   7920
         TabIndex        =   33
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   8640
         TabIndex        =   32
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TxtPStatus 
         Height          =   285
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10440
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9600
         TabIndex        =   29
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   5520
         TabIndex        =   26
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   6360
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1680
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   2880
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label39 
         Caption         =   "* - Listing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   600
         Width           =   2295
      End
   End
   Begin MSComctlLib.ListView lstPaymentTracking 
      Height          =   3975
      Left            =   120
      TabIndex        =   38
      Top             =   3360
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   7006
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PV No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "PV Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Amt Paid"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Invoice No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Invoice Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Invoice Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Payee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Member No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Claim No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Bordx No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "OP Batch No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Remarks"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Pymt Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "A/C Period"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label76 
      Caption         =   "Report Selection"
      Height          =   255
      Left            =   6000
      TabIndex        =   41
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label26 
      Caption         =   "Report Type"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Payment Tracking"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Public PaymentTrackingCriteria As String
Private AA As Integer
Private m_blnDirection(1 To 14) As Boolean

Private Sub CboRptSelection_Change()
    
    If CboReportType.ListIndex = 0 Then
        Call ListViewHeader1
    ElseIf CboReportType.ListIndex = 1 Then
        Call ListViewHeader2
    End If
    
End Sub

Private Sub CboRptSelection_Click()

    If CboReportType.ListIndex = 0 Then
        Call ListViewHeader1
    ElseIf CboReportType.ListIndex = 1 Then
        Call ListViewHeader2
    End If

End Sub

Private Sub CboRptSelection_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboRptSelection_KeyPress"
On Error GoTo debugs

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRptSelection.Text, CboRptSelection.SelStart) + Chr(KeyAscii) + Mid(CboRptSelection.Text, CboRptSelection.SelStart + CboRptSelection.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To CboRptSelection.ListCount - 1
        If NewText = LCase(Left(CboRptSelection.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboRptSelection.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboRptSelection.Text = ValidValue
                CboRptSelection.SelStart = 0
                CboRptSelection.SelLength = Len(ValidValue)
            End If
        End If
        
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub CboReportType_Click()
    CboRptSelection.Clear
    
    If CboReportType.ListIndex = 0 Then 'Pay to Clinic
        CboRptSelection.AddItem "By Claim"
        CboRptSelection.AddItem "By Bordx"
        CboRptSelection.AddItem "By PV No"
        CboRptSelection.AddItem "By Clinic"
        CboRptSelection.AddItem "By Group Company"
        CboRptSelection.AddItem "General Listing"
        CboRptSelection.Text = "General Listing"
        
        Call ListViewHeader1
    
    ElseIf CboReportType.ListIndex = 1 Then 'Pay to Member
        CboRptSelection.AddItem "By Claim"
        CboRptSelection.AddItem "By Bordx"
        CboRptSelection.AddItem "By PV No"
        CboRptSelection.AddItem "By Clinic"
        CboRptSelection.AddItem "By Group Company"
        CboRptSelection.AddItem "General Listing"
        CboRptSelection.Text = "General Listing"
        
        Call ListViewHeader2
        
    End If
End Sub

Private Sub CboReportType_KeyDown(KeyCode As Integer, Shift As Integer)
    CboRptSelection.Clear
    
    If CboReportType.ListIndex = 0 Then 'Pay to Clinic
        CboRptSelection.AddItem "By Claim"
        CboRptSelection.AddItem "By Bordx"
        CboRptSelection.AddItem "By PV No"
        CboRptSelection.AddItem "By Clinic"
        CboRptSelection.AddItem "By Group Company"
        CboRptSelection.AddItem "General Listing"
        CboRptSelection.Text = "General Listing"
        
        Call ListViewHeader1
    
    ElseIf CboReportType.ListIndex = 1 Then 'Pay to Member
        CboRptSelection.AddItem "By Claim"
        CboRptSelection.AddItem "By Bordx"
        CboRptSelection.AddItem "By PV No"
        CboRptSelection.AddItem "By Clinic"
        CboRptSelection.AddItem "By Group Company"
        CboRptSelection.AddItem "General Listing"
        CboRptSelection.Text = "General Listing"
        
        Call ListViewHeader2
    
    End If
End Sub

Private Sub CboReportType_KeyUp(KeyCode As Integer, Shift As Integer)
    CboRptSelection.Clear
    
    If CboReportType.ListIndex = 0 Then 'Pay to Clinic
        CboRptSelection.AddItem "By Claim"
        CboRptSelection.AddItem "By Bordx"
        CboRptSelection.AddItem "By PV No"
        CboRptSelection.AddItem "By Clinic"
        CboRptSelection.AddItem "By Group Company"
        CboRptSelection.AddItem "General Listing"
        CboRptSelection.Text = "General Listing"
        
        Call ListViewHeader1
    
    ElseIf CboReportType.ListIndex = 1 Then 'Pay to Member
        CboRptSelection.AddItem "By Claim"
        CboRptSelection.AddItem "By Bordx"
        CboRptSelection.AddItem "By PV No"
        CboRptSelection.AddItem "By Clinic"
        CboRptSelection.AddItem "By Group Company"
        CboRptSelection.AddItem "General Listing"
        CboRptSelection.Text = "General Listing"
        
        Call ListViewHeader2
        
    End If
    
End Sub

Private Sub CboReportType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboReportType.Text, CboReportType.SelStart) + Chr(KeyAscii) + Mid(CboReportType.Text, CboReportType.SelStart + CboReportType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboReportType.ListCount - 1
 
        If NewText = LCase(Left(CboReportType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboReportType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboReportType.Text = ValidValue
                CboReportType.SelStart = 0
                CboReportType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub ChkClinic_Click()
Const METHOD_NAME = "ChkClinic_Click"
On Error GoTo debugs
Dim ClnRec As New ADODB.Recordset
Dim StrSql As String
    CboClinic.Clear
    If ChkClinic.Value = 1 Then
        StrSql = "EXEC SearchCLN"
        ClnRec.Open StrSql, MediNeptuneConnCISWS, adOpenForwardOnly, adLockOptimistic
        Do Until ClnRec.EOF
            If Len(ClnRec!CLNName) Then
                CboClinic.AddItem ClnRec!CLNCode & " - " & Trim(ClnRec!CLNName)
            End If
            ClnRec.MoveNext
        Loop
        ClnRec.Close
        Set ClnRec = Nothing
    End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub ChkGrpComp_Click()
Const METHOD_NAME = "ChkGrpComp_Click"
On Error GoTo debugs

Dim Company As New ADODB.Recordset
Dim StrSql As String
    
CboGrpCompany.Clear
If ChkGrpComp.Value = 1 Then
    strAppend = ""
    StrSql = "EXEC PaymentReceiptModule '" & strAppend & "',16"
    Company.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockOptimistic
    
    Do Until Company.EOF
        CboGrpCompany.AddItem Trim(Company!GrpCoName)
        Company.MoveNext
    Loop
    Company.Close
    Set Company = Nothing
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdBack_Click()
    Frame1.Visible = True
    Label26.Visible = True
    CboReportType.Visible = True
    Label76.Visible = True
    CboRptSelection.Visible = True
    lstPaymentTracking.Top = 3360
    lstPaymentTracking.Left = 120
    lstPaymentTracking.Width = 11655
    lstPaymentTracking.Height = 3975
End Sub

Private Sub CmdView_Click()
    
    Frame1.Visible = False
    Label26.Visible = False
    CboReportType.Visible = False
    Label76.Visible = False
    CboRptSelection.Visible = False
    lstPaymentTracking.Top = 360
    lstPaymentTracking.Left = 120
    lstPaymentTracking.Width = 11655
    lstPaymentTracking.Height = 6975
    
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

Private Sub cmdExport_Click()
        
    Screen.MousePointer = vbHourglass
    
    If lstPaymentTracking.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstPaymentTracking)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
            
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "FrmPayment")
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    lstPaymentTracking.Visible = True
    Call ComboListing
End Sub

'******************* Start Combo Listing ********************************

Private Sub ComboListing()
Const METHOD_NAME = "ComboListing"
On Error GoTo debugs

Dim monthstart
Dim yearstart
    For monthstart = 1 To 12
        With cboACPeriodMthFr
            .AddItem monthstart
        End With
        With cboACPeriodMthTo
            .AddItem monthstart
        End With
    Next monthstart
    For yearstart = 1999 To 3000
        With cboACPeriodYRFr
            .AddItem yearstart
        End With
        With cboACPeriodYRTo
            .AddItem yearstart
        End With
    Next yearstart
    With CboPayment
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Partial Paid"
    End With
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub
'******************* End Combo Listing ********************************

'******************* Start Command Button *******************************

Private Sub CmdSearch_Click()
Const METHOD_NAME = "CmdSearch_Click"
On Error GoTo errRun

    Screen.MousePointer = vbHourglass
                
    If CboReportType.Text <> "" Then
        
        CmdClear.Enabled = False
        CmdExit.Enabled = False
        CmdExport.Enabled = False
        CmdSearch.Enabled = False
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
            Call SearchRecord
      '  MediConnCISDB.Close
      '  Set MediConnCISDB = Nothing
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
      '  MediConnCISWS.Close
      '  Set MediConnCISWS = Nothing
      '  MediConnCISPayment.Close
       ' Set MediConnCISPayment = Nothing
        
    Else
        MsgBox "Please select a report type.", vbInformation
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdExport.Enabled = True
        CmdSearch.Enabled = True
    End If
    
    Screen.MousePointer = Default
    Exit Sub
errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lblCurrent = ""
    LblTotalAmt = ""
    lstPaymentTracking.ListItems.Clear
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdExport.Enabled = True
    CmdSearch.Enabled = True
    
    If Err.Description <> "" Then
        Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
    End If

    
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdClear_Click()
    TxtClaimNo1 = ""
    TxtClaimNo2 = ""
    CboGrpCompany.Text = ""
    TxtMBMNo = ""
    txtPVNoFR = ""
    txtPVNoTo = ""
    TxtInvoiceNo1 = ""
    TxtInvoiceNo2 = ""
    TxtBordxNo1 = ""
    TxtBordxNo2 = ""
    txtOPBatchNo1 = ""
    txtOPBatchNo2 = ""
    CboClinic.Text = ""
    txtPVDate1 = "__/__/____"
    txtPVDate2 = "__/__/____"
    txtInvDate1 = "__/__/____"
    txtInvDate2 = "__/__/____"
    txtAmtFr = ""
    TxtAmtTo = ""
    cboACPeriodMthFr = ""
    cboACPeriodMthTo = ""
    cboACPeriodYRFr = ""
    cboACPeriodYRTo = ""
    lstPaymentTracking.ListItems.Clear
    lblCurrent = 0
    LblTotalAmt = 0
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
End Sub
'******************* End Command Button *******************************

'******************* Start Search Record*******************************

Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim StrSql As String
Dim strAppend As String
Dim sql As String
    
    strAppend = ""
        
    
    If Len(Trim(TxtClaimNo1)) Then
        strAppend = strAppend + " And CASNumber >= '" & Trim(TxtClaimNo1) & "'"
    End If
    
    If Len(Trim(TxtClaimNo2)) Then
        strAppend = strAppend + " And CASNumber <= '" & Trim(TxtClaimNo2) & "'"
    End If
    
    If Len(Trim(CboGrpCompany)) Then
        strAppend = strAppend + " And GroupCompany = '" & RemStr(Trim(CboGrpCompany)) & "'"
    End If
    
    If Len(Trim(TxtMBMNo)) Then
        strAppend = strAppend + " And MBMNumber = '" & ChkStr(TxtMBMNo) & "'"
    End If
    
    If Len(Trim(TxtInvoiceNo1)) Then
        strAppend = strAppend + " AND InvoiceNo >= '" & Trim(TxtInvoiceNo1) & "'"
    End If
       
    If Len(Trim(TxtInvoiceNo2)) Then
        strAppend = strAppend + " AND InvoiceNo <= '" & Trim(TxtInvoiceNo2) & "'"
    End If
    
    If Len(Trim(TxtBordxNo1)) Then
        strAppend = strAppend + " And BordxRefNo >= '" & Trim(TxtBordxNo1) & "'"
    End If
    
    If Len(Trim(TxtBordxNo2)) Then
        strAppend = strAppend + " And BordxRefNo <= '" & Trim(TxtBordxNo2) & "'"
    End If
    
    If Len(Trim(txtOPBatchNo1)) Then
        strAppend = strAppend + " And OPBatchNo >= '" & Trim(txtOPBatchNo1) & "'"
    End If
    
    If Len(Trim(txtOPBatchNo2)) Then
        strAppend = strAppend + " And OPBatchNo <= '" & Trim(txtOPBatchNo2) & "'"
    End If
    
    If Len(Trim(CboClinic)) Then
        strAppend = strAppend + " AND CLinicCode = '" & Mid(CboClinic, 1, 4) & "'"
    End If
    
    If Len(Trim(txtInvDate1)) > 0 And IsDate(txtInvDate1) = True Then
        strAppend = strAppend + " And InvoiceDate >= '" & Format(txtInvDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtInvDate2)) > 0 And IsDate(txtInvDate2) = True Then
        strAppend = strAppend + " And InvoiceDate <= '" & Format(txtInvDate2, "mm/dd/yyyy") & "'"
    End If
    
    
    If CboReportType.ListIndex = 0 Then
        
        If Mid(CboPayment, 1, 1) = 1 Then
            strAppend = strAppend & " and PVPayCLNStatus = 'F'"
        ElseIf Mid(CboPayment, 1, 1) = 2 Then
            strAppend = strAppend & " and PVPayCLNStatus = 'P'"
        End If
        
        If Len(Trim(txtPVNoFR)) Then
            strAppend = strAppend + " AND PVCLNNo >= '" & RemStr(Trim(txtPVNoFR)) & "'"
        End If
           
        If Len(Trim(txtPVNoTo)) Then
            strAppend = strAppend + " AND PVCLNNo <= '" & RemStr(Trim(txtPVNoTo)) & "'"
        End If
                
        If Len(Trim(txtPVDate1)) > 0 And IsDate(txtPVDate1) = True Then
            strAppend = strAppend + " And PVCLNDate >= '" & Format(txtPVDate1, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtPVDate2)) > 0 And IsDate(txtPVDate2) = True Then
            strAppend = strAppend + " And PVCLNDate <= '" & Format(txtPVDate2, "mm/dd/yyyy") & "'"
        End If
               
        If Len(txtAmtFr) Then
            strAppend = strAppend + " AND CLMPayableByMedix2C >= " & Trim(txtAmtFr) & ""
        End If
        
        If Len(TxtAmtTo) Then
            strAppend = strAppend + " AND CLMPayableByMedix2C <= " & Trim(TxtAmtTo) & ""
        End If
        
        If Len(Trim(cboACPeriodMthFr)) Then
            strAppend = strAppend + " And CLNAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
        End If
        If Len(Trim(cboACPeriodMthTo)) Then
            strAppend = strAppend + " And CLNAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
        End If
        
        If Len(Trim(cboACPeriodYRFr)) Then
            strAppend = strAppend + " And CLNAccPeriodYR >= '" & Trim(cboACPeriodYRFr) & "'"
        End If
        If Len(Trim(cboACPeriodYRTo)) Then
            strAppend = strAppend + " And CLNAccPeriodYR <= '" & Trim(cboACPeriodYRTo) & "'"
        End If

    Else
    
        If Mid(CboPayment, 1, 1) = 1 Then
            strAppend = strAppend & " and PVPayMBMStatus = 'F'"
        ElseIf Mid(CboPayment, 1, 1) = 2 Then
            strAppend = strAppend & " and PVPayMBMStatus = 'P'"
        End If
        
        If Len(Trim(txtPVNoFR)) Then
            strAppend = strAppend + " AND PVNo >= '" & RemStr(Trim(txtPVNoFR)) & "'"
        End If
           
        If Len(Trim(txtPVNoTo)) Then
            strAppend = strAppend + " AND PVNo <= '" & RemStr(Trim(txtPVNoTo)) & "'"
        End If
                
        If Len(Trim(txtPVDate1)) > 0 And IsDate(txtPVDate1) = True Then
            strAppend = strAppend + " And PVDate >= '" & Format(txtPVDate1, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtPVDate2)) > 0 And IsDate(txtPVDate2) = True Then
            strAppend = strAppend + " And PVDate <= '" & Format(txtPVDate2, "mm/dd/yyyy") & "'"
        End If
               
        If Len(txtAmtFr) Then
            strAppend = strAppend + " AND CLMPayableByMedix2M >= " & Trim(txtAmtFr) & ""
        End If
        
        If Len(TxtAmtTo) Then
            strAppend = strAppend + " AND CLMPayableByMedix2M <= " & Trim(TxtAmtTo) & ""
        End If
        
        If Len(Trim(cboACPeriodMthFr)) Then
            strAppend = strAppend + " And PAYAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
        End If
        If Len(Trim(cboACPeriodMthTo)) Then
            strAppend = strAppend + " And PAYAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
        End If
        
        If Len(Trim(cboACPeriodYRFr)) Then
            strAppend = strAppend + " And PAYAccPeriodYR >= '" & Trim(cboACPeriodYRFr) & "'"
        End If
        If Len(Trim(cboACPeriodYRTo)) Then
            strAppend = strAppend + " And PAYAccPeriodYR <= '" & Trim(cboACPeriodYRTo) & "'"
        End If
    
    End If
    
    If CboReportType.ListIndex = 0 And CboRptSelection.Text = "General Listing" Then
        Call PaymentCLNTrackingListing(strAppend)
    ElseIf CboReportType.ListIndex = 0 And CboRptSelection.Text = "By Bordx" Then
        Call PaymentCLNTrackingListing2(strAppend)
    ElseIf CboReportType.ListIndex = 0 And CboRptSelection.Text = "By Claim" Then
        Call PaymentCLNTrackingListing3(strAppend)
    ElseIf CboReportType.ListIndex = 0 And CboRptSelection.Text = "By Clinic" Then
        Call PaymentCLNTrackingListing4(strAppend)
    ElseIf CboReportType.ListIndex = 0 And CboRptSelection.Text = "By Group Company" Then
        Call PaymentCLNTrackingListing5(strAppend)
    ElseIf CboReportType.ListIndex = 0 And CboRptSelection.Text = "By PV No" Then
        Call PaymentCLNTrackingListing6(strAppend)
    ElseIf CboReportType.ListIndex = 1 And CboRptSelection.Text = "General Listing" Then
        Call PaymentMBMTrackingListing(strAppend)
    ElseIf CboReportType.ListIndex = 1 And CboRptSelection.Text = "By Bordx" Then
        Call PaymentMBMTrackingListing2(strAppend)
    ElseIf CboReportType.ListIndex = 1 And CboRptSelection.Text = "By Claim" Then
        Call PaymentMBMTrackingListing3(strAppend)
    ElseIf CboReportType.ListIndex = 1 And CboRptSelection.Text = "By Clinic" Then
        Call PaymentMBMTrackingListing4(strAppend)
    ElseIf CboReportType.ListIndex = 1 And CboRptSelection.Text = "By Group Company" Then
        Call PaymentMBMTrackingListing5(strAppend)
    ElseIf CboReportType.ListIndex = 1 And CboRptSelection.Text = "By PV No" Then
        Call PaymentMBMTrackingListing6(strAppend)
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
   
End Function
'******************* End Search Record*******************************

'******************* Start Payment MBM Tracking Listing *******************************
Private Sub PaymentMBMTrackingListing(Optional strAppend As String)

Const METHOD_NAME = "PaymentMBMTrackingListing"
On Error GoTo debugs


Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    total = 0
    Current = 0
    
    lstPaymentTracking.ListItems.Clear
    
    strAppendCase = "10"
    
    sql = sql + strAppend
            
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    'Sql = " Select * From VIEW_PaymentMBMTracking where 0 = 0 "
     
    'If Len(Trim(StrAppend)) Then
        'Sql = Sql + StrAppend
    'Else
        'Sql = Sql
    'End If
    
    'rst2.Open Sql, MediConnCISPayment, adOpenForwardOnly, adLockBatchOptimistic
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdExport.Enabled = True
    Else
        Do Until rst2.EOF
        TxtPStatus = ""
        If Trim(rst2!PVPayMBMStatus) <> "" Then
            TxtPStatus = Trim(rst2!PVPayMBMStatus)
        End If
        
            With rst2
                Set Record = lstPaymentTracking.ListItems.Add(, , !PVNo)
                If Len(!PVDate) Then
                   Record.SubItems(1) = Format(!PVDate, "dd/mm/yyyy")
                End If
                If Len(!MBMPaidAmt) Then
                   Record.SubItems(2) = Format(!MBMPaidAmt, "#,##0.00")
                End If
                If Len(!InvoiceNo) Then
                   Record.SubItems(3) = Trim(!InvoiceNo)
                End If
                If Len(!InvoiceDate) Then
                   Record.SubItems(4) = Format(!InvoiceDate, "dd/mm/yyyy")
                End If
                If Len(!CLMPayableByMedix2M) Then
                   Record.SubItems(5) = Format(!CLMPayableByMedix2M, "#,##0.00")
                End If
                If Len(!PVPayto) Then
                    Record.SubItems(6) = Trim(!PVPayto)
                End If
                If Len(!MBMNumber) Then
                    Record.SubItems(7) = !MBMNumber
                End If
                If Len(!CASNumber) Then
                    Record.SubItems(8) = Trim(!CASNumber)
                End If
                If Len(!Bordxrefno) Then
                    Record.SubItems(9) = !Bordxrefno
                End If
                If Len(!OPBatchNo) Then
                    Record.SubItems(10) = !OPBatchNo
                End If
                If Len(!Remark) Then
                    Record.SubItems(11) = Trim(!Remark)
                End If
                
                If TxtPStatus = "F" Then
                    Record.SubItems(12) = "Full"
                ElseIf TxtPStatus = "P" Then
                    Record.SubItems(12) = "Partial"
                Else
                    Record.SubItems(12) = ""
                End If
                
                If Len(!PAYAccPeriodMth & "/" & !PAYAccPeriodYr) Then
                    Record.SubItems(13) = (!PAYAccPeriodMth & "/" & !PAYAccPeriodYr)
                End If
                                                              
                       
                Current = Current + 1
                lblCurrent = Current
                If Len(!MBMPaidAmt) Then
                    TotalAmt = TotalAmt + !MBMPaidAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdExport.Enabled = True
    CmdSearch.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub PaymentMBMTrackingListing2(Optional strAppend As String)
Const METHOD_NAME = "PaymentMBMTrackingListing2"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    total = 0
    Current = 0
    
    lstPaymentTracking.ListItems.Clear
    
    strAppendCase = "11"
    
    sql = sql + strAppend
            
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdExport.Enabled = True
    Else
        Do Until rst2.EOF
                
            With rst2
                If Len(!Bordxrefno) Then
                    Set Record = lstPaymentTracking.ListItems.Add(, , !Bordxrefno)
                Else
                    Set Record = lstPaymentTracking.ListItems.Add(, , "Empty")
                End If
                If Len(!BordxDate) Then
                    Record.SubItems(1) = Format(!BordxDate, "dd/mm/yyyy")
                Else
                    Record.SubItems(1) = "Empty"
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(2) = !NoofCase
                Else
                    Record.SubItems(2) = 0
                End If
                If Len(!MBMPaidAmt) Then
                    Record.SubItems(3) = Format(!MBMPaidAmt, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                       
                Current = Current + 1
                lblCurrent = Current
                If Len(!MBMPaidAmt) Then
                    TotalAmt = TotalAmt + !MBMPaidAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdExport.Enabled = True
    CmdSearch.Enabled = True
    
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub PaymentMBMTrackingListing3(Optional strAppend As String)
Const METHOD_NAME = "PaymentMBMTrackingListing3"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    total = 0
    Current = 0
    
    lstPaymentTracking.ListItems.Clear
    
    strAppendCase = "12"
    
    sql = sql + strAppend
            
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdExport.Enabled = True
    Else
        Do Until rst2.EOF
            
            With rst2
                If Len(!CASNumber) Then
                    Set Record = lstPaymentTracking.ListItems.Add(, , !CASNumber)
                Else
                    Set Record = lstPaymentTracking.ListItems.Add(, , "Empty")
                End If
                If Len(!MBMPaidAmt) Then
                    Record.SubItems(1) = Format(!MBMPaidAmt, "#,##0.00")
                Else
                    Record.SubItems(1) = "0.00"
                End If
                       
                Current = Current + 1
                lblCurrent = Current
                If Len(!MBMPaidAmt) Then
                    TotalAmt = TotalAmt + !MBMPaidAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdExport.Enabled = True
    CmdSearch.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
   
    
End Sub

Private Sub PaymentMBMTrackingListing4(Optional strAppend As String)
Const METHOD_NAME = "PaymentMBMTrackingListing4"
On Error GoTo debugs


Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    total = 0
    Current = 0
    
    lstPaymentTracking.ListItems.Clear
    
    strAppendCase = "13"
    
    sql = sql + strAppend
            
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdExport.Enabled = True
    Else
        Do Until rst2.EOF
            
            With rst2
                
                If Len(!ClinicCode & "-" & !CLNName) Then
                    Set Record = lstPaymentTracking.ListItems.Add(, , !ClinicCode & " " & "-" & " " & !CLNName)
                Else
                    Set Record = lstPaymentTracking.ListItems.Add(, , "Empty")
                End If
                If Len(!PVDate) Then
                    Record.SubItems(1) = Format(!PVDate, "dd/mm/yyyy")
                Else
                    Record.SubItems(1) = "Empty"
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(2) = !NoofCase
                Else
                    Record.SubItems(2) = 0
                End If
                If Len(!MBMPaidAmt) Then
                    Record.SubItems(3) = Format(!MBMPaidAmt, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                       
                Current = Current + 1
                lblCurrent = Current
                If Len(!MBMPaidAmt) Then
                    TotalAmt = TotalAmt + !MBMPaidAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdExport.Enabled = True
    CmdSearch.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub PaymentMBMTrackingListing5(Optional strAppend As String)
Const METHOD_NAME = "PaymentMBMTrackingListing5"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String
    total = 0
    Current = 0
    lstPaymentTracking.ListItems.Clear
    strAppendCase = "14"
    sql = sql + strAppend
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdExport.Enabled = True
    Else
        Do Until rst2.EOF
            
            With rst2
                
                If Len(!GroupCompany) Then
                    Set Record = lstPaymentTracking.ListItems.Add(, , !GroupCompany)
                Else
                    Set Record = lstPaymentTracking.ListItems.Add(, , "Empty")
                End If
                If Len(!PVDate) Then
                    Record.SubItems(1) = Format(!PVDate, "dd/mm/yyyy")
                Else
                    Record.SubItems(1) = "Empty"
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(2) = !NoofCase
                Else
                    Record.SubItems(2) = 0
                End If
                If Len(!MBMPaidAmt) Then
                    Record.SubItems(3) = Format(!MBMPaidAmt, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                       
                Current = Current + 1
                lblCurrent = Current
                If Len(!MBMPaidAmt) Then
                    TotalAmt = TotalAmt + !MBMPaidAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdExport.Enabled = True
    CmdSearch.Enabled = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub PaymentMBMTrackingListing6(Optional strAppend As String)
Const METHOD_NAME = "PaymentMBMTrackingListing6"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    total = 0
    Current = 0
    
    lstPaymentTracking.ListItems.Clear
    
    strAppendCase = "15"
    
    sql = sql + strAppend
            
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdExport.Enabled = True
    Else
        Do Until rst2.EOF
            
            With rst2
                
                If Len(!PVNo) Then
                    Set Record = lstPaymentTracking.ListItems.Add(, , !PVNo)
                Else
                    Set Record = lstPaymentTracking.ListItems.Add(, , "Empty")
                End If
                If Len(!PVDate) Then
                    Record.SubItems(1) = Format(!PVDate, "dd/mm/yyyy")
                Else
                    Record.SubItems(1) = "Empty"
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(2) = !NoofCase
                Else
                    Record.SubItems(2) = 0
                End If
                If Len(!MBMPaidAmt) Then
                    Record.SubItems(3) = Format(!MBMPaidAmt, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                       
                Current = Current + 1
                lblCurrent = Current
                If Len(!MBMPaidAmt) Then
                    TotalAmt = TotalAmt + !MBMPaidAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdExport.Enabled = True
    CmdSearch.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'******************* End Payment MBM Tracking Listing *******************************

'******************* Start Payment Clinic Tracking Listing *******************************
Private Sub PaymentCLNTrackingListing(Optional strAppend As String)
Const METHOD_NAME = "PaymentCLNTrackingListing"
On Error GoTo debugs


Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    total = 0
    Current = 0
    
    lstPaymentTracking.ListItems.Clear
    
    strAppendCase = "4"
    
    sql = sql + strAppend
            
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdExport.Enabled = True
    Else
        Do Until rst2.EOF
        TxtPStatus = ""
        If Trim(rst2!PVPayCLNStatus) <> "" Then
            TxtPStatus = Trim(rst2!PVPayCLNStatus)
        End If
        
            With rst2
                Set Record = lstPaymentTracking.ListItems.Add(, , !PVCLNNo)
                If Len(!PVCLNDate) Then
                   Record.SubItems(1) = Format(!PVCLNDate, "dd/mm/yyyy")
                End If
                If Len(!CLNPaidAmt) Then
                   Record.SubItems(2) = Format(!CLNPaidAmt, "#,##0.00")
                End If
                If Len(!InvoiceNo) Then
                   Record.SubItems(3) = Trim(!InvoiceNo)
                End If
                If Len(!InvoiceDate) Then
                   Record.SubItems(4) = Format(!InvoiceDate, "dd/mm/yyyy")
                End If
                If Len(!CLMPayableByMedix2C) Then
                   Record.SubItems(5) = Format(!CLMPayableByMedix2C, "#,##0.00")
                End If
                If Len(!ClinicCode & "-" & !CLNName) Then
                    Record.SubItems(6) = (!ClinicCode & " " & "-" & " " & !CLNName)
                End If
                If Len(!MBMNumber) Then
                    Record.SubItems(7) = !MBMNumber
                End If
                If Len(!CASNumber) Then
                    Record.SubItems(8) = Trim(!CASNumber)
                End If
                If Len(!Bordxrefno) Then
                    Record.SubItems(9) = !Bordxrefno
                End If
                If Len(!OPBatchNo) Then
                    Record.SubItems(10) = !OPBatchNo
                End If
                If Len(!CLNRemark) Then
                    Record.SubItems(11) = Trim(!CLNRemark)
                End If
                
                If TxtPStatus = "F" Then
                    Record.SubItems(12) = "Full"
                ElseIf TxtPStatus = "P" Then
                    Record.SubItems(12) = "Partial"
                Else
                    Record.SubItems(12) = ""
                End If
                
                If Len(!CLNAccPeriodMth & "/" & !CLNAccPeriodYr) Then
                    Record.SubItems(13) = (!CLNAccPeriodMth & "/" & !CLNAccPeriodYr)
                End If
                                                              
                       
                Current = Current + 1
                lblCurrent = Current
                If Len(!CLNPaidAmt) Then
                    TotalAmt = TotalAmt + !CLNPaidAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdExport.Enabled = True
    CmdSearch.Enabled = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
 
End Sub

Private Sub PaymentCLNTrackingListing2(Optional strAppend As String)
Const METHOD_NAME = "PaymentCLNTrackingListing2"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    total = 0
    Current = 0
    
    lstPaymentTracking.ListItems.Clear
    
    strAppendCase = "5"
    
    sql = sql + strAppend
            
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdExport.Enabled = True
    Else
        Do Until rst2.EOF
            
            With rst2
                If Len(!Bordxrefno) Then
                    Set Record = lstPaymentTracking.ListItems.Add(, , !Bordxrefno)
                Else
                    Set Record = lstPaymentTracking.ListItems.Add(, , "Empty")
                End If
                If Len(!BordxDate) Then
                    Record.SubItems(1) = Format(!BordxDate, "dd/mm/yyyy")
                Else
                    Record.SubItems(1) = "Empty"
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(2) = !NoofCase
                Else
                    Record.SubItems(2) = 0
                End If
                If Len(!CLNPaidAmt) Then
                    Record.SubItems(3) = Format(!CLNPaidAmt, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                       
                Current = Current + 1
                lblCurrent = Current
                If Len(!CLNPaidAmt) Then
                    TotalAmt = TotalAmt + !CLNPaidAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdExport.Enabled = True
    CmdSearch.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub PaymentCLNTrackingListing3(Optional strAppend As String)
Const METHOD_NAME = "PaymentCLNTrackingListing3"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String
    total = 0
    Current = 0
    lstPaymentTracking.ListItems.Clear
    strAppendCase = "6"
    sql = sql + strAppend
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdExport.Enabled = True
    Else
        Do Until rst2.EOF
            
            With rst2
                If Len(!CASNumber) Then
                    Set Record = lstPaymentTracking.ListItems.Add(, , !CASNumber)
                Else
                    Set Record = lstPaymentTracking.ListItems.Add(, , "Empty")
                End If
                If Len(!CLNPaidAmt) Then
                    Record.SubItems(1) = Format(!CLNPaidAmt, "#,##0.00")
                Else
                    Record.SubItems(1) = "0.00"
                End If
                       
                Current = Current + 1
                lblCurrent = Current
                If Len(!CLNPaidAmt) Then
                    TotalAmt = TotalAmt + !CLNPaidAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdExport.Enabled = True
    CmdSearch.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub PaymentCLNTrackingListing4(Optional strAppend As String)
Const METHOD_NAME = "PaymentCLNTrackingListing4"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    total = 0
    Current = 0
    
    lstPaymentTracking.ListItems.Clear
    
    strAppendCase = "7"
    
    sql = sql + strAppend
            
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdExport.Enabled = True
    Else
        Do Until rst2.EOF
            
            With rst2
                
                If Len(!ClinicCode & "-" & !CLNName) Then
                    Set Record = lstPaymentTracking.ListItems.Add(, , !ClinicCode & " " & "-" & " " & !CLNName)
                Else
                    Set Record = lstPaymentTracking.ListItems.Add(, , "Empty")
                End If
                If Len(!PVCLNDate) Then
                    Record.SubItems(1) = Format(!PVCLNDate, "dd/mm/yyyy")
                Else
                    Record.SubItems(1) = "Empty"
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(2) = !NoofCase
                Else
                    Record.SubItems(2) = 0
                End If
                If Len(!CLNPaidAmt) Then
                    Record.SubItems(3) = Format(!CLNPaidAmt, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                       
                Current = Current + 1
                lblCurrent = Current
                If Len(!CLNPaidAmt) Then
                    TotalAmt = TotalAmt + !CLNPaidAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdExport.Enabled = True
    CmdSearch.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub PaymentCLNTrackingListing5(Optional strAppend As String)
Const METHOD_NAME = "PaymentCLNTrackingListing5"
On Error GoTo debugs


Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    total = 0
    Current = 0
    
    lstPaymentTracking.ListItems.Clear
    
    strAppendCase = "8"
    
    sql = sql + strAppend
            
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdExport.Enabled = True
    Else
        Do Until rst2.EOF
            
            With rst2
                
                If Len(!GroupCompany) Then
                    Set Record = lstPaymentTracking.ListItems.Add(, , !GroupCompany)
                Else
                    Set Record = lstPaymentTracking.ListItems.Add(, , "Empty")
                End If
                If Len(!PVCLNDate) Then
                    Record.SubItems(1) = Format(!PVCLNDate, "dd/mm/yyyy")
                Else
                    Record.SubItems(1) = "Empty"
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(2) = !NoofCase
                Else
                    Record.SubItems(2) = 0
                End If
                If Len(!CLNPaidAmt) Then
                    Record.SubItems(3) = Format(!CLNPaidAmt, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                       
                Current = Current + 1
                lblCurrent = Current
                If Len(!CLNPaidAmt) Then
                    TotalAmt = TotalAmt + !CLNPaidAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdExport.Enabled = True
    CmdSearch.Enabled = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub PaymentCLNTrackingListing6(Optional strAppend As String)
Const METHOD_NAME = "PaymentCLNTrackingListing6"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    total = 0
    Current = 0
    
    lstPaymentTracking.ListItems.Clear
    
    strAppendCase = "9"
    
    sql = sql + strAppend
            
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "PaymentReceiptModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdExport.Enabled = True
    Else
        Do Until rst2.EOF
                
            With rst2
                
                If Len(!PVCLNNo) Then
                    Set Record = lstPaymentTracking.ListItems.Add(, , !PVCLNNo)
                Else
                    Set Record = lstPaymentTracking.ListItems.Add(, , "Empty")
                End If
                If Len(!PVCLNDate) Then
                    Record.SubItems(1) = Format(!PVCLNDate, "dd/mm/yyyy")
                Else
                    Record.SubItems(1) = "Empty"
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(2) = !NoofCase
                Else
                    Record.SubItems(2) = 0
                End If
                If Len(!CLNPaidAmt) Then
                    Record.SubItems(3) = Format(!CLNPaidAmt, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                       
                Current = Current + 1
                lblCurrent = Current
                If Len(!CLNPaidAmt) Then
                    TotalAmt = TotalAmt + !CLNPaidAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdExport.Enabled = True
    CmdSearch.Enabled = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If


End Sub
'******************* End Payment Clinic Tracking Listing *******************************



'**********************Start Sorted ListView *******************************
Private Sub lstPaymentTracking_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
    If CboRptSelection.Text = "General Listing" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtstring, m_blnDirection(1)
        Case 2
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtDatetime, m_blnDirection(2)
        Case 3
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtstring, m_blnDirection(4)
        Case 5
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtDatetime, m_blnDirection(5)
        Case 6
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtstring, m_blnDirection(7)
        Case 8
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtstring, m_blnDirection(8)
        Case 9
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtstring, m_blnDirection(9)
        Case 10
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtstring, m_blnDirection(10)
        Case 11
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtstring, m_blnDirection(11)
        Case 12
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtstring, m_blnDirection(12)
        Case 13
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtstring, m_blnDirection(13)
        Case 14
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        End Select
    
    ElseIf CboRptSelection.Text = "By Bordx" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtstring, m_blnDirection(1)
        Case 2
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtDatetime, m_blnDirection(2)
        Case 3
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        End Select
    
    ElseIf CboRptSelection.Text = "By Claim" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtstring, m_blnDirection(1)
        Case 2
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        End Select
        
    ElseIf CboRptSelection.Text = "By Clinic" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtstring, m_blnDirection(1)
        Case 2
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtDatetime, m_blnDirection(2)
        Case 3
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        End Select
        
    ElseIf CboRptSelection.Text = "By Group Company" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtstring, m_blnDirection(1)
        Case 2
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtDatetime, m_blnDirection(2)
        Case 3
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        End Select
        
    ElseIf CboRptSelection.Text = "By PV No" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtstring, m_blnDirection(1)
        Case 2
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtDatetime, m_blnDirection(2)
        Case 3
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView lstPaymentTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        End Select
    
    End If
End Sub
'**********************End Sorted ListView *******************************


'******************* Start ComboBox Change/KeyPress *******************************

Private Sub CboClinic_Change()
    If InStr(CboClinic.Text, "null") Then
       CboClinic.Enabled = False
    End If
End Sub

Private Sub CboClinic_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClinic.Text, CboClinic.SelStart) + Chr(KeyAscii) + Mid(CboClinic.Text, CboClinic.SelStart + CboClinic.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClinic.ListCount - 1
 
        If NewText = LCase(Left(CboClinic.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClinic.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClinic.Text = ValidValue
                CboClinic.SelStart = 0
                CboClinic.SelLength = Len(ValidValue)
            End If
        End If
End Sub
'******************* End ComboBox Change/KeyPress *******************************

Private Sub CboGrpCompany_Change()
If CboGrpCompany.Text = "" Then
    If InStr(CboGrpCompany.Text, "null") Then
       CboGrpCompany.Enabled = False
    End If
End If
End Sub

Private Sub cboGrpCompany_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboGrpCompany.Text, CboGrpCompany.SelStart) + Chr(KeyAscii) + Mid(CboGrpCompany.Text, CboGrpCompany.SelStart + CboGrpCompany.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboGrpCompany.ListCount - 1
 
        If NewText = LCase(Left(CboGrpCompany.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboGrpCompany.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboGrpCompany.Text = ValidValue
                CboGrpCompany.SelStart = 0
                CboGrpCompany.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub txtAmtFr_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAmtTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtPVDate1_LostFocus()
    If IsDate(txtPVDate1) Then
    Else
        If txtPVDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtPVDate2_LostFocus()
    If IsDate(txtPVDate2) Then
    Else
        If txtPVDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtInvDate1_LostFocus()
    If IsDate(txtInvDate1) Then
    Else
        If txtInvDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtInvDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtInvDate2_LostFocus()
    If IsDate(txtInvDate2) Then
    Else
        If txtInvDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtInvDate2.SetFocus
        End If
    End If
End Sub

Private Sub ListViewHeader1()

    If CboReportType.ListIndex = 0 And CboRptSelection.Text = "General Listing" Then
        lstPaymentTracking.ColumnHeaders(1).Text = "PV No"
        lstPaymentTracking.ColumnHeaders(2).Text = "PV Date"
        lstPaymentTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstPaymentTracking.ColumnHeaders(3).Text = "Amt Paid"
        lstPaymentTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstPaymentTracking.ColumnHeaders(4).Text = "Invoice No"
        lstPaymentTracking.ColumnHeaders(4).Alignment = lvwColumnLeft
        lstPaymentTracking.ColumnHeaders(5).Text = "Invoice Date"
        lstPaymentTracking.ColumnHeaders(5).Alignment = lvwColumnCenter
        lstPaymentTracking.ColumnHeaders(6).Text = "Invoice Amt"
        lstPaymentTracking.ColumnHeaders(6).Alignment = lvwColumnRight
        lstPaymentTracking.ColumnHeaders(7).Text = "Payee"
        lstPaymentTracking.ColumnHeaders(7).Alignment = lvwColumnLeft
        lstPaymentTracking.ColumnHeaders(8).Text = "Member No"
        lstPaymentTracking.ColumnHeaders(8).Alignment = lvwColumnLeft
        lstPaymentTracking.ColumnHeaders(9).Text = "Claim No"
        lstPaymentTracking.ColumnHeaders(9).Alignment = lvwColumnLeft
        lstPaymentTracking.ColumnHeaders(10).Text = "Bordx No"
        lstPaymentTracking.ColumnHeaders(10).Alignment = lvwColumnLeft
        lstPaymentTracking.ColumnHeaders(11).Text = "OP Batch No"
        lstPaymentTracking.ColumnHeaders(11).Alignment = lvwColumnLeft
        lstPaymentTracking.ColumnHeaders(12).Text = "Remarks"
        lstPaymentTracking.ColumnHeaders(12).Alignment = lvwColumnLeft
        lstPaymentTracking.ColumnHeaders(13).Text = "Pymt Status"
        lstPaymentTracking.ColumnHeaders(13).Alignment = lvwColumnCenter
        lstPaymentTracking.ColumnHeaders(14).Text = "A/C Period"
        lstPaymentTracking.ColumnHeaders(14).Alignment = lvwColumnLeft
    
    ElseIf CboReportType.ListIndex = 0 And CboRptSelection.Text = "By Bordx" Then
        
        lstPaymentTracking.ColumnHeaders(1).Text = "Bordx No"
        lstPaymentTracking.ColumnHeaders(2).Text = "Bordx Date"
        lstPaymentTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstPaymentTracking.ColumnHeaders(3).Text = "No of Cases"
        lstPaymentTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstPaymentTracking.ColumnHeaders(4).Text = "Amt Paid to Claim"
        lstPaymentTracking.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 14
            lstPaymentTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf CboReportType.ListIndex = 0 And CboRptSelection.Text = "By PV No" Then
        
        lstPaymentTracking.ColumnHeaders(1).Text = "PV No"
        lstPaymentTracking.ColumnHeaders(2).Text = "PV Date"
        lstPaymentTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstPaymentTracking.ColumnHeaders(3).Text = "No of Cases"
        lstPaymentTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstPaymentTracking.ColumnHeaders(4).Text = "Amt Paid to Clinic"
        lstPaymentTracking.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 14
            lstPaymentTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf CboReportType.ListIndex = 0 And CboRptSelection.Text = "By Clinic" Then
    
        lstPaymentTracking.ColumnHeaders(1).Text = "Clinic"
        lstPaymentTracking.ColumnHeaders(2).Text = "PV Date"
        lstPaymentTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstPaymentTracking.ColumnHeaders(3).Text = "No of Cases"
        lstPaymentTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstPaymentTracking.ColumnHeaders(4).Text = "Amt Paid to Clinic"
        lstPaymentTracking.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 14
            lstPaymentTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf CboReportType.ListIndex = 0 And CboRptSelection.Text = "By Group Company" Then
    
        lstPaymentTracking.ColumnHeaders(1).Text = "Group Company"
        lstPaymentTracking.ColumnHeaders(2).Text = "PV Date"
        lstPaymentTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstPaymentTracking.ColumnHeaders(3).Text = "No of Cases"
        lstPaymentTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstPaymentTracking.ColumnHeaders(4).Text = "Amt Paid to Clinic"
        lstPaymentTracking.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 14
            lstPaymentTracking.ColumnHeaders(AA).Text = ""
        Next AA
            
    ElseIf CboReportType.ListIndex = 0 And CboRptSelection.Text = "By Claim" Then
    
        lstPaymentTracking.ColumnHeaders(1).Text = "Claim No"
        lstPaymentTracking.ColumnHeaders(2).Text = "Amt Paid to Clinic"
        lstPaymentTracking.ColumnHeaders(2).Alignment = lvwColumnRight
        
        For AA = 3 To 14
            lstPaymentTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
    End If
    
    txtPVNoFR = ""
    txtPVNoTo = ""
    txtPVDate1 = "__/__/____"
    txtPVDate2 = "__/__/____"
    txtAmtFr = ""
    TxtAmtTo = ""
    CboClinic.Text = ""
    TxtMBMNo.Text = ""
    'txtMBMName.Text = ""
    'TxtCreditFr.Text = ""
    'TxtCreditTo.Text = ""
    TxtClaimNo1.Text = ""
    TxtClaimNo2.Text = ""
    cboACPeriodMthFr.Text = ""
    cboACPeriodMthTo.Text = ""
    cboACPeriodYRFr.Text = ""
    cboACPeriodYRTo.Text = ""
    lstPaymentTracking.ListItems.Clear
    
End Sub

Private Sub ListViewHeader2()

    If CboReportType.ListIndex = 1 And CboRptSelection.Text = "General Listing" Then
        lstPaymentTracking.ColumnHeaders(1).Text = "PV No"
        lstPaymentTracking.ColumnHeaders(2).Text = "PV Date"
        lstPaymentTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstPaymentTracking.ColumnHeaders(3).Text = "Amt Paid"
        lstPaymentTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstPaymentTracking.ColumnHeaders(4).Text = "Invoice No"
        lstPaymentTracking.ColumnHeaders(4).Alignment = lvwColumnLeft
        lstPaymentTracking.ColumnHeaders(5).Text = "Invoice Date"
        lstPaymentTracking.ColumnHeaders(5).Alignment = lvwColumnCenter
        lstPaymentTracking.ColumnHeaders(6).Text = "Invoice Amt"
        lstPaymentTracking.ColumnHeaders(6).Alignment = lvwColumnRight
        lstPaymentTracking.ColumnHeaders(7).Text = "Payee"
        lstPaymentTracking.ColumnHeaders(7).Alignment = lvwColumnLeft
        lstPaymentTracking.ColumnHeaders(8).Text = "Member No"
        lstPaymentTracking.ColumnHeaders(8).Alignment = lvwColumnLeft
        lstPaymentTracking.ColumnHeaders(9).Text = "Claim No"
        lstPaymentTracking.ColumnHeaders(9).Alignment = lvwColumnLeft
        lstPaymentTracking.ColumnHeaders(10).Text = "Bordx No"
        lstPaymentTracking.ColumnHeaders(10).Alignment = lvwColumnLeft
        lstPaymentTracking.ColumnHeaders(11).Text = "OP Batch No"
        lstPaymentTracking.ColumnHeaders(11).Alignment = lvwColumnLeft
        lstPaymentTracking.ColumnHeaders(12).Text = "Remarks"
        lstPaymentTracking.ColumnHeaders(12).Alignment = lvwColumnLeft
        lstPaymentTracking.ColumnHeaders(13).Text = "Pymt Status"
        lstPaymentTracking.ColumnHeaders(13).Alignment = lvwColumnCenter
        lstPaymentTracking.ColumnHeaders(14).Text = "A/C Period"
        lstPaymentTracking.ColumnHeaders(14).Alignment = lvwColumnLeft
    
    ElseIf CboReportType.ListIndex = 1 And CboRptSelection.Text = "By Bordx" Then
        
        lstPaymentTracking.ColumnHeaders(1).Text = "Bordx No"
        lstPaymentTracking.ColumnHeaders(2).Text = "Bordx Date"
        lstPaymentTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstPaymentTracking.ColumnHeaders(3).Text = "No of Cases"
        lstPaymentTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstPaymentTracking.ColumnHeaders(4).Text = "Amt Paid to MBM"
        lstPaymentTracking.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 14
            lstPaymentTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf CboReportType.ListIndex = 1 And CboRptSelection.Text = "By PV No" Then
        
        lstPaymentTracking.ColumnHeaders(1).Text = "PV No"
        lstPaymentTracking.ColumnHeaders(2).Text = "PV Date"
        lstPaymentTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstPaymentTracking.ColumnHeaders(3).Text = "No of Cases"
        lstPaymentTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstPaymentTracking.ColumnHeaders(4).Text = "Amt Paid to MBM"
        lstPaymentTracking.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 14
            lstPaymentTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf CboReportType.ListIndex = 1 And CboRptSelection.Text = "By Clinic" Then
    
        lstPaymentTracking.ColumnHeaders(1).Text = "Clinic"
        lstPaymentTracking.ColumnHeaders(2).Text = "PV Date"
        lstPaymentTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstPaymentTracking.ColumnHeaders(3).Text = "No of Cases"
        lstPaymentTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstPaymentTracking.ColumnHeaders(4).Text = "Amt Paid to MBM"
        lstPaymentTracking.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 14
            lstPaymentTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf CboReportType.ListIndex = 1 And CboRptSelection.Text = "By Group Company" Then
    
        lstPaymentTracking.ColumnHeaders(1).Text = "Group Company"
        lstPaymentTracking.ColumnHeaders(2).Text = "PV Date"
        lstPaymentTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstPaymentTracking.ColumnHeaders(3).Text = "No of Cases"
        lstPaymentTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstPaymentTracking.ColumnHeaders(4).Text = "Amt Paid to MBM"
        lstPaymentTracking.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 14
            lstPaymentTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf CboReportType.ListIndex = 1 And CboRptSelection.Text = "By Claim" Then
    
        lstPaymentTracking.ColumnHeaders(1).Text = "Claim No"
        lstPaymentTracking.ColumnHeaders(2).Text = "Amt Paid to MBM"
        lstPaymentTracking.ColumnHeaders(2).Alignment = lvwColumnRight
        
        For AA = 3 To 14
            lstPaymentTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
    End If
    
    txtPVNoFR = ""
    txtPVNoTo = ""
    txtPVDate1 = "__/__/____"
    txtPVDate2 = "__/__/____"
    txtAmtFr = ""
    TxtAmtTo = ""
    CboClinic.Text = ""
    TxtMBMNo.Text = ""
    'txtMBMName.Text = ""
    'TxtCreditFr.Text = ""
    'TxtCreditTo.Text = ""
    TxtClaimNo1.Text = ""
    TxtClaimNo2.Text = ""
    cboACPeriodMthFr.Text = ""
    cboACPeriodMthTo.Text = ""
    cboACPeriodYRFr.Text = ""
    cboACPeriodYRTo.Text = ""
    lstPaymentTracking.ListItems.Clear
    
End Sub


