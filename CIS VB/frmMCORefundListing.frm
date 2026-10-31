VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMCORefundListing 
   Caption         =   "MCO Refund Listing"
   ClientHeight    =   8496
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11676
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8496
   ScaleWidth      =   11676
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   21
      Top             =   7080
      Width           =   11055
      Begin VB.TextBox Label5 
         Height          =   285
         Left            =   7320
         TabIndex        =   37
         Text            =   "0"
         Top             =   600
         Width           =   720
      End
      Begin VB.TextBox lblTotalMCORef 
         Height          =   285
         Left            =   5520
         TabIndex        =   36
         Text            =   "0"
         Top             =   480
         Width           =   1200
      End
      Begin VB.TextBox lblMCORefSupp 
         Height          =   285
         Left            =   3960
         TabIndex        =   35
         Text            =   "0"
         Top             =   480
         Width           =   1200
      End
      Begin VB.TextBox Label3 
         Height          =   285
         Left            =   3960
         TabIndex        =   34
         Text            =   "0"
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox txtMCOTotalSupplementary 
         Height          =   285
         Left            =   1200
         TabIndex        =   33
         Text            =   "0"
         Top             =   480
         Width           =   720
      End
      Begin VB.TextBox txtMCOTotalPrincipal 
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Text            =   "0"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   300
         Left            =   10200
         TabIndex        =   12
         Top             =   270
         Width           =   735
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   10320
         TabIndex        =   11
         Top             =   270
         Width           =   615
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   8520
         TabIndex        =   9
         Top             =   270
         Width           =   615
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   300
         Left            =   7080
         TabIndex        =   7
         Top             =   270
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   300
         Left            =   7800
         TabIndex        =   8
         Top             =   270
         Width           =   720
      End
      Begin VB.CommandButton cmdPrintCRNote 
         Caption         =   "&Print CR Note"
         Height          =   300
         Left            =   9120
         TabIndex        =   10
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Total MCO Ref"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   2
         Left            =   5520
         TabIndex        =   31
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label4 
         Caption         =   "Total MCO Ref (Supp)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   1
         Left            =   2040
         TabIndex        =   30
         Top             =   480
         Width           =   1920
      End
      Begin VB.Label Label200 
         Caption         =   "Total Principal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label201 
         Caption         =   "Total Supp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Total MCO Ref (Principal)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   2040
         TabIndex        =   23
         Top             =   240
         Width           =   1920
      End
      Begin VB.Label Label2 
         Caption         =   "Total MCO Refund (Min.)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   8160
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   1920
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Listing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5280
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   11055
      Begin MSComctlLib.ListView lvwMCO 
         Height          =   4890
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   10815
         _ExtentX        =   19071
         _ExtentY        =   8615
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   21
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Member No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "CovID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Policy No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Member Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "I/C No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Bordx Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Eff. Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Exp. Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Days Insured"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Cancellation Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Endt. Bordx Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Days Utilised"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Text            =   "MCO fees Charged"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Text            =   "MCO fees Utilised"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Text            =   "MCO fees Refund"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Text            =   "MCO fees Refund (Min.)"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Invoice No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   17
            Text            =   "Invoice Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Remarks"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Credit Note"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Credit Date"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame FrameMCO 
      Caption         =   "Enter Criteria For Searching"
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   11055
      Begin MSMask.MaskEdBox TxtDate1 
         Height          =   300
         Left            =   7320
         TabIndex        =   3
         Top             =   285
         Width           =   1125
         _ExtentX        =   1990
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDate2 
         Height          =   300
         Left            =   9000
         TabIndex        =   4
         Top             =   285
         Width           =   1125
         _ExtentX        =   1990
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboMCOListingVer 
         Height          =   315
         ItemData        =   "frmMCORefundListing.frx":0000
         Left            =   1680
         List            =   "frmMCORefundListing.frx":000D
         OLEDragMode     =   1  'Automatic
         TabIndex        =   0
         Top             =   285
         Width           =   4215
      End
      Begin VB.ComboBox CboPrintStatus 
         Height          =   315
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   560
         Width           =   1150
      End
      Begin VB.ComboBox CboDepartment 
         Height          =   315
         Left            =   7320
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox CboDateSelection 
         Height          =   315
         ItemData        =   "frmMCORefundListing.frx":0066
         Left            =   1680
         List            =   "frmMCORefundListing.frx":0070
         TabIndex        =   29
         Top             =   570
         Width           =   4215
      End
      Begin VB.ComboBox cboPayor 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   850
         Width           =   4215
      End
      Begin VB.ComboBox cboMCOGrpCom 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   850
         Width           =   4215
      End
      Begin VB.Label Label7 
         Caption         =   "Department"
         Height          =   255
         Left            =   6000
         TabIndex        =   28
         Top             =   1020
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl22 
         Caption         =   "Listing Versions"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lbl20 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1035
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Date Selection"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Date"
         Height          =   255
         Left            =   6000
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   8640
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Print Status"
         Height          =   255
         Left            =   6000
         TabIndex        =   14
         Top             =   675
         Width           =   1095
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Refund Fee Listing"
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
      TabIndex        =   26
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmMCORefundListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim MBMNo As String
Dim CurrentS As String
Private m_blnDirection(1 To 19) As Boolean
Dim printCRNote As Boolean
Dim MBMRefundFees As Variant
Dim strAppend As String
Dim strAppend2 As String

Private Sub CboDateSelection_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboDateSelection.Text, CboDateSelection.SelStart) + Chr(KeyAscii) + Mid(CboDateSelection.Text, CboDateSelection.SelStart + CboDateSelection.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboDateSelection.ListCount - 1
 
        If NewText = LCase(Left(CboDateSelection.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboDateSelection.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboDateSelection.Text = ValidValue
                CboDateSelection.SelStart = 0
                CboDateSelection.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboMCOGrpCom_Click()
Dim sql As String
Dim GRPCOMName As String
Dim rst2 As New ADODB.Recordset

    If cboMCOListingVer.ListIndex = 2 Then
        sql = ""
        GRPCOMName = ""
        GRPCOMName = RemStr(Trim(cboMCOGrpCom))
        '=Run Sp=================================================
        sql = "Exec SearchDepartment '" & Trim(GRPCOMName) & "'"
        rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        '===================================================
        CboDepartment.Clear
        Do Until rst2.EOF
            With rst2
                CboDepartment.AddItem Trim(!MBMDepartment)
                rst2.MoveNext
            End With
        Loop
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub

Private Sub cboMCOListingVer_Click()
Select Case cboMCOListingVer.ListIndex
    Case 0
        lbl20.Caption = "Payor"
        cboPayor.Visible = True
        cboMCOGrpCom.Visible = False
        Label7.Visible = False
        CboDepartment.Visible = False
        CboDateSelection.ListIndex = 0
    Case 1
        lbl20.Caption = "Group Company"
        cboMCOGrpCom.Visible = True
        cboPayor.Visible = False
        Label7.Visible = False
        CboDepartment.Visible = False
        CboDateSelection.ListIndex = 0
    Case 2
        lbl20.Caption = "Group Company"
        cboMCOGrpCom.Visible = True
        cboPayor.Visible = False
        Label7.Visible = True
        CboDepartment.Visible = True
        CboDateSelection.ListIndex = 0
    End Select

End Sub

Private Sub cboMCOListingVer_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboMCOListingVer.Text, cboMCOListingVer.SelStart) + Chr(KeyAscii) + Mid(cboMCOListingVer.Text, cboMCOListingVer.SelStart + cboMCOListingVer.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboMCOListingVer.ListCount - 1
        
        If NewText = LCase(Left(cboMCOListingVer.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboMCOListingVer.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboMCOListingVer.Text = ValidValue
               cboMCOListingVer.SelStart = 0
               cboMCOListingVer.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPayor_Change()
    If InStr(cboPayor.Text, "null") Then
       cboPayor.Enabled = False
    End If
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPayor.Text, cboPayor.SelStart) + Chr(KeyAscii) + Mid(cboPayor.Text, cboPayor.SelStart + cboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPayor.ListCount - 1
 
        If NewText = LCase(Left(cboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPayor.Text = ValidValue
                cboPayor.SelStart = 0
                cboPayor.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboPrintStatus_Change()
    If InStr(CboPrintStatus.Text, "null") Then
       CboPrintStatus.Enabled = False
    End If
End Sub

Private Sub CboPrintStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    'NewText = LCase(Left(CboPrintStatus.Text, CboPrintStatus.SelStart) + Chr(KeyAscii) + Mid(CboPrintStatus.Text, CboPrintStatus.SelStart + CboPrintStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPrintStatus.ListCount - 1
 
        If NewText = LCase(Left(CboPrintStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPrintStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPrintStatus.Text = ValidValue
                CboPrintStatus.SelStart = 0
                CboPrintStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CmdClear_Click()
    
    MCOFee = 0
    TotalAmt = 0
    TotalAmt2 = 0
    Current = 0
    CurrentS = 0
    MBMNo = ""
    MBMNumber = ""
    lvwMCO.ListItems.Clear
    txtMCOTotalPrincipal = 0
    txtMCOTotalSupplementary = 0
    CboPrintStatus = "N - No"
    Label3 = 0
    Label5 = 0
    TxtDate1.Text = "__/__/____"
    TxtDate2.Text = "__/__/____"
    cboPayor.Text = ""
    CboDateSelection.Text = ""
    CboDepartment.Text = ""
    cboMCOGrpCom.Text = ""
End Sub

Private Sub CmdPrint_Click()
    'Call ExportToExcelMBMMCORefund(lvwMCO)
    Screen.MousePointer = vbHourglass
    
    If lvwMCO.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lvwMCO)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = Default
End Sub

Private Sub cmdPrintCRNote_Click()
Const METHOD_NAME = "cmdPrintCRNote_Click"
On Error GoTo debugs


Dim sql As String
Dim PAY As New ADODB.Recordset
Dim sql2 As String
Dim rst2 As New ADODB.Recordset
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim II As Integer
Dim CRNote As String
Dim MBMNumber As String
Dim CRNoteDate As Date
Dim PAYCompanyName As String
Dim PAYAddress1 As String
Dim PAYAddress2 As String
Dim PAYAddress3 As String
Dim PAYMAINContact As String
Dim PAYTELno As String
Dim GRPCOMName As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    Screen.MousePointer = vbHourglass
    
    GRPCOMName = Replace(cboMCOGrpCom, "'", "''")
    
    If cboMCOListingVer = "" Then
        MsgBox "Please Enter Listing Version!", vbOKOnly + vbCritical, DialogBoxTitle
        cboMCOListingVer.SetFocus
    ElseIf lvwMCO.ListItems.Count = 0 Then
        MsgBox "Please Select Records!", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf CboDateSelection.Text = "" Then
        MsgBox "Please Select Date Selection!", vbOKOnly + vbCritical, DialogBoxTitle
        CboDateSelection.SetFocus
    ElseIf cboMCOListingVer.ListIndex = 0 And cboPayor.Text = "" Then
        MsgBox "Please Select Payor!", vbOKOnly + vbCritical, DialogBoxTitle
        cboPayor.SetFocus
    ElseIf cboMCOListingVer.ListIndex = 1 And cboMCOGrpCom.Text = "" Then
        MsgBox "Please Select Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
        cboMCOGrpCom.SetFocus
    ElseIf cboMCOListingVer.ListIndex = 2 And cboMCOGrpCom.Text = "" Then
        MsgBox "Please Select Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
        cboMCOGrpCom.SetFocus
    ElseIf cboMCOListingVer.ListIndex = 3 And CboDepartment.Text = "" Then
        MsgBox "Please Select Department!", vbOKOnly + vbCritical, DialogBoxTitle
        CboDepartment.SetFocus
    Else
    
        If cboMCOListingVer.ListIndex = 0 Then
        
            sql = " Select PAYCode, PAYCompanyName, PAYAddress1, " & _
                  " PAYAddress2, PAYAddress3, PAYMAINContact, PAYTELno " & _
                  " FROM PAY where PAYCode = '" & Mid(cboPayor, 1, 2) & "'"
            PAY.Open sql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
            If PAY.EOF = False Then '.RecordCount Then
                PAYCompanyName = PAY!PAYCompanyName
                PAYAddress1 = IIf(Len(PAY!PAYAddress1), PAY!PAYAddress1, "")
                PAYAddress2 = IIf(Len(PAY!PAYAddress2), PAY!PAYAddress2, "")
                PAYAddress3 = IIf(Len(PAY!PAYAddress3), PAY!PAYAddress3, "")
                PAYTELno = IIf(Len(PAY!PAYTELno), PAY!PAYTELno, "")
                PAYMAINContact = IIf(Len(PAY!PAYMAINContact), PAY!PAYMAINContact, "")
            End If
            PAY.Close
            Set PAY = Nothing
            
        Else
        
            sql = " Select GRPCoName, GRPCoAdd1, " & _
                  " GRPCoAdd2, GRPCoAdd3, GRPCoContact, GRPCoPhone " & _
                  " FROM GRPCompany where GRPCoName = '" & Trim(GRPCOMName) & "'"
            PAY.Open sql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            'PAY.Open sql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
            If PAY.EOF = False Then '.RecordCount Then
                PAYCompanyName = PAY!GrpCoName
                PAYAddress1 = IIf(Len(PAY!GRPCoAdd1), PAY!GRPCoAdd1, "")
                PAYAddress2 = IIf(Len(PAY!GRPCoAdd2), PAY!GRPCoAdd2, "")
                PAYAddress3 = IIf(Len(PAY!GRPCoAdd3), PAY!GRPCoAdd3, "")
                If Len(PAY!GRPCoPhone) Then
                    PAYTELno = PAY!GRPCoPhone
                End If
                If Len(PAY!GRPCoContact) Then
                    PAYMAINContact = PAY!GRPCoContact
                End If
            End If
            PAY.Close
            Set PAY = Nothing
            
        End If
    
    If Mid(CboPrintStatus, 1, 1) = "N" Then
        CRNote = GenerateCLNMCOCreditNo
    Else
        'sql2 = " Select MBMCLNCRNote, MBMCLNCRNoteDate " & _
               " FROM VIEW_CLN_MBM_Refund_Search Where 0 = 0 "
               
        strAppendCase = "12"
            
        sql2 = sql2 + strAppend
        
        Set cmd.ActiveConnection = MediConnCISDB
        cmd.CommandText = "MCOFeeModule"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql2)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
        'If Len(Trim(strappend)) Then
            'sql2 = sql2 + strappend
        'End If
       
        'rst2.Open sql2, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        If rst2.EOF = False Then
            CRNote = rst2!MBMCLNCRNote
            CRNoteDate = rst2!MBMCLNCRNoteDate
        End If
        rst2.Close
        Set rst2 = Nothing
    End If
    
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
        End If
        
        'If ERR.Number > 0 Then
        '    MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        '    GoTo ExitFunction
        'End If
        
        On Error GoTo HANDLE_ERROR
        ' Don't allow user to affect workbook
        objExcel.Interactive = False
    
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
        
            objExcel.WindowState = xlMaximized
    
        
        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MCO\Clinical_MCOCreditNote.xlt")
        'Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Tracking\MCOCreditNote.xlt")
        Set objWorksheet = objWorkbook.Sheets(1)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
            
            objWorksheet.Cells(8, "C") = PAYCompanyName
            objWorksheet.Cells(9, "C") = PAYAddress1
            objWorksheet.Cells(10, "C") = PAYAddress2
            objWorksheet.Cells(11, "C") = PAYAddress3
            objWorksheet.Cells(12, "C") = PAYTELno
            objWorksheet.Cells(12, "E") = PAYMAINContact
            
            If cboMCOListingVer.ListIndex = 2 Then
                objWorksheet.Cells(8, "A") = "Dept."
                objWorksheet.Cells(8, "C") = Trim(CboDepartment)
            End If
            
            objWorksheet.Cells(8, "J") = CRNote
            If Mid(CboPrintStatus, 1, 1) = "N" Then
                objWorksheet.Cells(9, "J") = Date
            Else
                objWorksheet.Cells(9, "J") = CRNoteDate
            End If
            objWorksheet.Cells(17, "H") = MBMRefundFees
            objWorksheet.Cells(17, "J") = MBMRefundFees
            
            If Mid(CboPrintStatus, 1, 1) = "N" Then
                
                For II = 1 To lvwMCO.ListItems.Count
                    
                    MBMNumber = Trim(lvwMCO.ListItems(II).Text)
                    Call UpdateCreditNote(CRNote, MBMNumber)
                        
                Next II
                
            End If
            
            objExcel.DisplayAlerts = True
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
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            GoTo ExitFunction
    'End If
    
    End If

    Screen.MousePointer = vbDefault
    cmdClear.Enabled = True
    cmdExit.Enabled = True
    cmdPrint.Enabled = True
    cmdPrintCRNote.Enabled = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdSearch_Click()
Const METHOD_NAME = "cmdSearch_Click"
On Error GoTo debugs
    
    cmdStop.Enabled = True
    cmdClear.Enabled = False
    cmdExit.Enabled = False
    cmdPrint.Enabled = False
    cmdPrintCRNote.Enabled = False
    txtMCOTotalPrincipal = 0
    txtMCOTotalSupplementary = 0
    lblMCORefSupp = 0
    lblTotalMCORef = 0
    CurrentS = "0"
    printCRNote = False
    Screen.MousePointer = vbHourglass

    If cboMCOListingVer = "" Then
        MsgBox "Please Enter Listing Version!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        If cboMCOListingVer.ListIndex = 0 Then
            If CboDateSelection.ListIndex = 0 Then
                If TxtDate2 = "__/__/____" Or TxtDate2 = "__/__/____" Then
                    MsgBox "Please Enter Endt. Eff Date!", vbOKOnly + vbCritical, DialogBoxTitle
                ElseIf cboPayor = "" Then
                    MsgBox "Please Select a Payor!", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    Call SearchRecord
                End If
            ElseIf CboDateSelection.ListIndex = 1 Then
                If TxtDate2 = "__/__/____" Or TxtDate2 = "__/__/____" Then
                    MsgBox "Please Enter Endt. Bordx Date!", vbOKOnly + vbCritical, DialogBoxTitle
                ElseIf cboPayor = "" Then
                    MsgBox "Please Select a Payor!", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    Call SearchRecord
                End If
            End If
        ElseIf cboMCOListingVer.ListIndex = 1 Then
            If CboDateSelection.ListIndex = 0 Then
                If TxtDate2 = "__/__/____" Or TxtDate2 = "__/__/____" Then
                    MsgBox "Please Enter Endt. Eff Date!", vbOKOnly + vbCritical, DialogBoxTitle
                ElseIf cboMCOGrpCom = "" Then
                    MsgBox "Please Select a Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    Call SearchRecord
                End If
            ElseIf CboDateSelection.ListIndex = 1 Then
                If TxtDate2 = "__/__/____" Or TxtDate2 = "__/__/____" Then
                    MsgBox "Please Enter Endt. Bordx Date!", vbOKOnly + vbCritical, DialogBoxTitle
                ElseIf cboMCOGrpCom = "" Then
                    MsgBox "Please Select a Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    Call SearchRecord
                End If
            End If
        ElseIf cboMCOListingVer.ListIndex = 2 Then
            If CboDateSelection.ListIndex = 0 Then
                If TxtDate2 = "__/__/____" Or TxtDate2 = "__/__/____" Then
                    MsgBox "Please Enter Endt. Eff Date!", vbOKOnly + vbCritical, DialogBoxTitle
                ElseIf cboMCOGrpCom = "" Then
                    MsgBox "Please Select a Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
                ElseIf CboDepartment = "" Then
                    MsgBox "Please Select a Department!", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    Call SearchRecord
                End If
            ElseIf CboDateSelection.ListIndex = 1 Then
                If TxtDate2 = "__/__/____" Or TxtDate2 = "__/__/____" Then
                    MsgBox "Please Enter Endt. Bordx Date!", vbOKOnly + vbCritical, DialogBoxTitle
                ElseIf cboMCOGrpCom = "" Then
                    MsgBox "Please Select a Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
                ElseIf CboDepartment = "" Then
                    MsgBox "Please Select a Department!", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    Call SearchRecord
                End If
            End If
        End If
    End If
        
    Screen.MousePointer = vbDefault
    
    cmdClear.Enabled = True
    cmdExit.Enabled = True
    cmdPrint.Enabled = True
    cmdPrintCRNote.Enabled = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
    
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    cmdExit.Enabled = True
    cmdClear.Enabled = True
    cmdPrint.Enabled = True
    cmdPrintCRNote.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CboDateSelection_Click()
    
    Select Case CboDateSelection.ListIndex
    Case 0
        Label8.Caption = "Endt. Eff Date"
        lvwMCO.ListItems.Clear
        txtMCOTotalSupplementary = 0
        txtMCOTotalPrincipal = 0
    Case 1
        Label8.Caption = "Endt. Bordx Date"
        lvwMCO.ListItems.Clear
        txtMCOTotalSupplementary = 0
        txtMCOTotalPrincipal = 0
    End Select

End Sub

Sub ComboListing()
Const METHOD_NAME = "ComboListing"
On Error GoTo debugs

Dim PAY As New ADODB.Recordset
Dim sqlstr As String
Dim GrpCom As New ADODB.Recordset
Dim sqlstr2 As String
    sqlstr = "select PAYCode, PAYCompanyName from PAY"
    PAY.Open sqlstr, MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do While Not PAY.EOF
        cboPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    sqlstr = "select GRPCoName from CISDB..GRPCompany order by GRPCoName ASC "
    GrpCom.Open sqlstr, MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do While Not GrpCom.EOF
        cboMCOGrpCom.AddItem Trim((GrpCom!GrpCoName))
        GrpCom.MoveNext
    Loop
    GrpCom.Close
    Set GrpCom = Nothing
    
    CboPrintStatus.AddItem "N - No"
    CboPrintStatus.AddItem "Y - Yes"
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
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

Private Sub Form_Load()
Call ActionLogUpdate(12, "frmMCORefundListing")
    Call ComboListing
    CboPrintStatus = "N - No"
End Sub

Private Sub SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

    strAppend = ""
    strAppend2 = ""
    If Len(Mid(cboPayor, 1, 2)) Then
        strAppend = strAppend & " AND PAYCode= '" & Mid(cboPayor, 1, 2) & "'"
    End If
    If Len(cboMCOGrpCom) Then
        strAppend = strAppend & " AND MBMGroupCompany= '" & Trim(cboMCOGrpCom) & "'"
    End If
    If Len(CboDepartment) Then
        strAppend = strAppend & " AND MBMDepartment= '" & Trim(CboDepartment) & "'"
    End If
    If CboDateSelection.ListIndex = 0 Then
        If IsDate(TxtDate1) = True Then
            strAppend = strAppend + " And MBMCLNEndtEffDate >= '" & Format(TxtDate1, "mm/dd/yyyy") & "'"
        End If
        If IsDate(TxtDate2) = True Then
            strAppend = strAppend + " And MBMCLNEndtEffDate <= '" & Format(TxtDate2, "mm/dd/yyyy") & "'"
        End If
    ElseIf CboDateSelection.ListIndex = 1 Then
        If IsDate(TxtDate1) = True Then
            strAppend = strAppend + " And MBMCLNEndtBordxDate >= '" & Format(TxtDate1, "mm/dd/yyyy") & "'"
        End If
        If IsDate(TxtDate2) = True Then
            strAppend = strAppend + " And MBMCLNEndtBordxDate <= '" & Format(TxtDate2, "mm/dd/yyyy") & "'"
        End If
    End If
    Call MCORefundListing(strAppend)
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub
Private Sub MCORefundListing(Optional strAppend As String)
Const METHOD_NAME = "MCORefundListing"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql1 As String
Dim TotalAmt As String
Dim Current As String
Dim TotalAmt2 As Variant
Dim MCOFee As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    MCOFee = 0
    TotalAmt = 0
    TotalAmt2 = 0
    Current = 0
    CurrentS = 0
    MBMNo = ""
    MBMRefundFees = 0
    lvwMCO.ListItems.Clear
    strAppendCase = "10"
    If Mid(CboPrintStatus, 1, 1) = "N" Then
        strAppend2 = strAppend2 & " AND (MBMCLNCRNote IS NULL OR MBMCLNCRNote = '') "
    Else
        strAppend2 = strAppend2 & " AND MBMCLNCRNote <> '' "
    End If
    sql = strAppend + strAppend2
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "MCOFeeModule"
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
        cmdClear.Enabled = True
        cmdExit.Enabled = True
        cmdPrint.Enabled = True
        cmdPrintCRNote.Enabled = True
    Else
    Do
        Do Until rst2.EOF
         
            With rst2
                Set Record = lvwMCO.ListItems.Add(, , !MBMNumber)
                    MBMNo = Trim(!MBMNumber)
                
                    If Len(!MBMCOVID) Then
                        Record.SubItems(1) = !MBMCOVID
                    End If
                    If !MBMCLNStatus = "C" Then
                        Record.SubItems(3) = Trim(!MBMName) & " (Cancelled)"
                    ElseIf !MBMCLNStatus = "P" Then
                        Record.SubItems(3) = Trim(!MBMName) & " (Pending)"
                    End If
                    If Len(!MBMIcBcPp) Then
                        Record.SubItems(4) = Trim(!MBMIcBcPp)
                    End If
                    If Len(Trim(!MBMBordxDate)) Then
                        Record.SubItems(5) = Format(!MBMBordxDate, "dd/mm/yyyy")
                    End If
                    If Len(!MBMPolEffDate) Then
                        Record.SubItems(6) = Format(!MBMPolEffDate, "dd/mm/yyyy")
                    End If
                    If Len(!MBMPolEXPDate) Then
                        Record.SubItems(7) = Format(!MBMPolEXPDate, "dd/mm/yyyy")
                    End If
                    If Len(!DayInsured) Then
                        Record.SubItems(8) = !DayInsured
                    End If
                    If Len(!MBMCLNEndtEffDate) Then
                        Record.SubItems(9) = Format(!MBMCLNEndtEffDate, "dd/mm/yyyy")
                    End If
                    If Len(!MBMCLNEndtBordxDate) Then
                        Record.SubItems(10) = Format(!MBMCLNEndtBordxDate, "dd/mm/yyyy")
                    End If
                    If Len(!DayUtilised) Then
                        Record.SubItems(11) = !DayUtilised
                    End If
                    If Len(!MBMCLNMCOFee) Then
                        Record.SubItems(12) = Format(!MBMCLNMCOFee, "#,##0.00")
                    End If
                    If Len(!MBMMCOUtilised) Then
                        Record.SubItems(13) = Format(!MBMMCOUtilised, "#,##0.00")
                    End If
                    If Len(!MBMMCORefund) Then
                        Record.SubItems(14) = Format(!MBMMCORefund, "#,##0.00")
                        MBMRefundFees = CDbl(MBMRefundFees) + CDbl(!MBMMCORefund)
                    End If
                    If Trim(!INSType) = "F" Or Trim(!INSType) = "H" Then
                        If Len(!MBMMCOUtilised) Then
                        If Format(!MBMMCOUtilised, "#,##0.00") <= 50# Then
                            MCOFee = (!MBMCLNMCOFee - 50)
                            Record.SubItems(15) = Format(MCOFee, "#,##0.00")
                        Else
                            Record.SubItems(15) = Format(!MBMMCOUtilised, "#,##0.00")
                        End If
                    Else
                        Record.SubItems(15) = "0.00"
                    End If
                    Else
                        If Len(!MBMMCOUtilised) Then
                            If Format(!MBMMCOUtilised, "#,##0.00") <= 25# Then
                                MCOFee = (!MBMCLNMCOFee - 25)
                                Record.SubItems(15) = Format(MCOFee, "#,##0.00")
                            Else
                                Record.SubItems(15) = Format(!MBMMCOUtilised, "#,##0.00")
                            End If
                        Else
                            Record.SubItems(15) = "0.00"
                        End If
                    
                    End If
                    
                    If Len(!MBMCInvoiceNo) Then
                        Record.SubItems(16) = Trim(!MBMCInvoiceNo)
                    End If
                    
                    If Len(!MBMCInvoiceDate) Then
                        Record.SubItems(17) = Format(!MBMCInvoiceDate, "dd/mm/yyyy")
                    End If
                    
                    If !TotalCase > 0 Then
                        Record.SubItems(18) = "This member had" & " " & !TotalCase & " " & "Case(s)"
                    End If
                    
                    If Trim(!INSType) = "F" Or Trim(!INSType) = "H" Then
                        Call MCORefundListingSupp(strAppend)
                    End If
                    
                    If Len(!MBMCLNCRNote) Then
                        Record.SubItems(19) = Format(!MBMCLNCRNote)
                    End If
                    
                    If Len(!MBMCLNCRNoteDate) Then
                        Record.SubItems(20) = Format(!MBMCLNCRNoteDate, "dd/mm/yyyy")
                    End If
                    
                    Current = Current + 1
                    txtMCOTotalPrincipal = Current
                    If Len(rst2!MBMMCORefund) Then
                        TotalAmt = TotalAmt + rst2!MBMMCORefund
                    End If
                    TotalAmt2 = TotalAmt2 + CDbl(Record.SubItems(15))
                    
                    Label3 = Format(TotalAmt, "#,##0.00")
                    Label5 = Format(TotalAmt2, "#,##0.00")
                    
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    lblTotalMCORef = CDbl(Label3) + CDbl(lblMCORefSupp)
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    cmdClear.Enabled = True
    cmdExit.Enabled = True
    cmdPrint.Enabled = True
    cmdPrintCRNote.Enabled = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub MCORefundListingSupp(Optional strAppend As String)
Const METHOD_NAME = "MCORefundListingSupp"
On Error GoTo debugs


Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim Prm4 As New ADODB.Parameter
Dim strAppendCase As String

    strAppendCase = "11"
    
    sql = "AND MBMNumber = '" & Trim(MBMNo) & "'"
    sql = sql + strAppend
        
    Set cmd2.ActiveConnection = MediConnCISDB
    cmd2.CommandText = "MCOFeeModule"
    cmd2.CommandType = adCmdStoredProc
    cmd2.CommandTimeout = 300
    Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd2.Parameters.Append Prm3
    Set Prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd2.Parameters.Append Prm4
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd2.Execute
    Do
        Do Until rst2.EOF
         
            With rst2
                Set Record = lvwMCO.ListItems.Add(, , !MBMNumber)
                                    
                    If Len(!MBMCCoverID) Then
                        Record.SubItems(1) = !MBMCCoverID
                    End If
                    If !MBMCLNStatus = "C" Then
                        Record.SubItems(3) = Trim(!MBMCName) & " (Cancelled)"
                    ElseIf !MBMCLNStatus = "P" Then
                        Record.SubItems(3) = Trim(!MBMCName) & " (Pending)"
                    End If
                    
                    If Len(!MBMCICBCPPNo) Then
                        Record.SubItems(4) = Trim(!MBMCICBCPPNo)
                    End If
                      
                    If Len(Trim(!MBMCBordxDate)) Then
                        Record.SubItems(5) = Format(!MBMCBordxDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPolEffDate) Then
                        Record.SubItems(6) = Format(!MBMPolEffDate, "dd/mm/yyyy")
                    End If
                      
                    If Len(!MBMPolEXPDate) Then
                        Record.SubItems(7) = Format(!MBMPolEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!DayInsured) Then
                        Record.SubItems(8) = !DayInsured
                    End If
                    
                    If Len(!MBMCLNEndtEffDate) Then
                        Record.SubItems(9) = Format(!MBMCLNEndtEffDate, "dd/mm/yyyy")
                    End If
                      
                    If Len(!MBMCLNEndtBordxDate) Then
                        Record.SubItems(10) = Format(!MBMCLNEndtBordxDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!DayUtilised) Then
                        Record.SubItems(11) = !DayUtilised
                    End If
                    
                    If Len(!MBMCCLNMCO) Then
                        Record.SubItems(12) = Format(!MBMCCLNMCO, "#,##0.00")
                    End If
                    
                    If Len(!MBMMCOUtilised) Then
                        Record.SubItems(13) = Format(!MBMMCOUtilised, "#,##0.00")
                    End If
                    
                    If Len(!MBMMCORefund) Then
                        Record.SubItems(14) = Format(!MBMMCORefund, "#,##0.00")
                        MBMRefundFees = CDbl(MBMRefundFees) + CDbl(!MBMMCORefund)
                    End If
                    If Len(!MBMMCOUtilised) Then
                        If Format(!MBMMCOUtilised, "#,##0.00") <= 50# Then
                            MCOFee = (!MBMCCLNMCO - 50)
                            Record.SubItems(15) = Format(MCOFee, "#,##0.00")
                        Else
                            Record.SubItems(15) = Format(!MBMMCOUtilised, "#,##0.00")
                        End If
                    Else
                        Record.SubItems(15) = "0.00"
                    End If
                    If Len(!MBMCInvoiceNo) Then
                        Record.SubItems(16) = Trim(!MBMCInvoiceNo)
                    End If
                    
                    If Len(!MBMCInvoiceDate) Then
                        Record.SubItems(17) = Format(!MBMCInvoiceDate, "dd/mm/yyyy")
                    End If


                    If !TotalCase > 0 Then
                        Record.SubItems(18) = "This member had" & " " & !TotalCase & " " & "Case(s)"
                    End If
                    
                    If Len(!MBMCCLNCRNote) Then
                        Record.SubItems(19) = Format(!MBMCCLNCRNote)
                    End If
                    
                    If Len(!MBMCCLNCRNoteDate) Then
                        Record.SubItems(20) = Format(!MBMCCLNCRNoteDate, "dd/mm/yyyy")
                    End If
                                                    
                    CurrentS = CurrentS + 1
                    txtMCOTotalSupplementary = CurrentS
                    If IsNumeric(!MBMMCORefund) Then
                        lblMCORefSupp = CDbl(lblMCORefSupp) + CDbl(!MBMMCORefund)
                    End If
            
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    'End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub lvwMCO_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
    Select Case ColumnHeader.Index
    Case 1
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
    Case 3
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lvwMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lvwMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
    Case 8
        SortListView lvwMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lvwMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
    Case 11
        SortListView lvwMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    Case 13
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
    Case 14
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
    Case 15
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
    Case 16
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
    Case 17
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lvwMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(18)
    Case 19
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(19)
    End Select
End Sub

Private Sub TxtDate1_LostFocus()
    If IsDate(TxtDate1) Then
    Else
        If TxtDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDate2_LostFocus()
    If IsDate(TxtDate2) Then
    Else
        If TxtDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDate2.SetFocus
        End If
    End If
End Sub

Private Sub UpdateCreditNote(CRNote As String, MBMNumber As String)
Const METHOD_NAME = "UpdateCreditNote"
On Error GoTo debugs


Dim strsqlCoveredPersons As String
Dim strsqlMBMOthers As String
Dim CRNoteDate As String

    CRNoteDate = Format(Date, "YYYY/MM/DD")
    strsqlMBMOthers = " update MBMCLNOthers set " & _
                      " MBMCLNCRNoteDate = '" & CRNoteDate & "', " & _
                      " MBMCLNCRNote = '" & Trim(CRNote) & "', " & _
                      " MBMCLNCRNoteby = '" & Logon & "' " & _
                      " Where MBMNumber = '" & Trim(MBMNumber) & "'"
    medixmasterconn.Execute strsqlMBMOthers
    
    strsqlCoveredPersons = " update MBMCoveredPersonsCLN set " & _
                           " MBMCCLNCRNoteDate = '" & CRNoteDate & "', " & _
                           " MBMCCLNCRNote = '" & Trim(CRNote) & "', " & _
                           " MBMCCLNCRNoteby = '" & Logon & "' " & _
                           " Where MBMCNumber = '" & Trim(MBMNumber) & "'"
    
    medixmasterconn.Execute strsqlCoveredPersons

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub






