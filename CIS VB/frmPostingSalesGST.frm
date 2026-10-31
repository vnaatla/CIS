VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPostingSalesGST 
   Caption         =   "Posting - Sales Invoice"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   60
      TabIndex        =   34
      Top             =   7320
      Width           =   11775
      Begin VB.CommandButton CmdPost 
         Caption         =   "&Post"
         Height          =   330
         Left            =   8520
         TabIndex        =   41
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   7800
         TabIndex        =   40
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10200
         TabIndex        =   39
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   38
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9240
         TabIndex        =   37
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   330
         Left            =   8640
         TabIndex        =   36
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   330
         Left            =   7080
         TabIndex        =   35
         Top             =   200
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Credit Amt"
         Height          =   255
         Left            =   4800
         TabIndex        =   47
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblCrAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   5760
         TabIndex        =   46
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1800
         TabIndex        =   45
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblDbAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3600
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "Debit Amt"
         Height          =   255
         Left            =   2760
         TabIndex        =   43
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblCurrent 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame FramePostInfo 
      Height          =   615
      Left            =   60
      TabIndex        =   26
      Top             =   6720
      Width           =   11775
      Begin VB.Timer Timer1 
         Interval        =   8000
         Left            =   7320
         Top             =   120
      End
      Begin VB.ComboBox cboAccPeriod 
         Height          =   315
         Left            =   1080
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "A/C Period"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Post Status"
         Height          =   255
         Left            =   4800
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LblPost 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   5760
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblBatchNo 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3600
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Batch No"
         Height          =   255
         Left            =   2760
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   60
      TabIndex        =   1
      Top             =   240
      Width           =   11715
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   4095
      End
      Begin VB.ComboBox CboPayor 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   520
         Width           =   4095
      End
      Begin VB.TextBox txtInvNo1 
         Height          =   315
         Left            =   7680
         TabIndex        =   7
         Top             =   495
         Width           =   1455
      End
      Begin VB.TextBox txtInvNo2 
         Height          =   315
         Left            =   9960
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox CboRptType 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   810
         Width           =   4095
      End
      Begin VB.ComboBox cboPostStatus 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Text            =   "cboPostStatus"
         Top             =   1100
         Width           =   4095
      End
      Begin VB.TextBox TxtBatch1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7680
         TabIndex        =   3
         Top             =   1050
         Width           =   1455
      End
      Begin VB.TextBox TxtBatch2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   9960
         TabIndex        =   2
         Top             =   1050
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtInvDate1 
         Height          =   315
         Left            =   7680
         TabIndex        =   10
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDate2 
         Height          =   315
         Left            =   9960
         TabIndex        =   11
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPostDate1 
         Height          =   315
         Left            =   7680
         TabIndex        =   12
         Top             =   765
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPostDate2 
         Height          =   315
         Left            =   9960
         TabIndex        =   13
         Top             =   765
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
         Caption         =   "Report Type"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   860
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Batch No. from"
         Height          =   255
         Left            =   6240
         TabIndex        =   24
         Top             =   1100
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "To"
         Height          =   255
         Left            =   9360
         TabIndex        =   23
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label35 
         Caption         =   "Health Type"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   9360
         TabIndex        =   21
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "To"
         Height          =   255
         Left            =   9360
         TabIndex        =   20
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         Height          =   255
         Left            =   9360
         TabIndex        =   19
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label6 
         Caption         =   "Exp Date from"
         Height          =   255
         Left            =   6240
         TabIndex        =   18
         Top             =   820
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Exp Status"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label lblRptNo 
         Caption         =   "Invoice No. from"
         Height          =   255
         Left            =   6240
         TabIndex        =   16
         Top             =   530
         Width           =   1455
      End
      Begin VB.Label lblRptDate 
         Caption         =   "Invoice Date from"
         Height          =   255
         Left            =   6240
         TabIndex        =   15
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   580
         Width           =   615
      End
   End
   Begin VB.CheckBox chkCheckAll 
      Caption         =   "Check All"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1900
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstMBMInvList 
      Height          =   4575
      Left            =   60
      TabIndex        =   33
      Top             =   2160
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   19
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Chk"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "J/T"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "J/S"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "DB"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Trans.Ref."
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "AccPeriod"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "TransDate"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "AccCode"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "D/C"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Hospital"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Grp Company"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Payor"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "CaseNo"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "Prod.Cat."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Bordx Ref"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "BatchNo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Index"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Sales Invoice Posting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmPostingSalesGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intExit As Boolean
Dim SearchCri As String
Dim PostingType As String
Dim PostingStatus As Boolean
Dim tmpVarDate As String
Dim tmpVarNo As String
Dim tmpFieldName As String
Dim TmpDone As Boolean
Dim tmpEndTime As String

Private Sub ComboListing()
Dim PayorRec As New ADODB.Recordset
Dim PayorSql As String
Dim HltRec As New ADODB.Recordset
Dim HltSql As String
Dim MthStart
Dim YrStart
Dim Acc As New ADODB.Recordset
Dim Accsql As String


    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    HltSql = "Select HLTCode, HLTDescription from HLT "
    HltRec.Open HltSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    Do While Not HltRec.EOF
        cboHLTCode.AddItem Trim(HltRec!HLTCode) & " - " & Trim(HltRec!HLTDescription)
        HltRec.MoveNext
    Loop
    HltRec.Close
    Set HltRec = Nothing

    PayorSql = "Select PayCode, HLTCode, PAYCompanyName from Pay "
    PayorRec.Open PayorSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    Do While Not PayorRec.EOF
        CboPayor.AddItem Trim(PayorRec!PAYCode) & "-" & Trim(PayorRec!PAYCompanyName)
        PayorRec.MoveNext
    Loop
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    cboPostStatus.Clear
    cboPostStatus.AddItem "0 - No"
    cboPostStatus.AddItem "1 - Yes"
    cboPostStatus.ListIndex = 0
    CboRptType.Clear
    CboRptType.AddItem "Sales Invoice"
    CboRptType.AddItem "Credit Note"
    CboRptType.ListIndex = 0
    'CboAccMth.Clear
    'For MthStart = 1 To 12
    '    With CboAccMth
    '        .AddItem Format(MthStart, "00")
    '    End With
    'Next MthStart
    cboAccPeriod.Clear
    Accsql = "Select AccMonth, AccYear from AccPeriod "
    Acc.Open Accsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    Do Until Acc.EOF
        cboAccPeriod.AddItem Trim(Acc!AccMonth) & "/" & Trim(Acc!AccYear)
        Acc.MoveNext
    Loop
        Acc.Close
        Set Acc = Nothing
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
    
    'For YrStart = 1999 To 3000
    '    With CboAccYear
    '        .AddItem YrStart
    '    End With
    'Next YrStart

End Sub

Private Sub cboAccPeriod_KeyPress(KeyAscii As Integer)
    Call CboKeyPress(cboAccPeriod, KeyAscii)
End Sub

Private Sub cboAccPeriod_LostFocus()
    If IsMonth(cboAccPeriod) Then
    Else
        If cboAccPeriod <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            cboAccPeriod.SetFocus
        End If
    End If
End Sub

Private Sub CboPayor_Change()
    If CboPayor.Text = "" Then
        If InStr(CboPayor.Text, "null") Then
           CboPayor.Enabled = False
        End If
    End If
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
    Call CboKeyPress(CboPayor, KeyAscii)
End Sub

Private Sub cboPostStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPostStatus.Text, cboPostStatus.SelStart) + Chr(KeyAscii) + Mid(cboPostStatus.Text, cboPostStatus.SelStart + cboPostStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPostStatus.ListCount - 1
 
        If NewText = LCase(Left(cboPostStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPostStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPostStatus.Text = ValidValue
                cboPostStatus.SelStart = 0
                cboPostStatus.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboPostStatus_LostFocus()
    TxtBatch1.Enabled = False
    TxtBatch2.Enabled = False
    txtPostDate1.Enabled = False
    txtPostDate2.Enabled = False
    If Mid(cboPostStatus, 1, 1) = "1" Then
        TxtBatch1.Enabled = True
        TxtBatch2.Enabled = True
        txtPostDate1.Enabled = True
        txtPostDate2.Enabled = True
    End If
End Sub

Private Sub CboRptType_Click()
    If Trim(CboRptType) = "Sales Invoice" Then
        lblRptDate = "Invoice Date From"
        lblRptNo = "Invoice No. From"
        tmpVarDate = "MBMInvoiceDate"
        tmpVarNo = "MBMInvoiceNo"
        PostingType = "SalesInv"
        tmpFieldName = "SalesInvStatus"
    Else
        lblRptDate = "CR Date From"
        lblRptNo = "CR No. From"
        tmpVarDate = "MBMCRNoteDate"
        tmpVarNo = "MBMCRNote"
        PostingType = "MBMCRNote"
        tmpFieldName = "MBMCRStatus"
    End If
End Sub

Private Sub chkCheckAll_Click()
Dim LstCnt As Integer
    LstCnt = lstMBMInvList.ListItems.Count
        If LblPost = "Yes" Then
            'MsgBox "Record(s) has been posted ! ", vbOKOnly + vbCritical, DialogBoxTitle
            chkCheckAll.Value = 0
            Exit Sub
        Else
        If LstCnt = 0 Then
            chkCheckAll.Value = 0
        Else

            For LstCnt = 1 To LstCnt
                If lstMBMInvList.ListItems(LstCnt).Checked = False Then
                    lstMBMInvList.ListItems(LstCnt).Checked = True
                    lstMBMInvList.ListItems(LstCnt).Text = "Post"
                ElseIf lstMBMInvList.ListItems(LstCnt).Checked = True Then
                    lstMBMInvList.ListItems(LstCnt).Checked = False
                    lstMBMInvList.ListItems(LstCnt).Text = ""
                End If
            Next LstCnt
        End If
    End If

End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    cboPostStatus.ListIndex = 0
    CboPayor.Text = ""
    lblCurrent = ""
    LblDbAmt = ""
    LblCrAmt = ""
    lstMBMInvList.ListItems.Clear
    FramePostInfo.Enabled = False
    TxtBatch1.Enabled = False
    TxtBatch2.Enabled = False
    txtPostDate1.Enabled = False
    txtPostDate2.Enabled = False
    CboRptType.Enabled = True
    CboRptType.ListIndex = 0
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPost_Click()
On Err GoTo ErrorSave
Dim startTrans As Boolean
Dim AA As Integer
Dim StrINvSql As String
Dim TmpInvIndex As String
Dim UpdAsk
Dim BatchNo As String
Dim CompTrans As Boolean
Dim UpdateSuc As Boolean
Dim ItemChecked As Boolean


    ItemChecked = False
    CompTrans = True
    UpdateSuc = False
    For AA = 1 To lstMBMInvList.ListItems.Count
        If lstMBMInvList.ListItems(AA).Checked = True Then
            ItemChecked = True
            Exit For
        End If
    Next AA
    If ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
        Exit Sub
    ElseIf LblPost = "Yes" Then
        MsgBox "Record(s) has been posted ! ", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    ElseIf Trim(cboAccPeriod.Text) = "" Then
        MsgBox "Please Keyin for Accounting Period ! ", vbOKOnly + vbCritical, DialogBoxTitle
        cboAccPeriod.SetFocus
        Exit Sub
    Else
        UpdAsk = MsgBox("Do you want to Post the record(s) ?", vbYesNo + vbExclamation)
        If UpdAsk = vbYes Then
            Screen.MousePointer = vbHourglass
            'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            medixmasterconnWS.beginTrans
            CompTrans = False
            BatchNo = GenSUNPostingBatchNo
            For AA = 1 To lstMBMInvList.ListItems.Count
                TmpInvIndex = lstMBMInvList.ListItems(AA).SubItems(18)
                If lstMBMInvList.ListItems(AA).Checked = True And Trim(TmpInvIndex) <> "" Then
                    StrINvSql = ""
                    StrINvSql = "Insert Into SUNPosting(TransRef, " & tmpFieldName & "," & _
                                "PostStatus, PostType, PostDate, PostBy, " & _
                                "AccMonth, AccYear, BatchNo) " & _
                                "Values('" & TmpInvIndex & "','1','1','" & PostingType & _
                                "','" & Format(Date, "yyyy/mm/dd") & _
                                "','" & Logon & "','" & Mid(cboAccPeriod.Text, 1, 2) & _
                                "','" & Mid(cboAccPeriod.Text, 4, 4) & "','" & BatchNo & "')"
                    medixmasterconnWS.Execute StrINvSql
                End If
            Next AA
            medixmasterconnWS.commitTrans
            CompTrans = True
            UpdateSuc = True
            lblBatchNo = BatchNo
            Screen.MousePointer = vbDefault
            MsgBox "Record(s) Posted Sucessfully... ! ", vbOKOnly
            cboPostStatus.ListIndex = 1
ErrorSave:
            Screen.MousePointer = vbDefault
            If CompTrans = False Then
                medixmasterconnWS.RollbackTrans
            End If
           ' medixmasterconnWS.Close
           ' Set medixmasterconnWS = Nothing
            If CompTrans = False Then
                MsgBox Err.Description
            Else
                If UpdateSuc = True Then
                    StrINvSql = ""
                    StrINvSql = " and BatchNo = '" & Trim(lblBatchNo) & "'"
                    Call MBMInvListing(StrINvSql)
                    PostingStatus = True
                    Call CmdPrintFile_Click
                End If
            End If
        End If
    End If
End Sub

Private Sub CmdPrintFile_Click()
Dim SVRPath As String
Dim Filename As String

    Screen.MousePointer = vbHourglass
        
        'If PostingStatus = False Then
        '    If SearchStatus = "G" Then
        '        Call ExportListViewtoExcel(lstRecPayList)
        '    Else
        '        Call ExportListViewtoExcel(lstRecPayCDList)
        '    End If
        'Else
        '               '\\Mars\accounts\Account Journal Import\JI Not Posted
        '    SVRPath = "\\Mars\accounts\Account Journal Import\JI Not Posted\"
         '   Filename = lblBatchNo & " CR-Payor " & Format(txtReceiptDate1, "yyyymmdd") & "-" & Mid(txtReceiptDate2, 1, 2)
         '   Call ExporttoExcelPosting(lstRecPayList, SVRPath, Filename)
         '   PostingStatus = False
        'End If
        
    If lstMBMInvList.ListItems.Count <> 0 Then
        If PostingStatus = False Then
            Call ExportListViewtoExcel(lstMBMInvList)
        Else
                       '\\Mars\accounts\Account Journal Import\JI Not Posted
            SVRPath = "\\Svr-Sun3\Finance\Account Journal Import\HIS Not Posted\"
            Filename = lblBatchNo & " SI-SalesInv " & Format(txtInvDate1, "yyyymmdd") & "-" & Mid(txtInvDate2, 1, 2)
            Call ExporttoExcelPosting(lstMBMInvList, SVRPath, Filename)
            PostingStatus = False
        End If
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub CmdSearch_Click()
Dim strsql As String
Dim PostSql As String
    strsql = ""
    PostSql = ""
    Call ClearInfo
    CboRptType.Enabled = False
    If Trim(cboHLTCode) <> "" Then
        strsql = strsql & " And HltCode='" & Mid(cboHLTCode, 1, 1) & "'"
    End If
    If Trim(CboPayor) <> "" Then
        strsql = strsql & " And PayCode='" & Mid(CboPayor, 1, 2) & "'"
    End If
    If IsDate(txtInvDate1) Then
        strsql = strsql & " And " & tmpVarDate & " >='" & Format(txtInvDate1, "yyyy/mm/dd") & "'"
    End If
    If IsDate(txtInvDate2) Then
        strsql = strsql & " And " & tmpVarDate & " <='" & Format(txtInvDate2, "yyyy/mm/dd") & "'"
    End If
    If Trim(txtInvNo1) <> "" Then
        strsql = strsql & " And " & tmpVarNo & " >='" & Trim(txtInvNo1) & "'"
    End If
    If Trim(txtInvNo2) <> "" Then
        strsql = strsql & " And " & tmpVarNo & " <='" & Trim(txtInvNo2) & "'"
    End If
    If Mid(cboPostStatus, 1, 1) = "0" Then
        PostSql = " And (" & tmpFieldName & " IS NULL OR " & tmpFieldName & " = '0')"
        PostSql = PostSql & " And (PostStatus IS NULL OR PostStatus = '0')"
    End If
    If Mid(cboPostStatus, 1, 1) = "1" Then
        PostSql = " And " & tmpFieldName & " ='1' And PostStatus ='1'"
    End If
    If IsDate(txtPostDate1) Then
        strsql = strsql & " And PostDate >='" & Format(txtPostDate1, "yyyy/mm/dd") & "'"
    End If
    If IsDate(txtPostDate2) Then
        strsql = strsql & " And PostDate <='" & Format(txtPostDate2, "yyyy/mm/dd") & "'"
    End If
    If Trim(TxtBatch1) <> "" Then
        strsql = strsql & " And BatchNo >='" & Trim(TxtBatch1) & "'"
    End If
    If Trim(TxtBatch2) <> "" Then
        strsql = strsql & " And BatchNo <='" & Trim(TxtBatch2) & "'"
    End If
    SearchCri = strsql
    Screen.MousePointer = vbHourglass
    Call MBMInvListing(strsql & PostSql)
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub CmdStop_Click()
    intExit = True
    CmdExit.Enabled = True
End Sub

Private Sub Form_Load()
    intExit = False
    TmpDone = False
    Call ComboListing
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

Private Sub lstMBMInvList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim FirstChk As Integer
Dim SEQNumber As Integer
Dim SecChk As Integer

Dim CurrChk As Integer
Dim ModResult As Integer
Dim NextChk As Integer
Dim NextChk2 As Integer
Dim NextChk3 As Integer
Dim NextChk4 As Integer
Dim lstRec As ListItem
Dim PrevText As String
Dim PostStatus As String
Dim ForLoop, ListCount As Integer
Dim CHKItem As Integer
ListCount = lstMBMInvList.ListItems.Count
    If Item.Index > 1 Then
        CHKItem = Item.Index - 1
    Else
        CHKItem = 1
    End If
    NextChk2 = 0
    If LblPost = "Yes" Then
        MsgBox "Record(s) has been posted ! ", vbOKOnly + vbCritical, DialogBoxTitle
        Item.Checked = False
        
    ElseIf Item.ListSubItems(10) = "D" Then
        MsgBox "Please select 'Credit' item. ", vbOKOnly + vbCritical, DialogBoxTitle
        If Item.Text = "Post" Then
            Item.Checked = True
        Else
            Item.Checked = False
        End If
        
    ElseIf Item.Index > 1 And lstMBMInvList.ListItems(CHKItem).SubItems(10) = "C" Then
        'CHKItem = Item.Index - 1
        'If lstMBMInvList.listitems(CHKItem).SubItems(10) = "C" Then
            MsgBox "Please select first 'Credit' item. ", vbOKOnly + vbCritical, DialogBoxTitle
            If Item.Text = "Post" Then
                Item.Checked = True
            Else
                Item.Checked = False
            End If
        'End If
    'ElseIf Item.ListSubItems(10) = "D" And Mid(cboPostType, 1, 1) = "N" Then
    '    MsgBox "Please select 'Credit' item. ", vbOKOnly + vbCritical, DialogBoxTitle
    '    If Item.Text = "Post" Then
    '        Item.Checked = True
    '    Else
    '        Item.Checked = False
    '    End If
        
    Else
        
        'If Item.Index > 1 Then
        '    CHKItem = Item.Index - 1
        '    If lstMBMInvList.listitems(CHKItem).SubItems(10) = "C" Then
        '        FirstChk = Item.Index - 1
        '    End If
        'Else
            FirstChk = Item.Index
        'End If
        'Item.ListSubItems(10) = "C"
        'FirstChk = Item.Index
        'Item.Text = "Post"
        SEQNumber = FirstChk Mod 2
        If SEQNumber = 0 Then
            SecChk = FirstChk - 1
        Else
            SecChk = FirstChk + 1
        End If

        PostStatus = Item.Text
        If Item.Text = "" Then
            Item.Text = "Post"
            For ForLoop = FirstChk + 1 To ListCount
                With lstMBMInvList.ListItems(FirstChk)
                    'If Mid(cboPostType, 1, 1) = "P" Then
                        If lstMBMInvList.ListItems(ForLoop).SubItems(10) = "D" Then
                            lstMBMInvList.ListItems(ForLoop).Checked = True
                            lstMBMInvList.ListItems(ForLoop).Text = "Post"
                            Exit Sub
                        ElseIf lstMBMInvList.ListItems(ForLoop).SubItems(10) <> "D" Then
                            lstMBMInvList.ListItems(ForLoop).Checked = True
                            lstMBMInvList.ListItems(ForLoop).Text = "Post"
                        End If
                    'Else
                    '    If lstMBMInvList.listitems(ForLoop).SubItems(10) = "C" Then
                    '        Exit Sub
                    '    ElseIf lstMBMInvList.listitems(ForLoop).SubItems(10) <> "C" Then
                    '        lstMBMInvList.listitems(ForLoop).Checked = True
                    '        lstMBMInvList.listitems(ForLoop).Text = "Post"
                    '    End If
                    'End If
                End With
            Next
        Else
            Item.Text = ""
            For ForLoop = FirstChk + 1 To ListCount
                With lstMBMInvList.ListItems(FirstChk)
                    'If Mid(cboPostType, 1, 1) = "P" Then
                        If lstMBMInvList.ListItems(ForLoop).SubItems(10) = "D" Then
                            lstMBMInvList.ListItems(ForLoop).Checked = False
                            lstMBMInvList.ListItems(ForLoop).Text = ""
                            Exit Sub
                        ElseIf lstMBMInvList.ListItems(ForLoop).SubItems(10) = "C" Then
                            lstMBMInvList.ListItems(ForLoop).Checked = False
                            lstMBMInvList.ListItems(ForLoop).Text = ""
                        End If
                    'Else
                    '    If lstMBMInvList.listitems(ForLoop).SubItems(10) = "C" Then
                    '        Exit Sub
                    '    ElseIf lstMBMInvList.listitems(ForLoop).SubItems(10) <> "C" Then
                    '        lstMBMInvList.listitems(ForLoop).Checked = False
                    '        lstMBMInvList.listitems(ForLoop).Text = ""
                    '    End If
                    'End If
                End With
            Next
        End If
    End If
    'If LblPost = "Yes" Then
    '    MsgBox "Record(s) has been posted ! ", vbOKOnly + vbCritical, DialogBoxTitle
    '    Item.Checked = False
    'Else
    '    CurrChk = Item.Index
    '    ModResult = CurrChk Mod 2
    '    If ModResult = 0 Then
    '        NextChk = CurrChk - 1
    '    Else
    '        NextChk = CurrChk + 1
    '    End If
  '
  '      PrevText = Item.Text
  '      If Item.Text = "" Then
  '          Item.Text = "Post"
  '          lstMBMInvList.listitems(NextChk).Text = "Post"
  '          Item.Checked = True
  '          lstMBMInvList.listitems(NextChk).Checked = True
  '      End If
  '      If PrevText = "Post" Then
  '          Item.Text = ""
   '         lstMBMInvList.listitems(NextChk).Text = ""
  '          Item.Checked = False
  '          lstMBMInvList.listitems(NextChk).Checked = False
  '      End If
  '  End If
End Sub

Private Sub Timer1_Timer()
Dim tmpTime As String
    If SpcAccess <> "Y" Then
        If TmpDone = False Then
            If CheckTime("POSTING") = False Then Exit Sub
            TmpDone = True
        End If
        'tmpTime = GetRemoteTime
        'tmpTime = DateDiff("s", Format(Now, "hh:mm AMPM"), Format(tmpEndTime, "hh:mm AMPM"))
        If DateDiff("s", Format(Now, "hh:mm AMPM"), Format(tmpEndTime, "hh:mm AMPM")) >= 0 Then
            MsgBox "Posting" & " is ended at  " & Format(tmpEndTime, "hh:mm:ss AMPM") & " ! "
            Unload Me
        End If
    End If
End Sub

Private Sub txtPostDate1_LostFocus()
    If IsDate(txtPostDate1) Then
    Else
        If txtPostDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPostDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtPostDate2_LostFocus()
    If IsDate(txtPostDate2) Then
    Else
        If txtPostDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPostDate2.SetFocus
        End If
    End If
End Sub

Private Sub MBMInvListing(StrAppend As String)
Dim lstRec As ListItem
Dim MBMInv As New ADODB.Recordset
Dim tmpPayor, tmpHltCode, tmpINSCode, tmpPatient, tmpCoGRPCode As String
Dim JournalType, AcCode, CreditType As String
Dim AA, CurrRec As Integer
Dim TotDbAmt, TotCrAmt As Double
Dim cmd As New ADODB.Command
Dim App1 As New ADODB.Parameter
Dim App2 As New ADODB.Parameter
Dim tmpProdCat As String
Dim DispName As String
Dim SuppRec As New ADODB.Recordset
Dim SuppSql As String
Dim tmpSuppFee As Double
Dim tmpPayMCO As Double
Dim TotalPrinc As String
Dim tmpTotalPayMCO As Double
Dim tmpTotalMCO As String
Dim PSStatus As Boolean
Dim PSRepeat As Boolean
Dim strPaySearch As String
Dim PaySearch As New ADODB.Recordset
Dim TempInvNo As String
Dim InvSearch As New ADODB.Recordset
Dim strsqlInvSearch As String
Dim RecCount As Integer

    tmpPayor = ""
    tmpHltCode = ""
    lstMBMInvList.ListItems.Clear
    CurrRec = 0
    LblPost = "No"
    tmpPayMCO = 0
    TotalPrinc = 0
    tmpTotalPayMCO = 0
    TempInvNo = ""
    RecCount = 0
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    ' VIEW_SUN_Posting_MBMInv
    Set cmd.ActiveConnection = medixmasterconnWS
    cmd.CommandText = "SUN_HIS_Posting"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set App1 = cmd.CreateParameter("RptType", adVarChar, adParamInput, 20, PostingType)
    cmd.Parameters.Append App1
    Set App2 = cmd.CreateParameter("StrCriterial", adVarChar, adParamInput, 1000, StrAppend)
    cmd.Parameters.Append App2
    MBMInv.CursorLocation = adUseClient
    MBMInv.CursorType = adOpenForwardOnly
    MBMInv.CacheSize = 100
    Set MBMInv = cmd.Execute
    
    
    Do While Not MBMInv.EOF
    
        If TempInvNo <> MBMInv!refno Then
            strsqlInvSearch = " SELECT MBMInvoiceNo From VIEW_SUN_Posting_SalesInvoice_Search Where MBMInvoiceNo = '" & Trim(MBMInv!refno) & "'"
            InvSearch.Open strsqlInvSearch, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            'RecCount = InvSearch.recordCount
            RecCount = 0
            Do While Not InvSearch.EOF
                RecCount = RecCount + 1
                InvSearch.MoveNext
            Loop
            InvSearch.Close
            Set InvSearch = Nothing
        Else
            GoTo NextRecord
        End If
    
        If RecCount > 1 Then
            'strsqlIdv = strsqlIdv & " And PVNo ='" & Trim(PayHosp!PVNo) & "'"
            TempInvNo = MBMInv!refno
            Call MCOIndvListing(TempInvNo)
            CurrRec = CurrRec + 1
        Else
        
        CurrRec = CurrRec + 1
        With MBMInv
            PSStatus = False
            If MBMInv!SalesStatus = "1" Then
                LblPost = "Yes"
            End If
            tmpPayor = IIf(Len(!PAYCode), !PAYCode, "")
            tmpHltCode = IIf(Len(!HLTCode), !HLTCode, "")
            'tmpCoGRPCode = IIf(Len(!CoGRPCode), !CoGRPCode, "")
            tmpProdCat = "SM"
            If tmpHltCode <> "S" Then
                tmpProdCat = GetProductCat(tmpHltCode, tmpPayor)
            End If
            DispName = tmpPayor & " " & IIf(IsDate(!MBMBordxDate), Format(!MBMBordxDate, "DD/MM/YYYY"), "")
            JournalType = "SI"
            tmpSuppFee = SuppFee(!refno)
            'tmpPayMCOFees
            
            TotDbAmt = TotDbAmt + IIf(IsNumeric(!TotalMCOFee), !TotalMCOFee, 0) + tmpSuppFee
            TotCrAmt = TotCrAmt + IIf(IsNumeric(!TotalMCOFee), !TotalMCOFee, 0) + tmpSuppFee
            
            For AA = 1 To 2
                
                If (PostingType = "SalesInv" And AA = 1) Or (PostingType = "MBMCRNote" And AA = 2) Then
                    CreditType = "C"
                Else
                    CreditType = "D"
                End If
                If AA = 2 Then lstRec.SubItems(18) = Trim(!refno)
                    If AA = 1 Then
                        If Trim(!MCOPayFeesStatus) = "3" Then
                            AcCode = "1002"
                        ElseIf Trim(!MCOPayFeesStatus) = "2" Then
                            AcCode = "1005"
                        Else
                            AcCode = "1000"
                        End If
                    Else
                    
                        strPaySearch = "Select PAYMCOAccCode, PAYCode from Pay where PAYCode = '" & Trim(tmpPayor) & "'"
                        PaySearch.Open strPaySearch, medixmasterconn, adOpenKeyset, adLockOptimistic
                        
                        If PaySearch.recordCount Then
                            AcCode = PaySearch!PAYMCOAccCode
                        End If
                        
                        PaySearch.Close
                        Set PaySearch = Nothing
                    
                    '    AcCode = "61A002"
                    End If
Continue:
                Set lstRec = lstMBMInvList.ListItems.Add(, , IIf(!SalesStatus = 1, "Post", ""))
                lstRec.SubItems(1) = JournalType
                
                'lstRec.SubItems(2) Journal Source SUN Log-in ID
                lstRec.SubItems(3) = "MED"
                lstRec.SubItems(4) = Trim(!refno)
                'lstRec.SubItems(5) Line No
                If Trim(!AccMonth & !AccYear) <> "" Then
                    lstRec.SubItems(5) = !AccMonth & "/" & !AccYear
                End If
                lstRec.SubItems(6) = IIf(IsDate(!REfDate), Format(!REfDate, "DD/MM/YYYY"), "")
                lstRec.SubItems(7) = Trim(DispName)
                lstRec.SubItems(8) = Trim(IIf(Len(AcCode), AcCode, ""))
                If PostingType = "SalesInv" Then
                    If Trim(!MCOPayFeesStatus) = "1" And PostingType = "SalesInv" And AA = 1 Then
                        tmpTotalMCO = IIf(IsNumeric(!TotalMCOFee), !TotalMCOFee, 0) + tmpSuppFee
                        tmpPayMCO = IIf(IsNumeric(!MCOPayFees), !MCOPayFees, 0)
                        TotalPrinc = IIf(IsNumeric(!NoPrincipal), !NoPrincipal, 0)
                        tmpTotalPayMCO = tmpPayMCO * TotalPrinc
                        If PSStatus = False Then
                            lstRec.SubItems(9) = Format(tmpTotalMCO - tmpTotalPayMCO, "#,##0.00")
                            PSRepeat = True
                        Else
                            lstRec.SubItems(9) = Format(tmpTotalPayMCO, "#,##0.00")
                            PSRepeat = False
                        End If
                    ElseIf Trim(!MCOPayFeesStatus) = "2" And PostingType = "SalesInv" And AA = 1 Then
                        tmpTotalMCO = IIf(IsNumeric(!TotalMCOFee), !TotalMCOFee, 0) + tmpSuppFee
                        tmpPayMCO = IIf(IsNumeric(!MCOPayFees), !MCOPayFees, 0)
                        TotalPrinc = IIf(IsNumeric(!NoPrincipal), !NoPrincipal, 0)
                        tmpTotalPayMCO = tmpPayMCO * TotalPrinc
                        If PSStatus = False Then
                            lstRec.SubItems(9) = Format(tmpTotalMCO - tmpTotalPayMCO, "#,##0.00")
                            PSRepeat = True
                        Else
                            lstRec.SubItems(9) = Format(tmpTotalPayMCO, "#,##0.00")
                            PSRepeat = False
                        End If
                        
                    Else
                        lstRec.SubItems(9) = Format(IIf(IsNumeric(!TotalMCOFee), !TotalMCOFee, 0) + tmpSuppFee, "#,##0.00")
                    End If
                Else
                    lstRec.SubItems(9) = Format(IIf(IsNumeric(!TotalMCOFee), !TotalMCOFee, 0) + tmpSuppFee, "#,##0.00")
                End If
                    'lstRec.SubItems(9) = Format(IIf(IsNumeric(!TotalMCOFee), !TotalMCOFee, 0) + tmpSuppFee, "#,##0.00")
                'End If
                lstRec.SubItems(10) = Trim(CreditType)
                'lstRec.SubItems(12) = tmpCoGRPCode
                lstRec.SubItems(13) = Trim(tmpPayor)
                lstRec.SubItems(15) = Trim(tmpProdCat)
                lstRec.SubItems(17) = IIf(Len(!BatchNo), Trim(!BatchNo), "")
                
                If PostingType = "SalesInv" Then
                    If Trim(!MCOPayFeesStatus) = "1" And PostingType = "SalesInv" And AA = 1 And PSRepeat = True Then
                        AcCode = "8760"
                        PSStatus = True
                        GoTo Continue
                    ElseIf Trim(!MCOPayFeesStatus) = "2" And PostingType = "SalesInv" And AA = 1 And PSRepeat = True Then
                        AcCode = "8761"
                        PSStatus = True
                        GoTo Continue
                    Else
                        PSStatus = False
                    End If
                End If
            Next AA
            LblDbAmt = Format(TotDbAmt, "#,##0.00")
            LblCrAmt = Format(TotCrAmt, "#,##0.00")
            lblCurrent = CurrRec
            
        End With
    End If
NextRecord:
        MBMInv.MoveNext
        DoEvents
        If intExit = True Then
           Exit Do
        End If
    Loop
    intExit = 0
    
    MBMInv.Close
    Set MBMInv = Nothing
 '   medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
  '  medixmasterconn.Close
 '   Set medixmasterconn = Nothing
    If CurrRec = 0 Then
        MsgBox "No record selected ! ", vbOKOnly + vbCritical
    ElseIf LblPost = "No" Then
        FramePostInfo.Enabled = True
    End If

End Sub

Private Sub CboKeyPress(tmpCbo As ComboBox, KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    If tmpCbo.Name <> "cboPostStatus" Then
        NewText = LCase(Left(tmpCbo.Text, tmpCbo.SelStart) + Chr(KeyAscii) + Mid(tmpCbo.Text, tmpCbo.SelStart + tmpCbo.SelLength + 1))
    End If
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To tmpCbo.ListCount - 1

        If NewText = LCase(Left(tmpCbo.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = tmpCbo.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            tmpCbo.Text = ValidValue
            tmpCbo.SelStart = 0
            tmpCbo.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub ClearInfo()
    lstMBMInvList.ListItems.Clear
    lblCurrent = ""
    LblDbAmt = ""
    LblCrAmt = ""
    lblBatchNo = ""
    cboAccPeriod.Text = ""
    'CboAccYear.Text = ""
    FramePostInfo.Enabled = False
End Sub


Private Function SuppFee(tmpREfNo As String)
Dim cmd As New ADODB.Command
Dim App1 As New ADODB.Parameter
Dim App2 As New ADODB.Parameter
Dim SuppRec As New ADODB.Recordset
Dim TotCMcoAmt As Double
Dim SuppSql As String
    TotCMcoAmt = 0
    If (PostingType = "SalesInv") Then
        SuppSql = "Select SUM(MBMCMCO) AS MBMCMCO from MBMCoveredPersons where MBMInvoiceNo = '" & Trim(tmpREfNo) & "'GROUP BY MBMInvoiceNo"
        SuppRec.Open SuppSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do While Not SuppRec.EOF
            TotCMcoAmt = IIf(IsNumeric(SuppRec!MBMCMCO), SuppRec!MBMCMCO, 0)
            SuppRec.MoveNext
        Loop
        SuppRec.Close
        Set SuppRec = Nothing
        
    Else
        SuppSql = "And MBMCRNote = '" & tmpREfNo & "'"
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "SUN_HIS_Posting"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set App1 = cmd.CreateParameter("RptType", adVarChar, adParamInput, 20, "MBMCRSupp")
        cmd.Parameters.Append App1
        Set App2 = cmd.CreateParameter("StrCriterial", adVarChar, adParamInput, 1000, SuppSql)
        cmd.Parameters.Append App2
        SuppRec.CursorLocation = adUseClient
        SuppRec.CursorType = adOpenForwardOnly
        SuppRec.CacheSize = 100
        Set SuppRec = cmd.Execute
        Do While Not SuppRec.EOF
            TotCMcoAmt = IIf(IsNumeric(SuppRec!MBMMCORefund), SuppRec!MBMMCORefund, 0)
            SuppRec.MoveNext
        Loop
        SuppRec.Close
        Set SuppRec = Nothing
    End If
    SuppFee = TotCMcoAmt
End Function

Private Sub MCOIndvListing(strAppendInv As String)
Dim lstRec As ListItem
Dim MBMInv As New ADODB.Recordset
Dim tmpPayor, tmpHltCode, tmpINSCode, tmpPatient, tmpCoGRPCode As String
Dim JournalType, AcCode, CreditType As String
Dim AA, CurrRec As Integer
Dim TotDbAmt, TotCrAmt As Double
Dim cmd As New ADODB.Command
Dim App1 As New ADODB.Parameter
Dim App2 As New ADODB.Parameter
Dim tmpProdCat As String
Dim DispName As String
Dim SuppRec As New ADODB.Recordset
Dim SuppSql As String
Dim tmpSuppFee As Double
Dim tmpPayMCO As Double
Dim TotalPrinc As String
Dim tmpTotalPayMCO As Double
Dim tmpTotalMCO As String
Dim PSStatus As Boolean
Dim PSRepeat As Boolean
Dim strPaySearch As String
Dim PaySearch As New ADODB.Recordset
Dim TempInvNo As String
Dim InvSearch As New ADODB.Recordset
Dim strsqlInvSearch As String
Dim RecCount As Integer

    tmpPayor = ""
    tmpHltCode = ""
    'lstMBMInvList.listitems.Clear
    CurrRec = 0
    LblPost = "No"
    tmpPayMCO = 0
    TotalPrinc = 0
    tmpTotalPayMCO = 0
    TempInvNo = ""
    RecCount = 0
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    
        'If TempInvNo <> MBMInv!MBMInvoiceNo Then
        strsqlInvSearch = " SELECT * From VIEW_SUN_Posting_SalesInvoice Where MBMInvoiceNo = '" & Trim(strAppendInv) & "'"
        InvSearch.Open strsqlInvSearch, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        'RecCount = InvSearch.recordCount
        'InvSearch.Close
        'Set InvSearch = Nothing
        'Else
        '    GoTo NextRecord
        'End If
    
        'If RecCount > 1 Then
        '    'strsqlIdv = strsqlIdv & " And PVNo ='" & Trim(PayHosp!PVNo) & "'"
        '    TempInvNo = MBMInv!MBMInvoiceNo
        '    Call MCOIndvListing(TempInvNo)
        '    CurrRec = CurrRec + 1
        'Else
        Do While Not InvSearch.EOF
        
        'CurrRec = CurrRec + 1
        With InvSearch
            PSStatus = False
            'If InvSearch!SalesStatus = "1" Then
            '    LblPost = "Yes"
            'End If
            tmpPayor = IIf(Len(!PAYCode), !PAYCode, "")
            tmpHltCode = IIf(Len(!HLTCode), !HLTCode, "")
            'tmpCoGRPCode = IIf(Len(!CoGRPCode), !CoGRPCode, "")
            tmpProdCat = "SM"
            If tmpHltCode <> "S" Then
                tmpProdCat = GetProductCat(tmpHltCode, tmpPayor)
            End If
            DispName = tmpPayor
            '& " " & IIf(IsDate(!MBMBordxDate), Format(!MBMBordxDate, "DD/MM/YYYY"), "")
            JournalType = "SI"
            tmpSuppFee = SuppFee(!MBMInvoiceNo)
            'tmpPayMCOFees
            
            TotDbAmt = TotDbAmt + IIf(IsNumeric(!TotalMCOFee), !TotalMCOFee, 0) + tmpSuppFee
            TotCrAmt = TotCrAmt + IIf(IsNumeric(!TotalMCOFee), !TotalMCOFee, 0) + tmpSuppFee
            
            For AA = 1 To 2
                
                If (PostingType = "SalesInv" And AA = 1) Or (PostingType = "MBMCRNote" And AA = 2) Then
                    CreditType = "C"
                Else
                    CreditType = "D"
                End If
                If AA = 2 Then lstRec.SubItems(18) = Trim(!MBMInvoiceNo)
                    If AA = 1 Then
                        If Trim(!MCOPayFeesStatus) = "1" Then
                            AcCode = "1002"
                        ElseIf Trim(!MCOPayFeesStatus) = "2" Then
                            AcCode = "1005"
                        Else
                            AcCode = "1000"
                        End If
                    Else
                    
                        strPaySearch = "Select PAYMCOAccCode, PAYCode from Pay where PAYCode = '" & Trim(tmpPayor) & "'"
                        PaySearch.Open strPaySearch, medixmasterconn, adOpenKeyset, adLockOptimistic
                        
                        If PaySearch.recordCount Then
                            AcCode = PaySearch!PAYMCOAccCode
                        End If
                        
                        PaySearch.Close
                        Set PaySearch = Nothing
                    
                    '    AcCode = "61A002"
                    End If
Continue:
                Set lstRec = lstMBMInvList.ListItems.Add(, , IIf(!SalesStatus = 1, "Post", ""))
                lstRec.SubItems(1) = JournalType
                
                'lstRec.SubItems(2) Journal Source SUN Log-in ID
                lstRec.SubItems(3) = "MED"
                lstRec.SubItems(4) = Trim(!MBMInvoiceNo)
                'lstRec.SubItems(5) Line No
                If Trim(!AccMonth & !AccYear) <> "" Then
                    lstRec.SubItems(5) = !AccMonth & "/" & !AccYear
                End If
                lstRec.SubItems(6) = IIf(IsDate(!MBMInvoiceDate), Format(!MBMInvoiceDate, "DD/MM/YYYY"), "")
                lstRec.SubItems(7) = Trim(DispName)
                lstRec.SubItems(8) = Trim(IIf(Len(AcCode), AcCode, ""))
                If PostingType = "SalesInv" Then
                    If Trim(!MCOPayFeesStatus) = "1" And PostingType = "SalesInv" And AA = 1 Then
                        tmpTotalMCO = IIf(IsNumeric(!TotalMCOFee), !TotalMCOFee, 0) + tmpSuppFee
                        tmpPayMCO = IIf(IsNumeric(!MCOPayFees), !MCOPayFees, 0)
                        TotalPrinc = IIf(IsNumeric(!NoPrincipal), !NoPrincipal, 0)
                        tmpTotalPayMCO = tmpPayMCO * TotalPrinc
                        If PSStatus = False Then
                            lstRec.SubItems(9) = Format(tmpTotalMCO - tmpTotalPayMCO, "#,##0.00")
                            PSRepeat = True
                        Else
                            lstRec.SubItems(9) = Format(tmpTotalPayMCO, "#,##0.00")
                            PSRepeat = False
                        End If
                    Else
                        lstRec.SubItems(9) = Format(IIf(IsNumeric(!TotalMCOFee), !TotalMCOFee, 0) + tmpSuppFee, "#,##0.00")
                    End If
                Else
                    lstRec.SubItems(9) = Format(IIf(IsNumeric(!TotalMCOFee), !TotalMCOFee, 0) + tmpSuppFee, "#,##0.00")
                End If
                    'lstRec.SubItems(9) = Format(IIf(IsNumeric(!TotalMCOFee), !TotalMCOFee, 0) + tmpSuppFee, "#,##0.00")
                'End If
                lstRec.SubItems(10) = Trim(CreditType)
                'lstRec.SubItems(12) = tmpCoGRPCode
                lstRec.SubItems(13) = Trim(tmpPayor)
                lstRec.SubItems(15) = Trim(tmpProdCat)
                lstRec.SubItems(17) = IIf(Len(!BatchNo), Trim(!BatchNo), "")
                
                If PostingType = "SalesInv" Then
                    If Trim(!MCOPayFeesStatus) = "1" And PostingType = "SalesInv" And AA = 1 And PSRepeat = True Then
                        AcCode = "8760"
                        PSStatus = True
                        GoTo Continue
                    ElseIf Trim(!MCOPayFeesStatus) = "2" And PostingType = "SalesInv" And AA = 1 And PSRepeat = True Then
                        AcCode = "8761"
                        PSStatus = True
                        GoTo Continue
                        
                    Else
                        PSStatus = False
                    End If
                End If
            Next AA
            'LblDbAmt = Format(TotDbAmt, "#,##0.00")
            'LblCrAmt = Format(TotCrAmt, "#,##0.00")
            'lblCurrent = CurrRec
            
        End With
    'End If
        InvSearch.MoveNext
        DoEvents
        If intExit = True Then
           Exit Do
        End If
    Loop
    intExit = 0
    
    InvSearch.Close
    Set InvSearch = Nothing
    'medixmasterconnWS.Close
    'Set medixmasterconnWS = Nothing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'If CurrRec = 0 Then
    '    MsgBox "No record selected ! ", vbOKOnly + vbCritical
    'ElseIf LblPost = "No" Then
    '    FramePostInfo.Enabled = True
    'End If
End Sub

Private Function CheckTime(tmpAction As String) As Boolean
Dim TimeRec As New ADODB.Recordset
Dim tmpStartTime As String
Dim TimeSql As String
Dim tmpUserTime As String

    CheckTime = True
    If SpcAccess = "Y" Then Exit Function
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    TimeSql = "Select EndTime from TimeMain where Process ='" & Trim(tmpAction) & "'"
    TimeRec.Open TimeSql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
    Do While Not TimeRec.EOF
        tmpEndTime = TimeRec!EndTime
        TimeRec.MoveNext
    Loop
    TimeRec.Close
    Set TimeRec = Nothing
 '   medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
End Function


