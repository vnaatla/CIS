VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvPosting 
   Caption         =   "Clinic Invoice Posting "
   ClientHeight    =   7968
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7968
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   7320
      Width           =   11775
      Begin VB.TextBox lblCreditAmt 
         Height          =   285
         Left            =   4800
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox LblDebitAmt 
         Height          =   285
         Left            =   2640
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   9480
         TabIndex        =   16
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   8160
         TabIndex        =   15
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   7440
         TabIndex        =   5
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10440
         TabIndex        =   14
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   11040
         TabIndex        =   13
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "&Post"
         Height          =   375
         Left            =   8760
         TabIndex        =   6
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Credit Amt"
         Height          =   255
         Left            =   3960
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Debit Amt"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   810
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Width           =   11895
      Begin VB.ComboBox cboPAYCode 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   180
         Width           =   3495
      End
      Begin VB.TextBox TxtINVNoFr 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   470
         Width           =   1455
      End
      Begin VB.TextBox TxtINVNoTo 
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Top             =   470
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtINVRegDateTo 
         Height          =   285
         Left            =   7320
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   23
         Top             =   440
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Posting Status"
         Height          =   255
         Left            =   4800
         TabIndex        =   24
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Inv No from"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   500
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "INV Reg Date"
         Height          =   255
         Left            =   4800
         TabIndex        =   9
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   7080
         TabIndex        =   8
         Top             =   195
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView lstInvPosting 
      Height          =   6015
      Left            =   0
      TabIndex        =   19
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
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "A/C Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Period"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Based Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Debit/Credit"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Trans Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Journal Type"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Trans Ref"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Desc"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "T1 Clinic"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "T2 Plan No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "T3 Pay Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "T4  InvDORMth"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "T5 Grp Com"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "T6 Bordx Ref"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Status"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Date"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label38 
      Caption         =   "Invoice Posting List"
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
      TabIndex        =   25
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Clinic Invoice Posting "
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
      TabIndex        =   20
      Top             =   0
      Width           =   11925
   End
End
Attribute VB_Name = "frmInvPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public searchCriteria As String
Public intExit As Integer
Dim strAppend As String
Dim SearchStatus As Boolean

Private Sub INVPostExcel(Optional strAppend As String)
Const METHOD_NAME = "INVPostExcel"
On Error GoTo debugs
Dim rst As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalDebitAmt As Variant
Dim TotalCreditAmt As Variant
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.Worksheet
Dim RecCnt As Integer
Dim RecCnt2 As Integer
    total = 0
    Current = 0
    
    If Mid(cboPostStatus, 1, 1) = "0" Then
        sql = " Select PAYCode, CLMPayableByMedix2C, InvoiceNo, MBMGroupCompany,PAYCode, " & _
              " InvoiceDate, BordxRefNo, ClinicCode, CLNName,RegisterDate, GRPCoCode, CASInvPost " & _
              " From VIEW_Inv_Posting where CASInvPost = '0'"
    ElseIf Mid(cboPostStatus, 1, 1) = "1" Then
        sql = " Select PAYCode, CLMPayableByMedix2C, InvoiceNo, MBMGroupCompany,PAYCode, " & _
          " InvoiceDate, BordxRefNo, ClinicCode, CLNName,RegisterDate, GRPCoCode, CASInvPost " & _
          " From VIEW_Inv_Posting where CASInvPost = '1'"
    End If
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst.Open sql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    'If rst.RecordCount = 0 Then
    '    MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    '    CmdClear.Enabled = True
    '    CmdPrintFile.Enabled = True
    '    CmdExit.Enabled = True
    'Else
    
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
    
        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\InvoicePosting")
        Set objWorksheet = objWorkbook.Sheets(1)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
        RecCnt = 0
        Do Until rst.EOF
            RecCnt = RecCnt + 1
            With rst
                '------- Debit Start --------------'
                objWorksheet.Cells(1 + CDbl(RecCnt), "A") = "6000"
                
                If Len(!InvoiceDate) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "B") = Format(!InvoiceDate, "mm/yyyy")
                End If
                
                If Len(!CLMPayableByMedix2C) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "C") = Format(!CLMPayableByMedix2C, "#,##0.00")
                End If
                   
                objWorksheet.Cells(1 + CDbl(RecCnt), "D") = "D"
                
                If Len(!InvoiceDate) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "E") = Format(!InvoiceDate, "dd/mm/yyyy")
                End If
                
                objWorksheet.Cells(1 + CDbl(RecCnt), "F") = "RC"
                
                If Len(!InvoiceNo) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "G") = !InvoiceNo
                End If

                If Len(!Bordxrefno) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "H") = !Bordxrefno
                End If
                
                If Len(!ClinicCode) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "I") = !ClinicCode
                End If
                
                objWorksheet.Cells(1 + CDbl(RecCnt), "J") = ""
                
                If Len(!PAYCode) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "K") = !PAYCode
                End If
                
                If Len(!InvoiceDate) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "L") = Format(!RegisterDate, "YYMM")
                End If
                
                If Len(!GRPCoCode) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "M") = (!GRPCoCode)
                End If
                
                If Len(!Bordxrefno) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "N") = Mid(!Bordxrefno, 4, 5)
                End If
                If Len(!CLMPayableByMedix2C) Then
                    TotalDebitAmt = TotalDebitAmt + !CLMPayableByMedix2C
                End If

                '------- Debit End --------------'
                
                '------- Credit Start --------------'
                RecCnt = RecCnt + 1
                objWorksheet.Cells(1 + CDbl(RecCnt), "A") = "84" & !ClinicCode
                
                If Len(!InvoiceDate) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "B") = Format(!InvoiceDate, "mm/yyyy")
                End If
                
                If Len(!CLMPayableByMedix2C) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "C") = Format(!CLMPayableByMedix2C, "#,##0.00")
                End If
                   
                objWorksheet.Cells(1 + CDbl(RecCnt), "D") = "C"
                
                If Len(!InvoiceDate) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "E") = Format(!InvoiceDate, "dd/mm/yyyy")
                End If
                
                objWorksheet.Cells(1 + CDbl(RecCnt), "F") = "RC"
                
                If Len(!InvoiceNo) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "G") = !InvoiceNo
                End If

                If Len(!CLNName) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "H") = !CLNName
                End If
                
                If Len(!ClinicCode) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "I") = !ClinicCode
                End If
                
                'If Len(!BordxPAYRecDate) Then
                objWorksheet.Cells(1 + CDbl(RecCnt), "J") = ""
                'End If
                
                If Len(!PAYCode) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "K") = !PAYCode
                End If
                
                If Len(!InvoiceDate) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "L") = Format(!RegisterDate, "YYMM")
                End If
                
                If Len(!GRPCoCode) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "M") = (!GRPCoCode)
                End If
                
                If Len(!Bordxrefno) Then
                    objWorksheet.Cells(1 + CDbl(RecCnt), "N") = Mid(!Bordxrefno, 4, 5)
                End If
                If Len(!CLMPayableByMedix2C) Then
                    TotalCreditAmt = TotalCreditAmt + !CLMPayableByMedix2C
                End If
                
                '------- Credit Start --------------'
                Call UpdateCase(!InvoiceNo, !ClinicCode)
            End With
            rst.MoveNext
        Loop
            'cmdCredit.Enabled = False
            objExcel.DisplayAlerts = True
            'objWorksheet.Columns.AutoFit
            objExcel.Interactive = True
            Screen.MousePointer = vbDefault
    'End If
    
    rst.Close
    Set rst = Nothing
    CmdClear.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
    cmdPost.Enabled = True
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
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub UpdateCase(InvNo As String, CLNCode As String)
Const METHOD_NAME = "UpdateCase"
On Error GoTo debugs
Dim CAS As New ADODB.Recordset
Dim sql As String

    If Mid(cboPostStatus, 1, 1) = "0" Then
        sql = " Select ClinicCode, InvoiceNo, CASInvPost,CASInvPostUser ,CASInvPostDate From CISCAS where InvoiceNo = '" & Trim(InvNo) & "' And ClinicCode = '" & Trim(CLNCode) & "' and CASInvPost = '0'"
        CAS.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If CAS.RecordCount Then
            Do Until CAS.EOF
                CAS!CASInvPost = "1"
                CAS!CASInvPostUser = Logon
                CAS!CASInvPostDate = Date
            CAS.update
            CAS.MoveNext
            Loop
            
        End If
        CAS.Close
        Set CAS = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub cboPAYCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPAYCode.Text, cboPAYCode.SelStart) + Chr(KeyAscii) + Mid(cboPAYCode.Text, cboPAYCode.SelStart + cboPAYCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPAYCode.ListCount - 1
 
        If NewText = LCase(Left(cboPAYCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPAYCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPAYCode.Text = ValidValue
                cboPAYCode.SelStart = 0
                cboPAYCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboPostStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPostStatus.Text, cboPostStatus.SelStart) + Chr(KeyAscii) + Mid(cboPostStatus.Text, cboPostStatus.SelStart + cboPostStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPAYCode.ListCount - 1
 
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

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstInvPosting.ListItems.Clear
    lblCurrent = 0
    LblDebitAmt = 0
    lblCreditAmt = 0
    Call ComboListing
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPost_Click()
    Screen.MousePointer = vbHourglass
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
        SearchStatus = False
        Call SearchRecord
        Call INVPostExcel(strAppend)
    
       ' MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
       ' MediConnCISWS.Close
       ' Set MediConnCISWS = Nothing
    Screen.MousePointer = Default
End Sub

Private Sub CmdPrintFile_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelBordxControlNonSihat(lstInvPosting)
    Screen.MousePointer = Default
End Sub

Private Sub CmdSearch_Click()
    LblDebitAmt = ""
    lblCreditAmt = ""
    lblCurrent = ""
    lblTotal = ""
    SearchStatus = True
    CmdClear.Enabled = False
    CmdPrintFile.Enabled = False
    cmdPost.Enabled = False
    CmdExit.Enabled = False

    Screen.MousePointer = vbHourglass

        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
            Call SearchRecord
            
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
        'medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        'MediConnCISWS.Close
       ' Set MediConnCISWS = Nothing
        
    Screen.MousePointer = Default
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
    cmdPost.Enabled = True
End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "frmInvPosting")
    Call ComboListing
End Sub


Private Sub txtINVRegDateFr_LostFocus()
    If txtINVRegDateFr <> "__/__/____" Then
        If IsDate(txtINVRegDateFr) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtINVRegDateFr.SetFocus
        ElseIf CDate(txtINVRegDateFr) > CDate(Date) Then
            MsgBox "Invoice Date cannot greater than system date ! ", vbCritical
            txtINVRegDateFr.SetFocus
        End If
    End If
End Sub


Private Sub txtINVRegDateTo_LostFocus()
    If txtINVRegDateTo <> "__/__/____" Then
        If IsDate(txtINVRegDateTo) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtINVRegDateTo.SetFocus
        ElseIf CDate(txtINVRegDateTo) > CDate(Date) Then
            MsgBox "Invoice Date cannot greater than system date ! ", vbCritical
            txtINVRegDateTo.SetFocus
        End If
    End If
End Sub

Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs
Dim strsql As String
Dim sql As String
Dim CAS As New ADODB.Recordset
Dim getSel As Boolean
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim GrpCompanyName As String
    
    
    strAppend = ""
    
    If Len(Trim(cboPAYCode)) Then
        strAppend = strAppend + " AND PAYCode = '" & Mid(cboPAYCode, 1, 2) & "'"
    End If
    
    If Len(Trim(TxtINVNoFr)) Then
        strAppend = strAppend + " AND InvoiceNo >= '" & RemStr(Trim(TxtINVNoFr)) & "'"
    End If
    
    If Len(Trim(TxtINVNoTo)) Then
        strAppend = strAppend + " AND InvoiceNo <= '" & RemStr(Trim(TxtINVNoTo)) & "'"
    End If
    
    If IsDate(txtINVRegDateFr) = True Then
        strAppend = strAppend + " And RegisterDate >= '" & Format(txtINVRegDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(txtINVRegDateTo) = True Then
        strAppend = strAppend + " And RegisterDate <= '" & Format(txtINVRegDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(cboPostStatus) Then
        strAppend = strAppend + " And  CASInvPost  = '" & Mid(cboPostStatus, 1, 1) & "'"
    End If

    
    
    searchCriteria = strAppend
    
    If SearchStatus = True Then
        Call INVPostListing(strAppend)
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Function

Private Sub INVPostListing(Optional strAppend As String)
Const METHOD_NAME = "INVPostListing"
On Error GoTo debugs
Dim rst As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalDebitAmt As Variant
Dim TotalCreditAmt As Variant
Dim INVSum As New ADODB.Recordset
Dim strsql As String
    total = 0
    Current = 0
        
    lstInvPosting.ListItems.Clear
            
    'strCount = "select count(*) as totalRecord From VIEW_CISBordx_ControlN Where 0=0 "

    sql = " Select PAYCode, CLMPayableByMedix2C, InvoiceNo, PAYCode, CASInvPostUser, CASInvPostDate," & _
          " InvoiceDate, ClinicCode, RegisterDate, CASInvPost, GRPCoCode,BordxRefNo" & _
          " From VIEW_InvSummary_Posting where 0 = 0 "
       
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    '    strCount = strCount + strAppend
    End If
    rst.Open sql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    
    Do
        Do Until rst.EOF
        'total = rstCount!totalrecord
        'lblTotal = CInt(total)
            With rst
                '------- Credit Start --------------'
                Set Record = lstInvPosting.ListItems.Add(, , "84" & !ClinicCode)
                
                If Len(!InvoiceDate) Then
                    Record.SubItems(1) = Format(!InvoiceDate, "mm/yyyy")
                End If
                
                If Len(!CLMPayableByMedix2C) Then
                    Record.SubItems(2) = Format(!CLMPayableByMedix2C, "#,##0.00")
                End If
                   
                Record.SubItems(3) = "C"
                
                If Len(!InvoiceDate) Then
                    Record.SubItems(4) = Format(!InvoiceDate, "dd/mm/yyyy")
                End If
                
                Record.SubItems(5) = "RC"
                
                If Len(!InvoiceNo) Then
                    Record.SubItems(6) = !InvoiceNo
                End If

                If Len(!InvoiceNo) Then
                    Record.SubItems(7) = !InvoiceNo
                End If
                
                If Len(!ClinicCode) Then
                    Record.SubItems(8) = !ClinicCode
                End If
                
                'If Len(!BordxPAYRecDate) Then
                '    Record.SubItems(9) = !BordxPAYRecDate
                'End If
                
                If Len(!PAYCode) Then
                    Record.SubItems(10) = !PAYCode
                End If
                
                If Len(!RegisterDate) Then
                    Record.SubItems(11) = Format(!RegisterDate, "YYMM")
                End If
                
                If Len(!GRPCoCode) Then
                    Record.SubItems(12) = (!GRPCoCode)
                End If
                
                If Len(!Bordxrefno) Then
                    Record.SubItems(13) = Mid(!Bordxrefno, 4, 5)
                End If
                
                If Len(!CASInvPost) Then
                    If !CASInvPost = "0" Then
                        Record.SubItems(14) = "No"
                    Else
                        Record.SubItems(14) = "Yes"
                    End If
                End If
                
                If Len(!CASInvPostUser) Then
                    Record.SubItems(15) = !CASInvPostUser
                End If
                
                If Len(!CASInvPostDate) Then
                    Record.SubItems(16) = !CASInvPostDate
                End If
                
                
                '------- Credit End --------------'
                
                '------- Debit Start --------------'
            strsql = " Select PAYCode, CLMPayableByMedix2C, InvoiceNo, MBMGroupCompany,PAYCode, CASInvPostUser, CASInvPostDate," & _
                  " InvoiceDate, BordxRefNo, ClinicCode, CLNName,RegisterDate, GRPCoCode, CASInvPost" & _
                  " From VIEW_Inv_Posting where InvoiceNo = '" & rst!InvoiceNo & "' And  RegisterDate = '" & Format(rst!RegisterDate, "yyyy/mm/dd") & "'"
            INVSum.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
            
            Do Until INVSum.EOF
            With INVSum
                If !PAYCode = "DC" Then
                    Set Record = lstInvPosting.ListItems.Add(, , "60" & !GRPCoCode)
                Else
                    Set Record = lstInvPosting.ListItems.Add(, , "60" & !PAYCode)
                End If
                If Len(!InvoiceDate) Then
                    Record.SubItems(1) = Format(!InvoiceDate, "mm/yyyy")
                End If
                
                If Len(!CLMPayableByMedix2C) Then
                    Record.SubItems(2) = Format(!CLMPayableByMedix2C, "#,##0.00")
                End If
                   
                Record.SubItems(3) = "D"
                
                If Len(!InvoiceDate) Then
                    Record.SubItems(4) = Format(!InvoiceDate, "dd/mm/yyyy")
                End If
                
                Record.SubItems(5) = "RC"
                
                If Len(!InvoiceNo) Then
                    Record.SubItems(6) = !InvoiceNo
                End If

                If Len(!Bordxrefno) Then
                    Record.SubItems(7) = !Bordxrefno
                End If
                
                If Len(!ClinicCode) Then
                    Record.SubItems(8) = !ClinicCode
                End If
                
                'If Len(!BordxPAYRecDate) Then
                Record.SubItems(9) = ""
                'End If
                
                If Len(!PAYCode) Then
                    Record.SubItems(10) = !PAYCode
                End If
                
                If Len(!RegisterDate) Then
                    Record.SubItems(11) = Format(!RegisterDate, "YYMM")
                End If
                
                If Len(!GRPCoCode) Then
                    Record.SubItems(12) = (!GRPCoCode)
                End If
                
                If Len(!Bordxrefno) Then
                    Record.SubItems(13) = Mid(!Bordxrefno, 4, 5)
                End If
                
                If Len(!CASInvPost) Then
                    If !CASInvPost = "0" Then
                        Record.SubItems(14) = "No"
                    Else
                        Record.SubItems(14) = "Yes"
                    End If
                End If
                
                If Len(!CASInvPostUser) Then
                    Record.SubItems(15) = !CASInvPostUser
                End If
                
                If Len(!CASInvPostDate) Then
                    Record.SubItems(16) = !CASInvPostDate
                End If
                
                If Len(!CLMPayableByMedix2C) Then
                    TotalDebitAmt = TotalDebitAmt + !CLMPayableByMedix2C
                End If
            End With
            INVSum.MoveNext
            Loop
            INVSum.Close
            Set INVSum = Nothing
                '------- Debit End --------------'
                
                
                Current = Current + 1
                lblCurrent = Current
                If Len(!CLMPayableByMedix2C) Then
                    TotalCreditAmt = TotalCreditAmt + !CLMPayableByMedix2C
                End If
                
                
                
            End With
            rst.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    lblCreditAmt = Format(TotalCreditAmt, "#,##0.00")
    LblDebitAmt = Format(TotalDebitAmt, "#,##0.00")
    'End If
    
    intExit = 0
    rst.Close
    Set rst = Nothing
    CmdClear.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
    cmdPost.Enabled = True

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub ComboListing()
Const METHOD_NAME = "ComboListing"
On Error GoTo debugs

    Dim PAY As New ADODB.Recordset
    Dim sqlPAY As String

   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
    cboPAYCode.Clear
    sqlPAY = "Select  PAYCode, PAYCompanyName From PAY where 0 = 0"
    PAY.Open sqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until PAY.EOF
        cboPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    
    PAY.Close
    Set PAY = Nothing
    
    cboPostStatus.Clear
    cboPostStatus.AddItem "1 - Yes"
    cboPostStatus.AddItem "0 - No"
    cboPostStatus.Text = "0" & " - " & "No"
    
    'MediConnCISDB.Close
  '  Set MediConnCISDB = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
  '  MediConnCISWS.Close
  '  Set MediConnCISWS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub
