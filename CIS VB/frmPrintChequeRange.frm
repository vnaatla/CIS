VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintChequeRange 
   Caption         =   "Print Cheque Range"
   ClientHeight    =   7476
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   7968
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7476
   ScaleWidth      =   7968
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   6840
      Width           =   7935
      Begin VB.CommandButton cmdPrintACCNo 
         Caption         =   "Print Acc No"
         Height          =   375
         Left            =   5760
         TabIndex        =   21
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txtPrintingStatus 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Text            =   "Printing Not In Progress !"
         Top             =   195
         Width           =   2175
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   6840
         TabIndex        =   15
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print Cheque"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4680
         TabIndex        =   14
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Printing Status"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblPymtType 
         Caption         =   "Label5"
         Height          =   255
         Left            =   5520
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTotalchqSel 
         Caption         =   "Label11"
         Height          =   255
         Left            =   4800
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTotalChqList 
         Caption         =   "Label11"
         Height          =   255
         Left            =   3840
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7935
      Begin VB.CheckBox chkPrintAccNo 
         Caption         =   "Print Account No"
         Height          =   375
         Left            =   5400
         TabIndex        =   22
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtChequeNoFr 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox txtChequeNoTo 
         Height          =   285
         Left            =   2880
         TabIndex        =   8
         Top             =   180
         Width           =   1455
      End
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   2880
         TabIndex        =   7
         Top             =   1440
         Width           =   3375
      End
      Begin VB.OptionButton optSihat 
         Caption         =   "CA - Sihat Citi1"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optNSMAA 
         Caption         =   "CC - NS MAA"
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   150
         Width           =   735
      End
      Begin VB.OptionButton optNS 
         Caption         =   "CB - NS Citi2"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Cheque No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Account"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lstChequeListing 
      Height          =   6255
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   7935
      _ExtentX        =   13991
      _ExtentY        =   11028
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cheque No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cheque Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Cheque Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Account No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Payment Type"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Principal Name"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Actual Amt"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Print Cheque Range"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmPrintChequeRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public chqMode As String
Dim clickStatus As String
Private Sub Form_Load()
  Call ActionLogUpdate(12, "frmPrintChequeRange")
End Sub

Private Sub cboAccount_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboAccount.Text, cboAccount.SelStart) + Chr(KeyAscii) + Mid(cboAccount.Text, cboAccount.SelStart + cboAccount.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboAccount.ListCount - 1
 
        If NewText = LCase(Left(cboAccount.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboAccount.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboAccount.Text = ValidValue
                cboAccount.SelStart = 0
                cboAccount.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub chkPrintAccNo_Click()
    If chkPrintAccNo.Value = 1 Then
        cmdPrintACCNo.Enabled = True
        cmdPrint.Enabled = False
        lstChequeListing.ListItems.Clear
    Else
        cmdPrintACCNo.Enabled = False
        cmdPrint.Enabled = True
        lstChequeListing.ListItems.Clear
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Const METHOD_NAME = "cmdPrint_Click"
On Error GoTo debugs

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim StrSql As String
Dim chq As New ADODB.Recordset
Dim RecFound As Boolean
Dim RecCnt As Integer
Dim tmpDD As String
Dim tmpMM As String
Dim tmpYY As String
    Screen.MousePointer = vbHourglass
    If lstChequeListing.ListItems.Count > 0 Then
        If MsgBox("Do you want to print now?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
    
            'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
            RecFound = False
            txtPrintingStatus = "Printing In Progress!"
            cmdPrint.Enabled = False
            cmdExit.Enabled = False
            If chqMode = "MedixPayMBM" Then
                Call ChequePrintingMBM
            ElseIf chqMode = "MedixPayCLN" Then
                Call ChequePrintingCLN
            End If
        End If
    Else
        MsgBox "Please select cheque to print", vbOKOnly + vbCritical
    End If
    Screen.MousePointer = vbDefault

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub ChequePrintingMBM()
Const METHOD_NAME = "ChequePrintingMBM"
On Error GoTo debugs

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim StrSql As String
Dim chq As New ADODB.Recordset
Dim RecFound As Boolean
Dim RecCnt As Integer
Dim tmpDD As String
Dim tmpMM As String
Dim tmpYY As String
Dim strsqlUpd As String
Dim RecordFound As Boolean
    RecordFound = False
    
     '   MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
     '   MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
     '   MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
            StrSql = "select PVChequeNo, PVChequeDate, PVCLNChequeAmt, " & _
            " PVWordingOne,PVWordingTwo, PVPayto,PVPayeeACNo,MBMGroupCompany from VIEW_Payment2MBM_Chq where PVChequeNo >= '" & Trim(txtChequeNoFr) & "' And PVChequeNo <= '" & Trim(txtChequeNoTo) & "'  order by PVChequeNo"
            chq.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
            
            Do While Not chq.EOF
                RecordFound = True
                If objExcel Is Nothing Then
                    Set objExcel = CreateObject("Excel.application")
                        objExcel.Interactive = False
                    'If objExcel.Visible = False Then
                        objExcel.Visible = False
                    'End If
                    objExcel.WindowState = xlMaximized
                    Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "PaymentVoucher\PaymentChequeCIS.xlt")
                    Set objWorksheet = objWorkbook.Sheets(1)
                    objExcel.DisplayAlerts = False
                    Set objWorksheet = objWorkbook.Worksheets.Item(1)
                End If
                    
                On Error GoTo HANDLE_ERROR
                    
                    If chq!MBMGroupCompany = "IBM MALAYSIA SDN BHD" Then
                        objWorksheet.Cells(2, "A") = chq!PVPayto
                        objWorksheet.Cells(3, "A") = chq!PVPayeeACNo
                    Else
                        objWorksheet.Cells(3, "A") = chq!PVPayto
                        'objWorksheet.Cells(3, "A") = chq!PVPayeeACNo
                    End If
                    objWorksheet.Cells(5, "A") = chq!PVWordingOne
                    objWorksheet.Cells(7, "A") = chq!PVWordingTwo
                    objWorksheet.Cells(5, "E") = Format(chq!PVCLNChequeAmt, "#,##0.00")
                    tmpDD = Format(chq!PVChequeDate, "DD")
                    tmpMM = Format(chq!PVChequeDate, "MM")
                    tmpYY = Format(chq!PVChequeDate, "YY")
                    objWorksheet.Cells(1, "H") = Mid(tmpDD, 1, 1)
                    objWorksheet.Cells(1, "I") = Mid(tmpDD, 2, 1)
                    objWorksheet.Cells(1, "J") = Mid(tmpMM, 1, 1)
                    objWorksheet.Cells(1, "K") = Mid(tmpMM, 2, 1)
                    objWorksheet.Cells(1, "L") = Mid(tmpYY, 1, 1)
                    objWorksheet.Cells(1, "M") = Mid(tmpYY, 2, 1)
                    objWorkbook.PrintOut
                    objWorkbook.SaveAs (LocCardPath & "MemeberCHQ\" & Trim(txtChequeNoFr) & ".xlt")
                    objExcel.DisplayAlerts = True
                    objExcel.Interactive = True
                    Set objWorksheet = Nothing
                    Set objWorkbook = Nothing
                    objExcel.Quit
                    Set objExcel = Nothing
                    
                chq.MoveNext
            Loop
            chq.Close
            Set chq = Nothing
            
            If RecordFound = True Then
                strsqlUpd = "Update Payment2MBM set PVChequeStatus = 'Y' Where (PVChequeNo >= '" & Trim(txtChequeNoFr) & "' And  PVChequeNo <= '" & Trim(txtChequeNoTo) & "')"
                MediConnCISPayment.Execute strsqlUpd
            End If

      '  MediConnCISDB.Close
      '  Set MediConnCISDB = Nothing
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
      '  MediConnCISWS.Close
      '  Set MediConnCISWS = Nothing
      '  MediConnCISPayment.Close
      '  Set MediConnCISPayment = Nothing
ExitFunction:
            txtPrintingStatus = "Printing Not In Progress!"
            cmdPrint.Enabled = False
            cmdExit.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            objExcel.Quit
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            GoTo ExitFunction
            txtPrintingStatus = "Printing Not In Progress!"
            cmdPrint.Enabled = True
            
            If chkPrintAccNo.Value = 1 Then
                cmdPrintACCNo.Enabled = True
                cmdPrint.Enabled = False
            Else
                cmdPrintACCNo.Enabled = False
                cmdPrint.Enabled = True
            End If
            cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub ChequePrintingMBMAcc()
Const METHOD_NAME = "ChequePrintingMBMAcc"
On Error GoTo debugs

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim StrSql As String
Dim chq As New ADODB.Recordset
Dim RecFound As Boolean
Dim RecCnt As Integer
Dim tmpDD As String
Dim tmpMM As String
Dim tmpYY As String
Dim strsqlUpd As String
Dim RecordFound As Boolean
    RecordFound = False
    StrSql = "select PVChequeNo, PVChequeDate, PVCLNChequeAmt, " & _
    " PVWordingOne,PVWordingTwo, PVPayto,PVPayeeACNo,MBMGroupCompany from VIEW_Payment2MBM_Acc where PVChequeNo >= '" & Trim(txtChequeNoFr) & "' And PVChequeNo <= '" & Trim(txtChequeNoTo) & "'  order by PVChequeNo"
    chq.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
    
    Do While Not chq.EOF
        RecordFound = True
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            'If objExcel.Visible = False Then
                objExcel.Visible = False
            'End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "PaymentVoucher\tmpPaychequeAcc.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
            
        On Error GoTo HANDLE_ERROR
            
            'If chq!MBMGroupCompany = "IBM MALAYSIA SDN BHD" Then
                objWorksheet.Cells(2, "A") = chq!PVPayeeACNo
                objWorksheet.Cells(3, "A") = chq!PVPayto
            'Else
            '    objWorksheet.Cells(3, "A") = chq!PVPayto
            '    'objWorksheet.Cells(3, "A") = chq!PVPayeeACNo
            'End If
            'objWorksheet.Cells(5, "A") = chq!PVWordingOne
            'objWorksheet.Cells(7, "A") = chq!PVWordingTwo
            'objWorksheet.Cells(5, "E") = Format(chq!PVCLNChequeAmt, "#,##0.00")
            'tmpDD = Format(chq!PVChequeDate, "DD")
            'tmpMM = Format(chq!PVChequeDate, "MM")
            'tmpYY = Format(chq!PVChequeDate, "YY")
            'objWorksheet.Cells(1, "H") = Mid(tmpDD, 1, 1)
            'objWorksheet.Cells(1, "I") = Mid(tmpDD, 2, 1)
            'objWorksheet.Cells(1, "J") = Mid(tmpMM, 1, 1)
            'objWorksheet.Cells(1, "K") = Mid(tmpMM, 2, 1)
            'objWorksheet.Cells(1, "L") = Mid(tmpYY, 1, 1)
            'objWorksheet.Cells(1, "M") = Mid(tmpYY, 2, 1)
    
            objWorkbook.PrintOut
            objWorkbook.SaveAs (LocCardPath & "PaymentVoucher\" & "tmpPaycheque" & ".xlt")
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            objExcel.Quit
            Set objExcel = Nothing
            
        chq.MoveNext
    Loop
    chq.Close
    Set chq = Nothing
    
ExitFunction:
            txtPrintingStatus = "Printing Not In Progress!"
            cmdExit.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            objExcel.Quit
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            GoTo ExitFunction
            txtPrintingStatus = "Printing Not In Progress!"
            cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub


Private Sub ChequePrintingCLN()
Const METHOD_NAME = "ChequePrintingCLN"
On Error GoTo debugs

'On Error Resume Next
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim StrSql As String
Dim chq As New ADODB.Recordset
Dim RecFound As Boolean
Dim RecCnt As Integer
Dim tmpDD As String
Dim tmpMM As String
Dim tmpYY As String
Dim strsqlUpd As String

            StrSql = "select PVCLNChequeNo, PVCLNChequeDate, PVCLNChequeAmt, " & _
            " PVWordingOne,PVWordingTwo, CLNPName from VIEW_Payment2CLN_Chq where PVCLNChequeNo >= '" & Trim(txtChequeNoFr) & "' And PVCLNChequeNo <= '" & Trim(txtChequeNoTo) & "'  order by PVCLNChequeNo"
            chq.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
            
            Do While Not chq.EOF
                If objExcel Is Nothing Then
                    Set objExcel = CreateObject("Excel.application")
                        objExcel.Interactive = False
                        objExcel.Visible = False
                    objExcel.WindowState = xlMaximized
                    Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "PaymentVoucher\PaymentCheque.xlt")
                    Set objWorksheet = objWorkbook.Sheets(1)
                    objExcel.DisplayAlerts = False
                    Set objWorksheet = objWorkbook.Worksheets.Item(1)
                End If
                    
                On Error GoTo HANDLE_ERROR
            
                    objWorksheet.Cells(3, "A") = chq!CLNPName
                    objWorksheet.Cells(5, "A") = chq!PVWordingOne
                    objWorksheet.Cells(7, "A") = chq!PVWordingTwo
                    objWorksheet.Cells(5, "E") = chq!PVCLNChequeAmt
                    tmpDD = Format(chq!PVCLNChequeDate, "DD")
                    tmpMM = Format(chq!PVCLNChequeDate, "MM")
                    tmpYY = Format(chq!PVCLNChequeDate, "YY")
                    objWorksheet.Cells(1, "H") = Mid(tmpDD, 1, 1)
                    objWorksheet.Cells(1, "I") = Mid(tmpDD, 2, 1)
                    objWorksheet.Cells(1, "J") = Mid(tmpMM, 1, 1)
                    objWorksheet.Cells(1, "K") = Mid(tmpMM, 2, 1)
                    objWorksheet.Cells(1, "L") = Mid(tmpYY, 1, 1)
                    objWorksheet.Cells(1, "M") = Mid(tmpYY, 2, 1)
        
                    objWorkbook.PrintOut
                    objWorkbook.SaveAs (LocCardPath & "PaymentVoucher\" & "tmpPaycheque" & ".xlt")
                    Set objWorksheet = Nothing
                    Set objWorkbook = Nothing
                    objExcel.Quit
                    Set objExcel = Nothing
        
                chq.MoveNext
            Loop
            chq.Close
            Set chq = Nothing
            
            strsqlUpd = "Update Payment2CLN set PVChequeStatus = 'Y' where (PVCLNChequeNo >= '" & Trim(txtChequeNoFr) & "' And  PVCLNChequeNo <= '" & Trim(txtChequeNoTo) & "')"
            MediConnCISPayment.Execute strsqlUpd
   
ExitFunction:
            txtPrintingStatus = "Printing Not In Progress!"
            cmdPrint.Enabled = False
            cmdExit.Enabled = True
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            txtPrintingStatus = "Printing Not In Progress!"
            cmdPrint.Enabled = True
            cmdExit.Enabled = True

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
            
End Sub

Private Sub cmdPrintACCNo_Click()
Const METHOD_NAME = "cmdPrintACCNo_Click"
On Error GoTo debugs

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim StrSql As String
Dim chq As New ADODB.Recordset
Dim RecFound As Boolean
Dim RecCnt As Integer
Dim tmpDD As String
Dim tmpMM As String
Dim tmpYY As String
    Screen.MousePointer = vbHourglass
    If lstChequeListing.ListItems.Count > 0 Then
        If MsgBox("Do you want to print now?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
            RecFound = False
            txtPrintingStatus = "Printing In Progress!"
            cmdPrint.Enabled = False
            cmdExit.Enabled = False
            Call ChequePrintingMBMAcc
        End If
    Else
        MsgBox "Please select cheque to print", vbOKOnly + vbCritical
    End If
    Screen.MousePointer = vbDefault
    

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdSearch_Click()
Dim RecSel As Double
    RecSel = 0
    lblTotalchqSel = 0
    If txtChequeNoFr = "" Then
        MsgBox "Please Enter Cheque No", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If txtChequeNoTo = "" Then
        MsgBox "Please Enter Cheque No", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    RecSel = (CDbl(Trim(Mid(txtChequeNoTo, 4, 8))) - CDbl(Trim(Mid(txtChequeNoFr, 4, 8)))) + 1
    If RecSel <= 0 Then
        MsgBox "Please ensure cheque range entered correctly", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    lblTotalchqSel = RecSel
    lstChequeListing.ListItems.Clear
            If chqMode = "MedixPayMBM" Then
                Call ChequeMBMListing
            ElseIf chqMode = "MedixPayCLN" Then
                Call ChequeCLNListing
            End If
End Sub

Private Sub ChequeCLNListing()
Const METHOD_NAME = "ChequeCLNListing"
On Error GoTo debugs



Dim StrSql As String
Dim chq As New ADODB.Recordset
Dim RecFound As Boolean
Dim RecCnt As Integer
Dim Record As ListItem
Dim RecordFound As Boolean
    RecordFound = False
    RecCnt = 0
    lblTotalChqList = 0
    cmdPrint.Enabled = False
    cmdExit.Enabled = False
    
    StrSql = ""
    StrSql = "select PVCLNChequeNo, PVCLNChequeDate, PVCLNChequeAmt, " & _
    " PVWordingOne,PVWordingTwo, CLNPName from VIEW_Payment2CLN_Chq where PVCLNChequeNo >= '" & Trim(txtChequeNoFr) & "' And PVCLNChequeNo <= '" & Trim(txtChequeNoTo) & "' order by PVCLNChequeNo"
    chq.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
    
    Do While Not chq.EOF
        RecordFound = True
        Set Record = lstChequeListing.ListItems.Add(, , chq!PVCLNChequeNo)
        
        If Len(chq!PVCLNChequeDate) Then
            Record.SubItems(1) = (chq!PVCLNChequeDate)
        End If
        If Len(chq!PVCLNChequeAmt) Then
            Record.SubItems(2) = (chq!PVCLNChequeAmt)
        End If
        'If Len(chq!ChequeNo) Then
        '    Record.SubItems(3) = (chq!ChequeNo)
        'End If
        'If Len(chq!AccountType) Then
        '    Record.SubItems(4) = (chq!AccountType)
        'End If
        RecCnt = RecCnt + 1
        lblTotalChqList = RecCnt
        chq.MoveNext
    Loop
    chq.Close
    Set chq = Nothing
    
    cmdExit.Enabled = True
    
    If RecordFound = True Then
        If CDbl(lblTotalChqList) = CDbl(lblTotalchqSel) Then
            cmdPrint.Enabled = True
            If chkPrintAccNo.Value = 1 Then
                cmdPrintACCNo.Enabled = True
                cmdPrint.Enabled = False
            Else
                cmdPrintACCNo.Enabled = False
                cmdPrint.Enabled = True
            End If
            
        Else
            cmdPrint.Enabled = False
            MsgBox "Cheque range not in series", vbOKOnly + vbCritical
        End If
    Else
        MsgBox "No record found", vbOKOnly + vbCritical
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ChequeMBMListing()

Const METHOD_NAME = "ChequeMBMListing"
On Error GoTo debugs

Dim i As Integer
Dim StrSql As String
Dim chq As New ADODB.Recordset
Dim RecFound As Boolean
Dim RecCnt As Integer
Dim tmpDD As String
Dim tmpMM As String
Dim tmpYY As String
Dim Record As ListItem
Dim RecordFound As Boolean
    
    RecordFound = False
    
    If chkPrintAccNo.Value = 0 Then
        StrSql = "select PVChequeNo, PVChequeDate, PVCLNChequeAmt, CLMTotalActual," & _
        " PVWordingOne,PVWordingTwo, PVPayto from VIEW_Payment2MBM_Chq where PVChequeNo >= '" & Trim(txtChequeNoFr) & "' And PVChequeNo <= '" & Trim(txtChequeNoTo) & "' order by PVChequeNo"
        chq.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
    Else
        StrSql = "select PVChequeNo, PVChequeDate, PVCLNChequeAmt, CLMTotalActual," & _
        " PVWordingOne,PVWordingTwo, PVPayto from VIEW_Payment2MBM_Acc where PVChequeNo >= '" & Trim(txtChequeNoFr) & "' And PVChequeNo <= '" & Trim(txtChequeNoTo) & "' order by PVChequeNo"
        chq.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
    End If
    
    Do While Not chq.EOF
        RecordFound = True
        Set Record = lstChequeListing.ListItems.Add(, , chq!PVChequeNo)
        
        If Len(chq!PVChequeDate) Then
            Record.SubItems(1) = (chq!PVChequeDate)
        End If
        If Len(chq!PVCLNChequeAmt) Then
            Record.SubItems(2) = (chq!PVCLNChequeAmt)
        End If
        If Len(chq!CLMTotalActual) Then
            Record.SubItems(6) = (chq!CLMTotalActual)
        End If

        RecCnt = RecCnt + 1
        lblTotalChqList = RecCnt
        chq.MoveNext
    Loop
    chq.Close
    Set chq = Nothing
    
    cmdExit.Enabled = True
    
    If RecordFound = True Then
        If CDbl(lblTotalChqList) = CDbl(lblTotalchqSel) Then
            cmdPrint.Enabled = True
            If chkPrintAccNo.Value = 1 Then
                cmdPrintACCNo.Enabled = True
                cmdPrint.Enabled = False
            Else
                cmdPrintACCNo.Enabled = False
                cmdPrint.Enabled = True
            End If
            
        Else
            cmdPrint.Enabled = False
            MsgBox "Cheque range not in series", vbOKOnly + vbCritical
        End If
    Else
        MsgBox "No record found", vbOKOnly + vbCritical
    End If
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


Private Sub optNS_Click()
    lblPymtType = ""
    lblPymtType = Mid(optNS.Caption, 1, 2)
    cmdPrint.Enabled = False
End Sub

Private Sub optNSMAA_Click()
    lblPymtType = ""
    lblPymtType = Mid(optNSMAA.Caption, 1, 2)
    cmdPrint.Enabled = False
End Sub

Private Sub optSihat_Click()
    lblPymtType = ""
    lblPymtType = Mid(optSihat.Caption, 1, 2)
    cmdPrint.Enabled = False
End Sub

Private Sub txtChequeNoFr_Click()
    cmdPrint.Enabled = False
End Sub

Private Sub txtChequeNoTo_Click()
    cmdPrint.Enabled = False
End Sub


