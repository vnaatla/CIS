VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdateClaimsData 
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3540
   ScaleWidth      =   10530
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   10455
      Begin VB.TextBox txtFileName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   15
         Top             =   240
         Width           =   7695
      End
      Begin VB.CommandButton cmdFileSearch 
         Caption         =   ">"
         Height          =   255
         Left            =   8490
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpdateBordxDate 
         Caption         =   "Bordx Date"
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdVerifyBordxDate 
         Caption         =   "Verify Bordx Date"
         Height          =   375
         Left            =   7200
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "FFS"
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "Verify FFS"
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   4800
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdReturnDate 
         Caption         =   "Copay"
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdVerifyReturn 
         Caption         =   "Verify Copay"
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdVerifySend 
         Caption         =   "Verify Bordx Amt"
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdUploadSend 
         Caption         =   "Bordx Amt"
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   6480
         TabIndex        =   3
         Top             =   1080
         Width           =   615
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSComctlLib.ListView lstSendDate 
      Height          =   4815
      Left            =   360
      TabIndex        =   0
      Top             =   4200
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8493
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Line No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bordx Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "SendDate"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstErrorList 
      Height          =   4815
      Left            =   360
      TabIndex        =   1
      Top             =   3120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8493
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Line No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Approval Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Approval Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Bordx Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ReturnDate"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Upload Send/Approval/Return"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10605
   End
End
Attribute VB_Name = "frmUpdateClaimsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xl As New Excel.Application
Dim xlsheet As Excel.Worksheet
Dim xlwbook As Excel.Workbook
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim tmpPathName As String

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFileSearch_Click()
 

  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  CommonDialog1.InitDir = "C:\"
  CommonDialog1.Flags = cdlOFNHideReadOnly
  CommonDialog1.Filter = "All Files (*.*)|*.*|Excel Files (*.xsl)|*.xsl"
  CommonDialog1.ShowOpen
  txtFileName = CommonDialog1.FileTitle
  Exit Sub
  
ErrHandler:
  Exit Sub
End Sub

Private Sub CmdPrint_Click()
        Screen.MousePointer = vbHourglass
        
        If lstSendDate.Visible = True Then
            If lstSendDate.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstSendDate)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Else
            If lstErrorList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstErrorList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        
        End If
        Screen.MousePointer = vbDefault
End Sub

Private Sub cmdReturnDate_Click()
Dim BordxNo As String
Dim SendDate As String
Dim AppAmt As Double
Dim commitTrans As Boolean
Dim Val1 As Integer
Dim strsqlupdate As String
Dim Number As Integer
Dim Record As ListItem
Dim RecordSaved
    Screen.MousePointer = vbHourglass

    strsqlupdate = ""
    commitTrans = False
    Val1 = 2
    lstSendDate.Visible = True
    lstSendDate.ListItems.Clear
    lstErrorList.Visible = False
    lstErrorList.ListItems.Clear
 '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If txtFileName <> "" Then
        RecordSaved = MsgBox("Do you want to Upload the records?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then

            Set xlwbook = xl.Workbooks.Open(CommonDialog1.Filename)
            Set xlsheet = xlwbook.Sheets.Item(1)
        
            
                MediConnCISWS.beginTrans
                On Err GoTo CisErrorSave
                Number = 1
    
                Do Until Val1 = 60000
                    If xlsheet.Cells(Val1, 1) <> "" Then
                        BordxNo = xlsheet.Cells(Val1, 2)
                        AppAmt = xlsheet.Cells(Val1, 2)
                        'SendDate = Format(xlsheet.Cells(Val1, 20), "dd/mm/yyyy")
                        Call UpdateCopay(BordxNo, AppAmt)
                        
                        'Set Record = lstSendDate.ListItems.Add(, , Number)
                        
                       ' If Len(BordxNo) Then
                       '     Record.SubItems(1) = BordxNo
                       ' End If
                       '
                       ' If Len(AppAmt) Then
                       '     Record.SubItems(2) = AppAmt
                       ' End If
                        
                       ' If Len(SendDate) Then
                       '     Record.SubItems(3) = SendDate
                       ' End If
                        
                        
                    Else
                        xl.DisplayAlerts = True
                        xl.Interactive = True
                        Set xlsheet = Nothing
                        Set xlwbook = Nothing
                        xl.Quit
                        Set xl = Nothing
                        commitTrans = True
                        MediConnCISWS.commitTrans
                        MsgBox "Records have been uploaded!", vbOKOnly + vbCritical, DialogBoxTitle
                        MediConnCISWS.Close
                        Set MediConnCISWS = Nothing
                        
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    Val1 = Val1 + 1
                    Number = Number + 1
                Loop
        End If
    End If
    MediConnCISWS.Close
    Set MediConnCISWS = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub
CisErrorSave:
    If commitTrans = False Then
        MediConnCISWS.RollbackTrans
        MediConnCISWS.Close
        Set MediConnCISWS = Nothing
    End If
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdUpdateBordxDate_Click()
Dim BordxNo As String
Dim SendDate As String
Dim AppAmt As Double
Dim commitTrans As Boolean
Dim Val1 As Integer
Dim strsqlupdate As String
Dim Number As Integer
Dim Record As ListItem
Dim RecordSaved
    Screen.MousePointer = vbHourglass

    strsqlupdate = ""
    commitTrans = False
    Val1 = 2
    lstSendDate.Visible = True
    lstSendDate.ListItems.Clear
    lstErrorList.Visible = False
    lstErrorList.ListItems.Clear
 '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If txtFileName <> "" Then
        RecordSaved = MsgBox("Do you want to Upload the records?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then

            Set xlwbook = xl.Workbooks.Open(CommonDialog1.Filename)
            Set xlsheet = xlwbook.Sheets.Item(1)
        
            
                MediConnCISWS.beginTrans
                On Err GoTo CisErrorSave
                Number = 1
    
                Do Until Val1 = 60000
                    If xlsheet.Cells(Val1, 1) <> "" Then
                        BordxNo = xlsheet.Cells(Val1, 2)
                        AppAmt = xlsheet.Cells(Val1, 2)
                        'SendDate = Format(xlsheet.Cells(Val1, 20), "dd/mm/yyyy")
                        Call UpdateBordxDate(BordxNo, SendDate)
                        
                        'Set Record = lstSendDate.ListItems.Add(, , Number)
                        
                       ' If Len(BordxNo) Then
                       '     Record.SubItems(1) = BordxNo
                       ' End If
                       '
                       ' If Len(AppAmt) Then
                       '     Record.SubItems(2) = AppAmt
                       ' End If
                        
                       ' If Len(SendDate) Then
                       '     Record.SubItems(3) = SendDate
                       ' End If
                        
                        
                    Else
                        xl.DisplayAlerts = True
                        xl.Interactive = True
                        Set xlsheet = Nothing
                        Set xlwbook = Nothing
                        xl.Quit
                        Set xl = Nothing
                        commitTrans = True
                        MediConnCISWS.commitTrans
                        MsgBox "Records have been uploaded!", vbOKOnly + vbCritical, DialogBoxTitle
                        MediConnCISWS.Close
                        Set MediConnCISWS = Nothing
                        
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    Val1 = Val1 + 1
                    Number = Number + 1
                Loop
        End If
    End If
    MediConnCISWS.Close
    Set MediConnCISWS = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub
CisErrorSave:
    If commitTrans = False Then
        MediConnCISWS.RollbackTrans
        MediConnCISWS.Close
        Set MediConnCISWS = Nothing
    End If
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdUpload_Click()
Dim BordxNo As String
Dim SendDate As String
Dim AppAmt As Double
Dim commitTrans As Boolean
Dim Val1 As Integer
Dim strsqlupdate As String
Dim Number As Integer
Dim Record As ListItem
Dim RecordSaved
    Screen.MousePointer = vbHourglass

    strsqlupdate = ""
    commitTrans = False
    Val1 = 2
    lstSendDate.Visible = True
    lstSendDate.ListItems.Clear
    lstErrorList.Visible = False
    lstErrorList.ListItems.Clear
 '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If txtFileName <> "" Then
        RecordSaved = MsgBox("Do you want to Upload the records?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then

            Set xlwbook = xl.Workbooks.Open(CommonDialog1.Filename)
            Set xlsheet = xlwbook.Sheets.Item(1)
        
            
                MediConnCISWS.beginTrans
                On Err GoTo CisErrorSave
                Number = 1
    
                Do Until Val1 = 60000
                    If xlsheet.Cells(Val1, 1) <> "" Then
                        BordxNo = xlsheet.Cells(Val1, 2)
                        AppAmt = xlsheet.Cells(Val1, 2)
                        'SendDate = Format(xlsheet.Cells(Val1, 20), "dd/mm/yyyy")
                        Call UpdateFFS(BordxNo, AppAmt)
                    Else
                        xl.DisplayAlerts = True
                        xl.Interactive = True
                        Set xlsheet = Nothing
                        Set xlwbook = Nothing
                        xl.Quit
                        Set xl = Nothing
                        commitTrans = True
                        MediConnCISWS.commitTrans
                        MsgBox "Records have been uploaded!", vbOKOnly + vbCritical, DialogBoxTitle
                        MediConnCISWS.Close
                        Set MediConnCISWS = Nothing
                        
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    Val1 = Val1 + 1
                    Number = Number + 1
                Loop
        End If
    End If
    MediConnCISWS.Close
    Set MediConnCISWS = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub
CisErrorSave:
    If commitTrans = False Then
        MediConnCISWS.RollbackTrans
        MediConnCISWS.Close
        Set MediConnCISWS = Nothing
    End If
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdUploadSend_Click()
Dim BordxNo As String
Dim SendDate As String
Dim AppAmt As Double
Dim commitTrans As Boolean
Dim Val1 As Integer
Dim strsqlupdate As String
Dim Number As Integer
Dim Record As ListItem
Dim RecordSaved
    Screen.MousePointer = vbHourglass

    strsqlupdate = ""
    commitTrans = False
    Val1 = 2
    lstSendDate.Visible = True
    lstSendDate.ListItems.Clear
    lstErrorList.Visible = False
    lstErrorList.ListItems.Clear
 '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If txtFileName <> "" Then
        RecordSaved = MsgBox("Do you want to Upload the records?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then

            Set xlwbook = xl.Workbooks.Open(CommonDialog1.Filename)
            Set xlsheet = xlwbook.Sheets.Item(1)
        
            
                MediConnCISWS.beginTrans
                On Err GoTo CisErrorSave
                Number = 1
    
                Do Until Val1 = 60000
                    If xlsheet.Cells(Val1, 1) <> "" Then
                        BordxNo = xlsheet.Cells(Val1, 1)
                        AppAmt = xlsheet.Cells(Val1, 2)
                        'SendDate = Format(xlsheet.Cells(Val1, 20), "dd/mm/yyyy")
                        'Call UpdateSent(BordxNo, AppAmt)
                        
                        'Set Record = lstSendDate.ListItems.Add(, , Number)
                        
                       ' If Len(BordxNo) Then
                       '     Record.SubItems(1) = BordxNo
                       ' End If
                       '
                       ' If Len(AppAmt) Then
                       '     Record.SubItems(2) = AppAmt
                       ' End If
                        
                       ' If Len(SendDate) Then
                       '     Record.SubItems(3) = SendDate
                       ' End If
                        
                        
                    Else
                        xl.DisplayAlerts = True
                        xl.Interactive = True
                        Set xlsheet = Nothing
                        Set xlwbook = Nothing
                        xl.Quit
                        Set xl = Nothing
                        commitTrans = True
                        MediConnCISWS.commitTrans
                        MsgBox "Records have been uploaded!", vbOKOnly + vbCritical, DialogBoxTitle
                        'MediConnCISWS.Close
                        'Set MediConnCISWS = Nothing
                        
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    Val1 = Val1 + 1
                    Number = Number + 1
                Loop
        End If
    End If
    MediConnCISWS.Close
    Set MediConnCISWS = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub
CisErrorSave:
    If commitTrans = False Then
        MediConnCISWS.RollbackTrans
        MediConnCISWS.Close
        Set MediConnCISWS = Nothing
    End If
    MsgBox Err.Description
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdVerify_Click()
Dim BordxNo, AppCode, AppDate As String
Dim AppAmt As Double
Dim commitTrans As Boolean
Dim Val1 As Integer
Dim error As Boolean
Dim Record As ListItem
Dim lSize As Long
Dim FileExists As Integer
Dim ReturnDate As String

    Screen.MousePointer = vbHourglass
    If txtFileName <> "" Then
            
    
        lstSendDate.Visible = False
        lstErrorList.Visible = True
        lstErrorList.ListItems.Clear
        commitTrans = False
        Val1 = 2
        error = False
        cmdUpload.Enabled = False
        cmdUploadSend.Enabled = False
        cmdReturnDate.Enabled = False
        FileExists = 0
    
        lSize = -1
        lSize = FileLen(CommonDialog1.Filename)
        If lSize <= 0 Then
            FileExists = 0
        ElseIf lSize > 0 Then
            FileExists = 1
        End If
        'FileLen (dlgOpen.Filename)

        If FileExists = 1 Then
            Set xlwbook = xl.Workbooks.Open(txtFileName)
            Set xlsheet = xlwbook.Sheets.Item(1)
        
            
            'Con.beginTrans
            
            Do Until Val1 = 60000
                If xlsheet.Cells(Val1, 1) <> "" Then
                    error = False
                    BordxNo = xlsheet.Cells(Val1, 1)
                    AppCode = xlsheet.Cells(Val1, 8)
                    AppDate = xlsheet.Cells(Val1, 9)
                    AppAmt = xlsheet.Cells(Val1, 10)
                    'ReturnDate = xlsheet.Cells(Val1, 12)
                    commitTrans = True
                    
                    If Len(BordxNo) <> 8 And Len(BordxNo) <> 10 Then
                        error = True
                    End If
                    
                    If Mid(BordxNo, 1, 1) <> "N" And Mid(BordxNo, 1, 1) <> "S" Then
                        error = True
                    End If
                    
                    If Len(AppCode) <> 12 Then
                        error = True
                    End If
                    
                    If IsDate(AppDate) = False Then
                        error = True
                    End If
                    
                    If IsNumeric(AppAmt) = False Then
                        error = True
                    End If
                    
                    'If IsDate(ReturnDate) = False Then
                    '    error = True
                    'End If
                    
                    If error = True Then
                        Set Record = lstErrorList.ListItems.Add(, , Val1)
                        
                        If Len(BordxNo) Then
                            Record.SubItems(1) = BordxNo
                        End If
                        
                        If Len(AppCode) Then
                            Record.SubItems(2) = AppCode
                        End If
                        
                        If Len(AppDate) Then
                            Record.SubItems(3) = AppDate
                        End If
                        
                        If Len(AppAmt) Then
                            Record.SubItems(4) = AppAmt
                        End If
                        
                        If Len(ReturnDate) Then
                            Record.SubItems(5) = ReturnDate
                        End If
                        
                    '    cmdUpload.Enabled = False
                    'Else
                    '    cmdUpload.Enabled = True
                    End If
                        
                '    Con.Execute "Insert into TABLE values (2, " & xlsheet.Cells(Val1, 1) & ", " _
                '  & " " & xlsheet.Cells(Val1, 2) & ", '" & Format(MyDate, "dd-mmm-yyyy") & "', " _
                '  & " " & xlsheet.Cells(Val1, 9) & ", " & flag1 & ", '" & Trim(remarks) & "')"
                Else
                    xl.DisplayAlerts = True
                    xl.Interactive = True
                    Set xlsheet = Nothing
                    Set xlwbook = Nothing
                    xl.Quit
                    Set xl = Nothing
                    If lstErrorList.ListItems.Count = 0 Then
                        MsgBox "File verified successfully!", vbOKOnly + vbCritical, DialogBoxTitle
                        cmdUpload.Enabled = True
                    Else
                        MsgBox "Error encoutered. Please check the file!", vbOKOnly + vbCritical, DialogBoxTitle
                        cmdUpload.Enabled = False
                    End If
                    
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                Val1 = Val1 + 1
            Loop
        Else
            MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
        End If
     
    Else
        MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
     
     'Con.CommitTrans
    Screen.MousePointer = vbDefault
            
End Sub

Private Sub cmdVerifyReturn_Click()
Dim BordxNo, AppCode, AppDate As String
Dim AppAmt As Double
Dim commitTrans As Boolean
Dim Val1 As Integer
Dim error As Boolean
Dim Record As ListItem
Dim lSize As Long
Dim FileExists As Integer
Dim ReturnDate As String

    Screen.MousePointer = vbHourglass
    If txtFileName <> "" Then
        lstErrorList.ListItems.Clear
        lstSendDate.ListItems.Clear
        lstErrorList.Visible = True
        lstSendDate.Visible = False
        
        commitTrans = False
        Val1 = 2
        error = False
        cmdUpload.Enabled = False
        cmdUploadSend.Enabled = False
        cmdReturnDate.Enabled = False
    
        FileExists = 0
    
        lSize = -1
        lSize = FileLen(CommonDialog1.Filename)
        If lSize <= 0 Then
            FileExists = 0
        ElseIf lSize > 0 Then
            FileExists = 1
        End If
        'FileLen (dlgOpen.Filename)
    
        If FileExists = 1 Then
            Set xlwbook = xl.Workbooks.Open(txtFileName)
            Set xlsheet = xlwbook.Sheets.Item(1)
        
            
            'Con.beginTrans
            
            Do Until Val1 = 60000
                If xlsheet.Cells(Val1, 1) <> "" Then
                    error = False
                    BordxNo = xlsheet.Cells(Val1, 2)
                    AppCode = xlsheet.Cells(Val1, 18)
                    AppDate = xlsheet.Cells(Val1, 19)
                    AppAmt = xlsheet.Cells(Val1, 3)
                    ReturnDate = xlsheet.Cells(Val1, 20)
                    commitTrans = True
                    
                    If Len(BordxNo) <> 8 And Len(BordxNo) <> 10 Then
                        error = True
                    End If
                    
                    If Mid(BordxNo, 1, 1) <> "N" And Mid(BordxNo, 1, 1) <> "S" Then
                        error = True
                    End If
                    
                    'If Len(AppCode) <> 12 Then
                    '    error = True
                    'End If
                    
                    'If IsDate(AppDate) = False Then
                    '    error = True
                    'End If
                    
                    If IsNumeric(AppAmt) = False Then
                        error = True
                    End If
                    
                    If IsDate(ReturnDate) = False Then
                        error = True
                    End If
                    
                    If error = True Then
                        Set Record = lstErrorList.ListItems.Add(, , Val1)
                        
                        If Len(BordxNo) Then
                            Record.SubItems(1) = BordxNo
                        End If
                        
                        If Len(AppCode) Then
                            Record.SubItems(2) = AppCode
                        End If
                        
                        If Len(AppDate) Then
                            Record.SubItems(3) = AppDate
                        End If
                        
                        If Len(AppAmt) Then
                            Record.SubItems(4) = AppAmt
                        End If
                        
                        If Len(ReturnDate) Then
                            Record.SubItems(5) = ReturnDate
                        End If
                        
                    '    cmdUpload.Enabled = False
                    'Else
                    '    cmdUpload.Enabled = True
                    End If
                        
                '    Con.Execute "Insert into TABLE values (2, " & xlsheet.Cells(Val1, 1) & ", " _
                '  & " " & xlsheet.Cells(Val1, 2) & ", '" & Format(MyDate, "dd-mmm-yyyy") & "', " _
                '  & " " & xlsheet.Cells(Val1, 9) & ", " & flag1 & ", '" & Trim(remarks) & "')"
                Else
                    xl.DisplayAlerts = True
                    xl.Interactive = True
                    Set xlsheet = Nothing
                    Set xlwbook = Nothing
                    xl.Quit
                    Set xl = Nothing
                    If lstErrorList.ListItems.Count = 0 Then
                        MsgBox "File verified successfully!", vbOKOnly + vbCritical, DialogBoxTitle
                        cmdReturnDate.Enabled = True
                    Else
                        MsgBox "Error encoutered. Please check the file!", vbOKOnly + vbCritical, DialogBoxTitle
                        cmdReturnDate.Enabled = False
                    End If
                    
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                Val1 = Val1 + 1
            Loop
        Else
            MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    Else
        MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
     
     'Con.CommitTrans
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdVerifySend_Click()
Dim BordxNo As String
Dim AppAmt As Double
Dim commitTrans As Boolean
Dim Val1 As Integer
Dim error As Boolean
Dim Record As ListItem
Dim lSize As Long
Dim FileExists As Integer
Dim SendDate As String
    
    Screen.MousePointer = vbHourglass
    If txtFileName <> "" Then
            lstSendDate.Visible = True
            lstErrorList.Visible = False
            lstSendDate.ListItems.Clear
            commitTrans = False
            Val1 = 2
            error = False
            cmdUpload.Enabled = False
            FileExists = 0
        
            lSize = -1
            lSize = FileLen(CommonDialog1.Filename)
            If lSize <= 0 Then
                FileExists = 0
            ElseIf lSize > 0 Then
                FileExists = 1
            End If
            'FileLen (dlgOpen.Filename)
            txtFileName = CommonDialog1.Filename
            If FileExists = 1 Then
                Set xlwbook = xl.Workbooks.Open(txtFileName)
                Set xlsheet = xlwbook.Sheets.Item(1)
            
                
                'Con.beginTrans
                
                Do Until Val1 = 60000
                    If xlsheet.Cells(Val1, 1) <> "" Then
                        error = False
                        BordxNo = xlsheet.Cells(Val1, 2)
                        AppAmt = xlsheet.Cells(Val1, 3)
                        SendDate = xlsheet.Cells(Val1, 20)
                        commitTrans = True
                        
                        If Len(BordxNo) <> 8 And Len(BordxNo) <> 10 Then
                            error = True
                        End If
                        
                        If Mid(BordxNo, 1, 1) <> "N" And Mid(BordxNo, 1, 1) <> "S" Then
                            error = True
                        End If
                        
                        
                        If IsNumeric(AppAmt) = False Then
                            error = True
                        End If
                        
                        If IsDate(SendDate) = False Then
                            error = True
                        End If
                        
                        If error = True Then
                            Set Record = lstErrorList.ListItems.Add(, , Val1)
                            
                            If Len(BordxNo) Then
                                Record.SubItems(1) = BordxNo
                            End If
                            
                            If Len(AppAmt) Then
                                Record.SubItems(2) = AppAmt
                            End If
                            
                            If Len(SendDate) Then
                                Record.SubItems(3) = SendDate
                            End If
                            
                        End If
                            
                    Else
                        xl.DisplayAlerts = True
                        xl.Interactive = True
                        Set xlsheet = Nothing
                        Set xlwbook = Nothing
                        xl.Quit
                        Set xl = Nothing
                        If lstSendDate.ListItems.Count = 0 Then
                            MsgBox "File verified successfully!", vbOKOnly + vbCritical, DialogBoxTitle
                            cmdUploadSend.Enabled = True
                        Else
                            MsgBox "Error encoutered. Please check the file!", vbOKOnly + vbCritical, DialogBoxTitle
                            cmdUploadSend.Enabled = False
                        End If
                        
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    Val1 = Val1 + 1
                Loop
            Else
                MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
            End If
        Else
            MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
        End If
            
             'Con.CommitTrans
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
   ' tmpPathName = "\\MARS\BordxReturned\"
    tmpPathName = "\\Svr-Doc\BordxReturned\"
End Sub

Private Sub UpdateSent(tBordxNo As String, AppAmt As Double)
Dim StrSql As String
Dim strsqlREM As String
Dim rst As New ADODB.Recordset
Dim strsql2 As String
Dim rst2 As New ADODB.Recordset
Dim tmpRemarks1 As String
Dim tmpRemarks2 As String
Dim tmpRemarks3 As String
Dim tmpRemarks4 As String
Dim tmpRemarks5 As String
Dim strsqlWeb As String
Dim WEB As New ADODB.Recordset

    
    StrSql = ""
    StrSql = " update Claim set " & _
            " CLMBordxAmt = '" & AppAmt & "', " & _
            " ClaimLastUpdateDate = '" & Date & "', " & _
            " ClaimLastUpdateUser = '" & Logon & "' " & _
            " Where CASNumber = '" & Trim(tBordxNo) & "'"
    MediConnCISWS.Execute StrSql
                
End Sub

Private Sub UpdateFFS(tBordxNo As String, ProcFee As Double)
Dim StrSql As String
Dim strsqlREM As String
Dim rst As New ADODB.Recordset
Dim strsql2 As String
Dim rst2 As New ADODB.Recordset
Dim tmpRemarks1 As String
Dim tmpRemarks2 As String
Dim tmpRemarks3 As String
Dim tmpRemarks4 As String
Dim tmpRemarks5 As String
Dim strsqlWeb As String
Dim WEB As New ADODB.Recordset

    
    StrSql = ""
    StrSql = " update Claim set " & _
            " CLMMCOProcFee = '" & ProcFee & "', " & _
            " ClaimLastUpdateDate = '" & Date & "', " & _
            " ClaimLastUpdateUser = '" & Logon & "' " & _
            " Where CASNumber = '" & Trim(tBordxNo) & "'"
    MediConnCISWS.Execute StrSql
                
End Sub

Private Sub UpdateCopay(tBordxNo As String, copay As Double)
Dim StrSql As String
Dim strsqlREM As String
Dim rst As New ADODB.Recordset
Dim strsql2 As String
Dim rst2 As New ADODB.Recordset
Dim tmpRemarks1 As String
Dim tmpRemarks2 As String
Dim tmpRemarks3 As String
Dim tmpRemarks4 As String
Dim tmpRemarks5 As String
Dim strsqlWeb As String
Dim WEB As New ADODB.Recordset

    
    StrSql = ""
    StrSql = " update Claim set " & _
            " CLMCoPayColl = '" & copay & "', " & _
            " ClaimLastUpdateDate = '" & Date & "', " & _
            " ClaimLastUpdateUser = '" & Logon & "' " & _
            " Where CASNumber = '" & Trim(tBordxNo) & "'"
    MediConnCISWS.Execute StrSql
                
End Sub

Private Sub UpdateBordxDate(tBordxNo As String, SentDate As String)
Dim StrSql As String
Dim strsqlREM As String
Dim rst As New ADODB.Recordset
Dim strsql2 As String
Dim rst2 As New ADODB.Recordset
Dim tmpRemarks1 As String
Dim tmpRemarks2 As String
Dim tmpRemarks3 As String
Dim tmpRemarks4 As String
Dim tmpRemarks5 As String
Dim strsqlWeb As String
Dim WEB As New ADODB.Recordset
Dim tmpCLMBordxMonth As String
Dim tmpCLMBordxYear As String

    tmpCLMBordxMonth = Month(SentDate)
    tmpCLMBordxYear = Year(SentDate)
    
    StrSql = ""
    StrSql = " update Claim set " & _
            " CLMBordxMonth = '" & tmpCLMBordxMonth & "', " & _
             " CLMBordxYear = '" & tmpCLMBordxYear & "', " & _
            " ClaimLastUpdateDate = '" & Date & "', " & _
            " ClaimLastUpdateUser = '" & Logon & "' " & _
            " Where BordxRefNo = '" & Trim(tBordxNo) & "'"
    MediConnCISWS.Execute StrSql
                
    StrSql = ""
    StrSql = " update Bordx set " & _
            " BordxDate = '" & SentDate & "' " & _
            " Where BordxRefNo = '" & Trim(tBordxNo) & "'"
    MediConnCISWS.Execute StrSql
    
End Sub
Private Function UpdateReturnPayor(BordxNo As String, ReturnDate As String)
Dim StrSql As String
Dim strsql2 As String
'Dim FRPayor As New ADODB.Recordset
Dim RecDocDate As String
Dim ii As Integer
Dim TempBodrxNo As String
Dim TempCASNumber As String
Dim TempCLMVersion As String
Dim CLM As New ADODB.Recordset
Dim strsqlRem1 As String
Dim strsqlRem2 As String
Dim strsqlRem3 As String
Dim REM1 As New ADODB.Recordset
Dim REM2 As New ADODB.Recordset
Dim REM3 As New ADODB.Recordset
Dim RemLen1 As Integer
Dim RemLen2 As Integer
Dim RemLen3 As Integer
Dim RemLen4 As Integer
Dim RemLen5 As Integer
Dim RemLen6 As Integer
Dim RemLen7 As Integer
Dim Remarks1 As String
Dim Remarks2 As String
Dim Remarks3 As String
Dim Remarks4 As String
Dim Remarks5 As String
Dim Remarks6 As String
Dim Remarks7 As String
Dim AddRemark As String
Dim strsqlREM As String

    ' update bordx table
    RecDocDate = Format(ReturnDate, "YYYY/MM/DD")
       
        RemLen1 = 0
        RemLen2 = 0
        RemLen3 = 0
        RemLen4 = 0
        RemLen5 = 0
        RemLen6 = 0
        RemLen7 = 0
        
        Remarks1 = ""
        Remarks2 = ""
        Remarks3 = ""
        Remarks4 = ""
        Remarks5 = ""
        Remarks6 = ""
        Remarks7 = ""
    
        AddRemark = vbNewLine & Date & "   Received Returned Bordx  -  " & BordxNo & " on " & ReturnDate & "                           " & Logon & " "
        AddRemark = ChkStr(Replace(AddRemark, Chr(13), "~"))
        
        strsql2 = ""
        strsql2 = " update Claim set " & _
                    " BordxPAYRecDocDate = '" & Format(RecDocDate, "yyyy/mm/dd") & "', " & _
                    " CLMAppStatus = 'A' " & _
                    " Where BordxRefNo = '" & Trim(BordxNo) & "'"
        medixmasterconnWS.Execute strsql2
        
        strsql2 = ""
        strsql2 = "Select CASNumber from Claim Where BordxRefNo = '" & Trim(BordxNo) & "'"
        CLM.Open strsql2, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        If CLM.RecordCount Then
            TempCASNumber = CLM!CasNumber
        End If
        CLM.Close
        Set CLM = Nothing
        
        If Len(TempCASNumber) Then
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

                        strsqlRem1 = "Select * from CASRemarks where CASNumber = '" & Trim(TempCASNumber) & "'"
                        REM1.Open strsqlRem1, medixmasterconn, adOpenKeyset, adLockOptimistic
                            If REM1.RecordCount > 0 Then
                                If REM1!CASORemarks <> "" Then
                                    RemLen1 = Len(REM1!CASORemarks)
                                    Remarks1 = REM1!CASORemarks & AddRemark
                                    Remarks1 = ChkStr(Replace(Remarks1, "'", "''"))
                                    If REM1!CASORemarksTwo <> "" Then
                                        RemLen2 = Len(REM1!CASORemarksTwo)
                                        Remarks2 = REM1!CASORemarksTwo & AddRemark
                                        Remarks2 = ChkStr(Replace(Remarks2, "'", "''"))
                                    End If
                                Else
                                    Remarks1 = AddRemark
                                End If
                            Else
                                Remarks1 = AddRemark
                            End If
                        REM1.Close
                        Set REM1 = Nothing
                        
                        strsqlRem2 = "Select * from CASRemarksTwo where CASNumber= '" & Trim(TempCASNumber) & "'"
                        REM2.Open strsqlRem2, medixmasterconn, adOpenKeyset, adLockOptimistic
                            If REM2.RecordCount Then
                                If REM2!CASORemarks <> "" Then
                                    RemLen3 = Len(REM2!CASORemarks)
                                    Remarks3 = REM2!CASORemarks & AddRemark
                                    Remarks3 = ChkStr(Replace(Remarks3, "'", "''"))
                                    RemLen4 = Len(REM2!CASORemarksTwo)
                                    Remarks4 = REM2!CASORemarksTwo & AddRemark
                                    Remarks4 = ChkStr(Replace(Remarks4, "'", "''"))
                                End If
                            End If
                        REM2.Close
                        Set REM2 = Nothing
                        'REM3.Close
                        'Set REM3 = Nothing
                        strsqlRem3 = ""
                        strsqlRem3 = "Select * from CASRemarkThree where CASNumber= '" & Trim(TempCASNumber) & "'"
                        REM3.Open strsqlRem3, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If REM3.RecordCount Then
                            RemLen5 = Len(REM3!CasRemarks)
                            Remarks5 = REM3!CasRemarks & AddRemark
                            Remarks5 = ChkStr(Replace(Remarks5, "'", "''"))
                            RemLen6 = Len(REM3!CASRemarksFour)
                            Remarks6 = REM3!CASRemarksFour & AddRemark
                            Remarks6 = ChkStr(Replace(Remarks6, "'", "''"))
                            RemLen7 = Len(REM3!CASRemarksFive)
                            Remarks7 = REM3!CASRemarksFive & AddRemark
                            Remarks7 = ChkStr(Replace(Remarks7, "'", "''"))
                        Else
                            RemLen5 = 0
                            Remarks5 = ""
                            RemLen6 = 0
                            Remarks6 = ""
                            RemLen7 = 0
                            Remarks7 = ""
                        End If
                        REM3.Close
                        Set REM3 = Nothing
                        
                        strsqlREM = ""
                        If (RemLen7 <> 0 And RemLen7 < 3700) Or RemLen6 > 3700 Then
                            strsqlREM = " update CASRemarkThree set " & _
                                     " CASRemarksFive = '" & Remarks7 & "'" & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"
                        ElseIf (RemLen6 <> 0 And RemLen6 < 3700) Or RemLen5 > 3700 Then
                            strsqlREM = " update CASRemarkThree set " & _
                                     " CASRemarksFour = '" & Remarks6 & "'" & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"
                        ElseIf (RemLen5 <> 0 And RemLen5 < 3700) Or RemLen4 > 1800 Then
                            strsqlREM = " update CASRemarkThree set " & _
                                     " CasRemarks = '" & Remarks5 & "'" & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"
                        ElseIf (RemLen4 <> 0 And RemLen4 < 1800) Or RemLen3 > 1800 Then
                            strsqlREM = " update CASRemarksTwo set " & _
                                     " CASORemarksTwo = '" & Remarks4 & "'" & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"
                        ElseIf (RemLen3 <> 0 And RemLen3 < 1800) Or RemLen2 > 1800 Then
                            strsqlREM = " update CASRemarksTwo set " & _
                                     " CASORemarks = '" & Remarks3 & "' " & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"
                        ElseIf (RemLen2 <> 0 And RemLen2 < 1800) Or RemLen1 > 1800 Then
                            strsqlREM = " update CasRemarks set " & _
                                     " CASORemarksTwo = '" & Remarks2 & "'" & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"
                        ElseIf RemLen1 <> 0 And RemLen1 < 1800 Then
                            strsqlREM = " update CasRemarks set " & _
                                     " CASORemarks = '" & Remarks1 & "'" & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"
                        Else
                            strsqlREM = " update CasRemarks set " & _
                                     " CASORemarks = '" & Remarks1 & "'" & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"

                        End If
                    medixmasterconn.Execute strsqlREM
             '       medixmasterconn.Close
             '       Set medixmasterconn = Nothing
                End If
            
    
    
End Function








