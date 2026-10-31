VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMaintenanceBenefitsCln 
   Caption         =   "Clinic Benefits"
   ClientHeight    =   7368
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7368
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11805
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Clinic Benefits Maintenance"
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
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   3285
      End
   End
   Begin VB.Frame Frame10 
      Height          =   705
      Left            =   360
      TabIndex        =   12
      Top             =   6600
      Width           =   11055
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   3240
         TabIndex        =   6
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   2400
         TabIndex        =   5
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Default         =   -1  'True
         Height          =   315
         Left            =   7080
         TabIndex        =   7
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   8760
         TabIndex        =   9
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   7920
         TabIndex        =   8
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   285
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   11055
      Begin VB.ComboBox cboCardType 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   1545
      End
      Begin MSMask.MaskEdBox txtTEffdate 
         Height          =   315
         Left            =   6480
         TabIndex        =   2
         Top             =   240
         Width           =   1545
         _ExtentX        =   2731
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         Caption         =   "BNF Card Type *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Eff Date           *"
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
         Left            =   4800
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FGrid 
      Height          =   5115
      Left            =   240
      TabIndex        =   15
      Top             =   1455
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   9017
      _Version        =   393216
      Rows            =   40
      Cols            =   3
      FixedCols       =   2
      AllowUserResizing=   3
   End
   Begin VB.Label txtTotalRecord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "* - Mandatory Fields"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label Label32 
      Caption         =   "TOTAL RECORDS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "Bold - Searchable "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   7440
      Width           =   1695
   End
End
Attribute VB_Name = "frmMaintenanceBenefitsCln"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recordInBNF As Integer
Public CLNBNFCardType As String

Private Sub Form_Activate()
    Call ClinicBNFcombolisting
End Sub

Private Sub cmdAdd_Click()
    
    Call ClearForm(Me)
    Call ClearGrid
    Call EntriesEnable(True)
    CmdSave.Enabled = True
    CmdAdd.Enabled = True
        
    cboCardType.SetFocus
    Call ClinicBNFcombolisting
        
End Sub
Private Sub Form_Load()
Const METHOD_NAME = "Form_Load"
On Error GoTo debugs

Call ActionLogUpdate(12, "frmMaintenanceBenefitsCln")


debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdprint_Click()
    Screen.MousePointer = vbHourglass
        Call PrintClinicBNF
        CmdAdd.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub CmdRefresh_Click()
    Call ClinicBNFcombolisting
End Sub

Private Sub CmdSave_Click()
    Screen.MousePointer = vbHourglass
        Call SaveClinicBNF
        CmdAdd.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub cmdSearch_Click()
    Screen.MousePointer = vbHourglass
        Call SearchClinicBNF
    Screen.MousePointer = Default
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    Call ClearGrid
    CmdAdd.Enabled = True
    txtTotalRecord = 0
    Call EntriesEnable(True)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


'************************Start EntriesEnable*********************************
Private Sub EntriesEnable(status As Boolean)
    CmdSearch.Enabled = status
End Sub
'************************End EntriesEnable*********************************

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub

Private Sub FGrid_Click()
Dim temp, tempCol, tempRow, tempTitle As String
Dim Calculation, value1 As Double
Dim value2, value3, value4 As Double
    
    tempCol = FGrid.Col
    tempRow = FGrid.Row
    FGrid.Row = 0
    tempTitle = FGrid.Text
    FGrid.Row = tempRow
    FGrid.Col = 1
    temp = InputBox(tempTitle, FGrid.Text)
    FGrid.Col = tempCol
    If temp <> "" Then
        FGrid = temp
        
        If IsNumeric(FGrid.Text) = True Then
            Value = FGrid.Text
        Else
            FGrid.Text = ""
        End If
        
        If tempCol = 2 Then
            FGrid.Col = 2
            If IsNumeric(FGrid.Text) = True Then
                value1 = FGrid.Text
            Else
                FGrid.Text = ""
            End If
        End If
    End If
End Sub

Private Sub SaveClinicBNF()
Const METHOD_NAME = "SaveClinicBNF"
On Error GoTo debugs

On Error GoTo UpdateErr
Dim BNF As New ADODB.Recordset
Dim SaveBNF As New ADODB.Recordset
Dim RecordSaved
Dim strsql As String
Dim currentRecord As Integer
Dim beginTrans As Boolean
Dim Pervisit As String
Dim Percentage As String
Dim Calculation As String

Calculation = "0"
Pervisit = "0"
Percentage = "0"

    If Len(Trim(cboCardType)) = 0 Then
        MsgBox "Please Enter Plan Code!", vbOKOnly + vbCritical, DialogBoxTitle
        cboCardType.SetFocus
    ElseIf txtTEffdate = "__/__/____" Or IsDate(txtTEffdate) = False Then
        MsgBox "Please Enter (Valid )effectice Date!", vbOKOnly + vbCritical, DialogBoxTitle
        txtTEffdate.SetFocus
    Else

        strsql = " Select * From STDClinicBNFSpe where CLNBNFType = '" & RemStr(cboCardType) & "'" & _
                 " AND CLNBNFEffDate ='" & Format(txtTEffdate, "mm/dd/yyyy") & "'"
    
        BNF.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic

        If BNF.RecordCount > 0 Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            BNF.Close
            cboCardType.SetFocus
            cboCardType.SelStart = 0
            cboCardType.SelLength = Len(cboTHLTCode)
            Set BNF = Nothing
            Exit Sub
        Else
            RecordSaved = MsgBox("Do you want to save this record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                For currentRecord = 1 To recordInBNF - 1 Step 1
                
                SaveBNF.Open " Select * from STDClinicBNFSpe ", MediConnCISDB, adOpenKeyset, adLockOptimistic
                SaveBNF.AddNew
               
                    With SaveBNF
                        FGrid.Row = currentRecord
                        FGrid.Col = 2
                        Calculation = FGrid.Text
                    
                        !CLNBNFType = Trim(cboCardType)
                        !CLNBNFEffDate = txtTEffdate
                        !CLNBNFLastUpdateUser = Logon
                        !CLNBNFLastUpdateDate = Date
                        
                        FGrid.Col = 0
                            !BNFCode = FGrid.Text
                        FGrid.Col = 2
                        If Trim(FGrid.Text) <> "" Then
                            !CLNBNFRMPerVisit = ChkStr(FGrid.Text)
                        End If
                        SaveBNF.update
                        SaveBNF.Close
                        Set SaveBNF = Nothing
            
                End With
                'SaveBNF.Close
                'Set SaveBNF = Nothing
            Next
            
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            
       End If
       
   End If
End If
'End If

Exit Sub
UpdateErr:
    
  MsgBox Err.Description
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If


End Sub

Private Sub ClinicBNFcombolisting()
Const METHOD_NAME = "ClinicBNFcombolisting"
On Error GoTo debugs
Dim subCat As New ADODB.Recordset
Dim Cnt As Integer

    FGrid.Row = 0
    FGrid.ColWidth(0) = 1000
    FGrid.ColWidth(1) = 3500
    FGrid.ColWidth(2) = 2000
    FGrid.Col = 0
    FGrid = "Cat Code"
    FGrid.Col = 1
    FGrid = "Cat Name"
    FGrid.Col = 2
    FGrid = "RM Payable Fee"
    
    subCat.Open "select BNFCATCode, BNFCATDescription from BNFCategoryNew ", MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    cboCardType.Clear
        
    cboCardType.AddItem "Classic"
    cboCardType.AddItem "Silver"
    cboCardType.AddItem "Gold"
    
    
    Cnt = 1
    FGrid.Rows = subCat.RecordCount + 1
    recordInBNF = subCat.RecordCount + 1
    Do Until subCat.EOF
         FGrid.Col = 0
         FGrid.Row = Cnt
         FGrid.Text = subCat!BNFCatCode
    
         FGrid.Col = 1
         FGrid.Row = Cnt
         FGrid.Text = subCat!BNFCATDescription
         Cnt = Cnt + 1
         subCat.MoveNext
    Loop
        
    
    subCat.Close
    Set subCat = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub SearchClinicBNF()
Const METHOD_NAME = "SearchClinicBNF"
On Error GoTo debugs
Dim strAppend As String
Dim strAppend2 As String
   
    strAppend = ""
    strAppend2 = ""
    
    If Len(Trim(cboCardType)) = 0 Then
        MsgBox "Please key in BNF Card Type!", vbCritical
    ElseIf txtTEffdate = "__/__/____" Then
        MsgBox "Please key in Effective Date!", vbCritical
    ElseIf IsDate(txtTEffdate) = False Then
        MsgBox "Please key in VALID Effective Date!", vbCritical
    Else
        strAppend = strAppend + " And CLNBNFType = '" & Trim(cboCardType) & "'"
        strAppend = strAppend + " And CLNBNFEffDate = '" & Format(txtTEffdate, "mm/dd/yyyy") & "'"
        strAppend2 = strAppend2 + " And CLNBNFEffDate = '" & Format(txtTEffdate, "mm/dd/yyyy") & "'"
               
        CLNBNFCardType = strAppend
        Call ClinicBNFlisting(strAppend)
        'Call ClinicBNFlisting2(strAppend2)
        
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ClinicBNFlisting(Optional strAppend As String)
Const METHOD_NAME = "ClinicBNFlisting"
On Error GoTo debugs

Dim BNF As New ADODB.Recordset
Dim strsql As String
Dim counter As Integer
    
    strsql = "Select * from STDClinicBNFSpe where 0=0 " & strAppend
    strsql = strsql & " order by BNFCode "
    BNF.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    Call ClearGrid
    If BNF.RecordCount = 0 Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until BNF.EOF
         With BNF
            For counter = 1 To recordInBNF - 1 Step 1
             FGrid.Row = counter
             FGrid.Col = 0
                If FGrid.Text = !BNFCode Then
                    FGrid.Col = 2
                    If !CLNBNFRMPerVisit <> "" Then
                        FGrid.Text = !CLNBNFRMPerVisit
                    End If
                End If
            Next
            .MoveNext
         End With
    Loop
    End If
    BNF.Close
    Set BNF = Nothing

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub ClinicBNFlisting2(Optional strAppend2 As String)
Const METHOD_NAME = "ClinicBNFlisting2"
On Error GoTo debugs
Dim BNF As New ADODB.Recordset
Dim strsql As String
Dim counter As Integer
    
    strsql = "Select * from STDClinicBNF where 0=0 " & strAppend2
    strsql = strsql & " order by BNFCode "
    BNF.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If BNF.RecordCount = 0 Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until BNF.EOF
         With BNF
            For counter = 1 To recordInBNF - 1 Step 1
             FGrid.Row = counter
             FGrid.Col = 0
                If FGrid.Text = !BNFCode Then
                    FGrid.Col = 2
                    If !CLNBNFRMPerVisit <> "" Then
                        FGrid.Text = !CLNBNFRMPerVisit
                    End If
                End If
            Next
            .MoveNext
         End With
    Loop
    End If
    BNF.Close
    Set BNF = Nothing

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub PrintClinicBNF()
    RptClinicBNF.Show
End Sub

Private Sub ClearGrid()
    
    For FCol = 2 To 2
        FGrid.Col = FCol
        For Frow = 1 To 29
            FGrid.Row = Frow
            FGrid = ""
        Next
    Next
    
End Sub

Private Sub cboCardType_Change()
    If InStr(cboCardType.Text, "null") Then
       cboCardType.Enabled = False
    End If
End Sub

Private Sub cboCardType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCardType.Text, cboCardType.SelStart) + Chr(KeyAscii) + Mid(cboCardType.Text, cboCardType.SelStart + cboCardType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCardType.ListCount - 1
 
        If NewText = LCase(Left(cboCardType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCardType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCardType.Text = ValidValue
                cboCardType.SelStart = 0
                cboCardType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub txtTEffdate_LostFocus()
    If txttxtTEffdateDateVisit <> "__/__/____" Then
        If IsDate(txtTEffdate) = False Then
            MsgBox "Invalid date format!", vbCritical
        ElseIf Format(txtTEffdate, "yyyymmdd") > Format(Date, "yyyymmdd") Then
            MsgBox "Date cannot greater than system date ! ", vbCritical
            txtTEffdate.SetFocus
        End If
    End If
End Sub


