VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCheckingItems 
   Caption         =   "Checking Items"
   ClientHeight    =   7488
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9528
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7488
   ScaleWidth      =   9528
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame10 
      Height          =   705
      Left            =   240
      TabIndex        =   22
      Top             =   6840
      Width           =   11055
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   315
         Left            =   3240
         TabIndex        =   7
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   6
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   8160
         TabIndex        =   3
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   9000
         TabIndex        =   10
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Default         =   -1  'True
         Height          =   315
         Left            =   7320
         TabIndex        =   9
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   4080
         TabIndex        =   8
         Top             =   285
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   240
      TabIndex        =   16
      Top             =   720
      Width           =   11055
      Begin VB.TextBox txtCatCode 
         Height          =   285
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   2
         Top             =   960
         Width           =   4815
      End
      Begin VB.TextBox txtCatName 
         Height          =   285
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   1
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label14 
         Caption         =   "Category Name*"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   645
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Category Code *"
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
         Left            =   360
         TabIndex        =   18
         Top             =   285
         Width           =   1740
      End
      Begin VB.Label Label16 
         Caption         =   "Amount *"
         Height          =   255
         Left            =   405
         TabIndex        =   17
         Top             =   960
         Width           =   1650
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11900
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Checking Items"
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
         TabIndex        =   12
         Top             =   120
         Width           =   3285
      End
   End
   Begin MSComctlLib.ListView lstCheckItemList 
      Height          =   3900
      Left            =   240
      TabIndex        =   20
      Top             =   2760
      Width           =   11115
      _ExtentX        =   19600
      _ExtentY        =   6879
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Category Code"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Category Name"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "CHECK ITEM LIST"
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
      Left            =   240
      TabIndex        =   21
      Top             =   2520
      Width           =   1515
   End
   Begin VB.Label txtTotalRecord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6720
      TabIndex        =   15
      Top             =   7800
      Width           =   1815
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
      Left            =   5280
      TabIndex        =   14
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "Bold - Searchable        * - Mandatory Fields"
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
      Left            =   360
      TabIndex        =   13
      Top             =   7800
      Width           =   4215
   End
End
Attribute VB_Name = "frmCheckingItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim EditFlag As Boolean
Public CheckCriteria As String
Private m_blnDirection(1 To 4) As Boolean

Private Sub cmdAdd_Click()
    EditFlag = False
    Call ClearForm(Me)
    Call DisableCmdButton(Me)
    Call EntriesEnable(True)
    CmdSave.Enabled = True
    CmdAdd.Enabled = True
    txtCatCode.SetFocus
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    Call EntriesEnable(True)
    CmdSave.Enabled = True
    CmdDelete.Enabled = False
     
    If txtCatCode <> "" Then
       txtCatCode.Enabled = False
       txtCatName.Enabled = True
       txtCatName.SetFocus
       txtAmount.Enabled = True
       txtAmount.SetFocus
    End If
End Sub
Private Sub Form_Load()
Const METHOD_NAME = "Form_Load"
On Error GoTo debugs

Call ActionLogUpdate(12, "frmCheckingItems")


debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub



Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
    Call PrintCheck
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    Screen.MousePointer = Default
End Sub

Private Sub cmdSave_Click()
    Screen.MousePointer = vbHourglass
    Call SaveCheck
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    Screen.MousePointer = Default
End Sub

Private Sub cmdDelete_Click()
    Screen.MousePointer = vbHourglass
    Call DeleteCheck
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    Screen.MousePointer = Default
End Sub

Private Sub cmdSearch_Click()
    Screen.MousePointer = vbHourglass
    Call SearchCheck
    Call ClearForm(Me)
    txtCatCode.Enabled = True
    txtCatName.Enabled = True
    txtAmount.Enabled = True
    txtCatCode.SetFocus
    Screen.MousePointer = Default
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    CmdAdd.Enabled = True
    lstCheckItemList.ListItems.Clear
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    txtTotalRecord = ""
    Call EntriesEnable(True)
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub


'************************Start EntriesEnable*********************************

Private Sub EntriesEnable(status As Boolean)
     
    txtCatCode.Enabled = status
    txtCatName.Enabled = status
    txtAmount.Enabled = status
   
End Sub
'************************End EntriesEnable*********************************
    
'####### @@@ START SAVE, EDIT, ADD, DELETE CHECKING ITEMS @@@ #######'

Private Sub DeleteCheck()
Const METHOD_NAME = "DeleteCheck"
On Error GoTo debugs
Dim rsDeleteCHK As New ADODB.Recordset
Dim rsCheckCHK As New ADODB.Recordset
Dim strsql, sqlCHK As String
Dim DeleteRecord
    
EditFlag = False
       
If txtCatCode = "" Then
   MsgBox "Please Enter Category Code", vbOKOnly + vbCritical, DialogBoxTitle
ElseIf lstCheckItemList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlCHK = "Select CatCode From CheckItems where CatCode = '" & Trim(lstCheckItemList.SelectedItem.Text) & "'"
    
'   rsCheckCOP.MaxRecords = 1
'   rsCheckCOP.Open sqlCOP, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
'    If rsCheckCOP.RecordCount Then
'        MsgBox "This record cannot be deleted because it is used in Calculation Table!", vbCritical, "Error in deleting record..."
'        Call ClearForm(Me)
'        CmdDelete.Enabled = False
'        CmdEdit.Enabled = False
    
'    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
             strsql = "Delete From CheckItems Where CatCode = '" & ChkStr(lstCheckItemList.SelectedItem.Text) & "'"
             rsDeleteCHK.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call ClearForm(Me)
             Call CheckListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             txtCatCode.SetFocus
        Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

    
Private Sub SearchCheck()
Const METHOD_NAME = "SearchCheck"
On Error GoTo debugs
Dim strAppend As String
    
  Screen.MousePointer = vbHourglass
  strAppend = ""
  
  If Len(Trim(txtCatCode)) Then
    strAppend = strAppend + " And CatCode = '" & Trim(Mid(txtCatCode, 1, 3)) & "'"
  'HLTCode like '%" & Mid(ChkStr(txtHLTcode), 1, 2) & "%'"
  End If
    
  CheckCriteria = strAppend
  Call CheckListing(strAppend)
  Screen.MousePointer = Default
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub PrintCheck()
    RptCheckList.Show
End Sub

Private Sub SaveCheck()
Const METHOD_NAME = "SaveCheck"
On Error GoTo debugs
   On Error Resume Next
   Dim RecordSaved
   Dim CHK As New ADODB.Recordset
   Dim SaveCHK As New ADODB.Recordset
   Dim strsql As String
   
    If txtCatCode = "" Then
        MsgBox "Please Enter Category Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtCatCode.SetFocus
    ElseIf Len(txtCatCode) < 3 Then
        MsgBox "Category Code must be 3 integer!", vbOKOnly + vbCritical, DialogBoxTitle
        txtCatCode.SetFocus
    ElseIf txtCatName = "" Then
        MsgBox "Please Enter Category Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtCatName.SetFocus
    ElseIf txtAmount = "" Then
        MsgBox "Please Enter Amount", vbOKOnly + vbCritical, DialogBoxTitle
        txtAmount.SetFocus
    Else
        strsql = "Select * From CheckItems where CatCode =  '" & Trim(txtCatCode) & "'"
        CHK.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If CHK.RecordCount > 0 And EditFlag = False Then
           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           CHK.Close
           txtCatCode.SetFocus
           txtCatCode.SelStart = 0
           txtCatCode.SelLength = Len(txtCatCode)
           Set CHK = Nothing
           Exit Sub
        Else
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           If RecordSaved = vbYes Then
              If EditFlag = False Then
                 SaveCHK.Open "CheckItems", MediConnCISDB, adOpenKeyset, adLockOptimistic
                 SaveCHK.AddNew
              Else
                 strsql = "Select * From CheckItems Where CatCode = '" & ChkStr(txtCatCode) & "'"
                 SaveCHK.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
              End If
              
              With SaveCHK
                !CatCode = ChkStr(txtCatCode)
                !CatName = ChkStr(txtCatName)
                !Amount = ChkStr(txtAmount)
                !ChkLastUpdateUser = Logon
                !ChkLastUpdateDate = Date
                
              End With
              SaveCHK.Update
            
             Call ClearForm(Me)
             'Call SearchCOP
             Call CheckListing(CheckCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             CHK.Close
             Set SaveCHK = Nothing
        End If
        EditFlag = False
     End If
   End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub CheckListing(Optional strAppend As String)
Const METHOD_NAME = "CheckListing"
On Error GoTo debugs
 'On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim CHK As ListItem
    Dim sql As String
    Dim TotalCHKCount As Integer
   
    TotalCHKCount = 0
    lstCheckItemList.ListItems.Clear
    
    sql = " Select CatCode, CatName, Amount " & _
      " From CheckItems where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            With rst2
            Set CHK = lstCheckItemList.ListItems.Add(, , rst2!CatCode)
                CHK.SubItems(1) = rst2!CatName
                CHK.SubItems(2) = rst2!Amount
                rst2.MoveNext
            End With
        TotalCHKCount = TotalCHKCount + 1
        txtTotalRecord = TotalCHKCount
        Loop
    End If
    CheckCriteria = ""
    rst2.Close
    Set rst2 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstCheckItemList_Click()
Const METHOD_NAME = "lstCheckItemList_Click"
On Error GoTo debugs
Dim strsql As String
    
    If lstCheckItemList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtCatCode = lstCheckItemList.SelectedItem.Text
        txtCatCode.Enabled = False
        txtCatName = lstCheckItemList.SelectedItem.SubItems(1)
        txtCatName.Enabled = False
        txtAmount = lstCheckItemList.SelectedItem.SubItems(2)
        txtAmount.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstCheckItemList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstCheckItemList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtCatCode = lstCheckItemList.SelectedItem.Text
        txtCatCode.Enabled = False
        txtCatName = lstCheckItemList.SelectedItem.SubItems(1)
        txtCatName.Enabled = False
        txtAmount = lstCheckItemList.SelectedItem.SubItems(2)
        txtAmount.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
    

End Sub

Private Sub lstCheckItemList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstCheckItemList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        txtCatCode = lstCheckItemList.SelectedItem.Text
        txtCatCode.Enabled = False
        txtCatName = lstCheckItemList.SelectedItem.SubItems(1)
        txtCatName.Enabled = False
        txtAmount = lstCheckItemList.SelectedItem.SubItems(2)
        txtAmount.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstCheckItemList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstCheckItemList.SortKey = ColumnHeader.Index - 1
    lstCheckItemList.Sorted = True
    If lstCheckItemList.SortOrder = lvwAscending Then
       lstCheckItemList.SortOrder = lvwDescending
    Else
       lstCheckItemList.SortOrder = lvwAscending
    End If
    
    txtCatCode.Enabled = False
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtCatCode_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
'####### @@@ END SAVE, EDIT , ADD, DELETE CHECKING ITEMS @@@ #######'

