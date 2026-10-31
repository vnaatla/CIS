VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBordxRevise 
   Caption         =   "Bordereaux Revise"
   ClientHeight    =   8592
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8592
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Bordereaux Revise"
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
         TabIndex        =   18
         Top             =   120
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   11655
      Begin VB.TextBox TxtBordx 
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   10
         Top             =   360
         Width           =   3855
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   280
         Left            =   5640
         TabIndex        =   9
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx No "
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblBordxPrintedDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Bordx Printed Date"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblBordxDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4920
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblSearhCriteria 
         Height          =   375
         Left            =   6720
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   11655
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   7440
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdReject 
         Caption         =   "&Remove "
         Height          =   375
         Left            =   8520
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10440
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9480
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   6720
      Width           =   11655
      Begin VB.TextBox lblTotalAmt 
         Height          =   285
         Left            =   9240
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Items"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Approved Amount"
         Height          =   255
         Left            =   7680
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lstWS 
      Height          =   4935
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   8700
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Claim No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Claim Reg Date"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Mem No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Patient"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Date Visit"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Approved Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Case Status"
         Object.Width           =   1940
      EndProperty
   End
End
Attribute VB_Name = "FrmBordxRevise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bExit As Boolean
Dim WSNoArray() As String
Dim ItemChecked As Boolean
Private m_blnDirection(1 To 7) As Boolean
Dim BordxAmt As Double
Dim ScanCode As String
Dim TotalApproved As Double


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '--- Enable PrintScreen feature
    '--- IMPORTANT - Don't forget to set .KeyPreview = True for this form at design time!!
    Const METHOD_NAME = "Form_KeyDown"
    On Error GoTo debugs
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdClear_Click()
    TxtBordx.Text = ""
    lblBordxPrintedDate = ""
    lblBordxDate = ""
    lblTotal = ""
    lblTotalAmt = ""
    lstWS.ListItems.Clear
    ReDim WSNoArray(0)
End Sub

Private Sub CmdPrint_Click()
Const METHOD_NAME = "CmdPrint_Click"
On Error GoTo debugs
Screen.MousePointer = vbHourglass
    
    If lstWS.ListItems.Count <> 0 Then
        'RptBordxRevise.Show
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
       
    Screen.MousePointer = vbDefault
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_Load()
Const METHOD_NAME = "Form_Load"
On Error GoTo debugs
Call ActionLogUpdate(12, "FrmBordxRevise")

    ReDim WSNoArray(0)
    bExit = False
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
'***************************

Private Sub CmdGo_Click()
    ScanCode = TxtBordx
    If Mid(ScanCode, 1, 2) = "LO" Then
        TxtBordx = ""
        'Call SearchForm(ScanCode)
    Else
        Screen.MousePointer = vbHourglass
        ReDim WSNoArray(0)
        If Trim(TxtBordx) = "" Then
            MsgBox "Please select a Bordx No ", vbOKOnly + vbCritical, DialogBoxTitle
            TxtBordx.SetFocus
        Else
            Call SearchRecord
        End If
        Screen.MousePointer = Default
    End If
End Sub


Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs
Dim strAppend As String
    
    strAppend = ""
    If Len(Trim(TxtBordx)) Then
        strAppend = strAppend + " And BordxRefNo = '" & Trim(TxtBordx) & "'"
    End If
    
    lblSearhCriteria = strAppend
    Call BordxReviseListing(strAppend)
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
'****************** End Search Record *****************************

'****************** Start BordxReviseListing Listing *****************************

Private Sub BordxReviseListing(Optional strAppend As String)
Const METHOD_NAME = "BordxReviseListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim ClaimAdmin As Boolean
Dim total As String
Dim TotalAmt As Double
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    total = 0
    
    lstWS.ListItems.Clear

    strAppendCase = "3"
    
    sql = sql + strAppend
            
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimBordxModule"
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
    
    'sql = "Select * From VIEW_CIS_BordxRevise where 0=0 "
    
    'If Len(Trim(strAppend)) Then
        'sql = sql + strAppend
    'End If
    
       
    'rst2.Open sql, MediConnCISWS, adOpenKeyset, adLockOptimistic
   
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = False
    Else
        CmdPrint.Enabled = True
        If Len(rst2!BordxPrintedDate) Then
            lblBordxPrintedDate = rst2!BordxPrintedDate
        End If
        lblBordxDate = rst2!BordxDate
        'Total = rst2.RecordCount
        BordxAmt = rst2!BordxClaimAmt
        'lblTotal = Total
        Do Until rst2.EOF
            With rst2
                Set Record = lstWS.ListItems.Add(, , !CASNumber)
                If Len(!ClaimOpenDate) Then
                    Record.SubItems(1) = Format(!ClaimOpenDate, "dd/mm/yyyy")
                End If
                If Len(!MBMNumber) Then
                   Record.SubItems(2) = Trim(!MBMNumber)
                End If
                If Len(!PatientName) Then
                    Record.SubItems(3) = Trim(!PatientName)
                End If
                If Len(!DateVisit) Then
                   Record.SubItems(4) = Format(!DateVisit, "dd/mm/yyyy")
                End If
                If Len(!CLMTotalApproved) Then
                    Record.SubItems(5) = Format(!CLMTotalApproved, "#,##0.00")
                End If
                If Len(!CASStatus) Then
                    Record.SubItems(6) = !CASStatus
                End If
                If Len(rst2!CLMTotalApproved) Then
                    TotalAmt = TotalAmt + rst2!CLMTotalApproved
                End If
            End With
            
            total = total + 1
            lblTotal = total
            
            rst2.MoveNext
        Loop
        lblTotalAmt = Format(TotalAmt, "#,##0.00;(#,#.00)")
    End If
    
    rst2.Close
    Set rst2 = Nothing
    
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'**********************Start Check ListView Record*******************************

Private Sub lstWS_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Const METHOD_NAME = "lstWS_ItemCheck"
On Error GoTo debugs
Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim temp
Static Arraycnt As Integer
    If Item.Checked = True Then
            ReDim Preserve WSNoArray(Arraycnt)
            WSNoArray(Arraycnt) = Item.Text
            Arraycnt = Arraycnt + 1
            ItemChecked = True
    Else
        ' Remove from the array
        newbound = 0
    
        If UBound(WSNoArray) <> 0 Then
            For i = 0 To UBound(WSNoArray)
                If Trim(WSNoArray(i)) <> "" Then
                    If (Item.SubItems(2)) <> WSNoArray(i) Then
                        ReDim Preserve tempArray(newbound)
                        tempArray(newbound) = WSNoArray(i)
                        newbound = newbound + 1
                    End If
                End If
            Next
            Erase WSNoArray
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
            ReDim WSNoArray(UBound(tempArray))
            For i = 0 To UBound(tempArray)
                WSNoArray(i) = tempArray(i)
            Next
            Erase tempArray
        Else
            Erase WSNoArray
            ReDim WSNoArray(0)
            Arraycnt = 0
            ItemChecked = False
        End If
    End If
'**********************End Check ListView Record*******************************
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

'**********************Start Sorted ListView *******************************
Private Sub lstWS_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Const METHOD_NAME = "lstWS_ColumnClick"
    On Error GoTo debugs
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    
    Case 1
        SortListView lstWS, ColumnHeader.Index, ldtstring, m_blnDirection(1)
    Case 2
        SortListView lstWS, ColumnHeader.Index, ldtDatetime, m_blnDirection(2)
    Case 3
        SortListView lstWS, ColumnHeader.Index, ldtstring, m_blnDirection(3)
    Case 4
        SortListView lstWS, ColumnHeader.Index, ldtstring, m_blnDirection(4)
    Case 5
        SortListView lstWS, ColumnHeader.Index, ldtDatetime, m_blnDirection(5)
    Case 6
        SortListView lstWS, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstWS, ColumnHeader.Index, ldtstring, m_blnDirection(7)
    End Select
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'**********************End Sorted ListView *******************************

Private Sub CmdReject_Click()
Const METHOD_NAME = "CmdReject_Click"
On Error GoTo debugs
Dim ii As Integer
Dim Cnt As Integer
Dim Update
Dim recordFound As Boolean
Static Arraycnt As Integer
    If ItemChecked = True Then
        Update = MsgBox("You are going to remove a worksheet from current Bordereaux. " & vbNewLine & _
                "Are you sure?", vbYesNo + vbExclamation)
        
         If Update = vbYes Then
            Cnt = 0
            recordFound = False
            If lstWS.ListItems.Count <> 0 Then
                For ii = 0 To UBound(WSNoArray)
                    If Trim(WSNoArray(ii)) <> "" Then
                       RemoveFromBordx WSNoArray(ii)
                       recordFound = True
                    End If
                Next ii
                If recordFound = True Then
                    MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                 End If
                Erase WSNoArray
                Arraycnt = 0
                ReDim WSNoArray(0)
                Call ClearForm(Me)
                lstWS.ListItems.Clear
                lblBordxPrintedDate = ""
                lblBordxDate = ""
                lblTotal = ""
                lblTotalAmt = ""
                
            End If
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

'**********************Start Update all Record functions*******************************

Private Sub RemoveFromBordx(tWSNo As String)
Const METHOD_NAME = "RemoveFromBordx"
On Error GoTo debugs
Dim StrSql As String
Dim strReduceBordx As String
Dim Claim As New ADODB.Recordset
Dim BordxArchive As New ADODB.Recordset
Dim strsqlArchive As String
Dim pos1 As Integer
Dim CASNumber As String
Dim strsqlBordx As String
Dim strsqlClmInv As String
Dim Bordx As New ADODB.Recordset
Dim ClmInv As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim cmd3 As New ADODB.Command
Dim prm5 As New ADODB.Parameter
Dim prm6 As New ADODB.Parameter
Dim strAppendCase As String
    
    
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
    pos1 = InStr(1, tWSNo)
    CASNumber = Trim(tWSNo)
  
    strAppendCase = "4"
    
    StrSql = "AND CASNumber ='" & Trim(CASNumber) & "'"
            
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimBordxModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrSql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    Claim.CursorLocation = adUseClient
    Claim.CursorType = adOpenForwardOnly
    Claim.CacheSize = 100
    Set Claim = cmd.Execute
  
  
  'strsql = " Select CASNUmber, ClaimBordxStatus, " & _
           " BordxRefNo, CLMFinalApprove, CLMBordxWSStatus, " & _
           " CLMBordxWSAscii From Claim where CASNumber ='" & Trim(CASNumber) & "'"
           
  'Claim.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
  
    TotalApproved = Claim!CLMFinalApprove
    strReduceBordx = " Update Bordx set BordxNoClaims = BordxNoClaims - 1 " & _
                     " Where BordxRefNo = '" & Claim!Bordxrefno & "'"
                    
    MediConnCISWS.Execute strReduceBordx
  
  
    strAppendCase = "5"
    
    strsqlBordx = "AND BordxRefNo = '" & Trim(TxtBordx) & "'"
            
    Set cmd2.ActiveConnection = MediConnCISWS
    cmd2.CommandText = "ClaimBordxModule"
    cmd2.CommandType = adCmdStoredProc
    cmd2.CommandTimeout = 300
    Set prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsqlBordx)
    cmd2.Parameters.Append prm3
    Set prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd2.Parameters.Append prm4
    Bordx.CursorLocation = adUseClient
    Bordx.CursorType = adOpenForwardOnly
    Bordx.CacheSize = 100
    Set Bordx = cmd2.Execute
  
  
  'strsqlBordx = "Select BordxClaimAmt, BordxDate from Bordx WHERE BordxRefNo = '" & Trim(TxtBordx) & "'"
  'Bordx.Open strsqlBordx, MediConnCISWS, adOpenKeyset, adLockOptimistic

  
    strAppendCase = "6"
    
    strsqlClmInv = "AND INVBordxRefNo = '" & Trim(TxtBordx) & "' AND CASNumber = '" & Trim(CASNumber) & "'"
            
    Set cmd3.ActiveConnection = MediConnCISWS
    cmd3.CommandText = "ClaimBordxModule"
    cmd3.CommandType = adCmdStoredProc
    cmd3.CommandTimeout = 300
    Set prm5 = cmd3.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsqlClmInv)
    cmd3.Parameters.Append prm5
    Set prm6 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd3.Parameters.Append prm6
    ClmInv.CursorLocation = adUseClient
    ClmInv.CursorType = adOpenForwardOnly
    ClmInv.CacheSize = 100
    Set ClmInv = cmd3.Execute
  
  'strsqlClmInv = " Select CASNumber, INVBordxRefNo, INVBordxed from ClaimInvoice where INVBordxRefNo = '" & Trim(TxtBordx) & "'" & _
                 " AND CASNumber = '" & Trim(CASNumber) & "'"
  'ClmInv.Open strsqlClmInv, MediConnCISWS, adOpenKeyset, adLockOptimistic
        
    BordxArchive.Open "BordxWSArchive", MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    
    BordxArchive.AddNew
     With BordxArchive
        !CASNumber = Trim(CASNumber)
        !Bordxrefno = Trim(TxtBordx)
        !BordxDate = Bordx!BordxDate
        !LastUpdateDate = Format(Date, "dd/mm/yyyy")
        !LastUpdateUser = Logon
        .Update
    End With
    BordxArchive.Close
    Set BordxArchive = Nothing
    
    'If Bordx.RecordCount > 0 Then
    If Bordx.EOF = False Then
        Bordx!BordxClaimAmt = CDbl(Bordx!BordxClaimAmt) - TotalApproved
        Bordx.Update
        Bordx.Close
        Set Bordx = Nothing
    End If
    
    'If ClmInv.RecordCount > 0 Then
    If ClmInv.EOF = False Then
       With ClmInv
          !INVBordxRefNo = ""
          !INVBordxed = "0"
       End With
       ClmInv.Update
       ClmInv.Close
       Set ClmInv = Nothing
    End If
    
    'If Claim.RecordCount > 0 Then
    If Claim.EOF = False Then
       With Claim
          !ClaimBordxStatus = "0"
          !Bordxrefno = ""
          If !CLMBordxWSASCII = "Y" Then
            !CLMBordxWSASCII = "N"
            !CLMBordxWSStatus = "D"
          End If
       End With
       Claim.Update
       Claim.Close
       Set Claim = Nothing
    End If
    
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub TxtBordx_GotFocus()
    CmdGo.Default = True
End Sub

Private Sub TxtBordx_LostFocus()
    TxtBordx = ChkStr(TxtBordx)
    CmdGo.Default = False
End Sub


