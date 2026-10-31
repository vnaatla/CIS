VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWSBordx2 
   Caption         =   "Insert Invoices for Bordx"
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
      Height          =   400
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Select Invoices for Bordx"
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
         TabIndex        =   21
         Top             =   80
         Width           =   3285
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   7200
      Width           =   11655
      Begin VB.CommandButton CmdCheck 
         Caption         =   "&Check All"
         Height          =   375
         Left            =   7560
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10920
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10320
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   8400
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   9120
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdBordx 
         Caption         =   "&Bordx"
         Height          =   375
         Left            =   9720
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblBordxNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   6480
         TabIndex        =   19
         Top             =   330
         Width           =   975
      End
      Begin VB.Label LblTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "of"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Last Generated Bordx No"
         Height          =   375
         Left            =   5280
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtTempPayor 
      Height          =   315
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   0
      Top             =   8040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.ListView lstWorksheetTrackList 
      Height          =   5535
      Left            =   120
      TabIndex        =   22
      Top             =   1680
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   9758
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Payor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Invoice No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Invoice Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Group Company"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "App Amt"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   120
      TabIndex        =   23
      Top             =   360
      Width           =   11655
      Begin MSMask.MaskEdBox txtInvDOR1 
         Height          =   300
         Left            =   1320
         TabIndex        =   1
         Top             =   255
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDOR2 
         Height          =   300
         Left            =   3240
         TabIndex        =   2
         Top             =   255
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboGrpCompany 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   525
         Width           =   3975
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   810
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "Payor Code"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   930
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Invoice DOR"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         ToolTipText     =   "Date of Admission"
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmWSBordx2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim CaseNoArray() As String
Dim CLMBordxNo As String
Dim CASNo As String
Dim TotalAPP As Variant
Dim TotCLM As Integer
Dim listloop1 As Integer
Dim listloop2 As Integer
Dim Check As Boolean
Dim Check2 As Boolean
Private m_blnDirection(1 To 6) As Boolean


Private Sub CboGrpCompany_LostFocus()
Const METHOD_NAME = "CboGrpCompany_LostFocus"
On Error GoTo debugs
Dim Sql As String
Dim rst2 As New ADODB.Recordset

    Sql = ""
        
    Sql = " Select GRPCoCode, GrpCoPaycode From GRPCompany Where GRPCoCode = '" & Mid(CboGrpCompany, 1, 3) & "'"
    
    rst2.Open Sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    
    If rst2.RecordCount Then
        TxtTempPayor = Trim(rst2!GRPCoPAYCode)
        CboPayor = Trim(rst2!GRPCoPAYCode)
    End If
    
    rst2.Close
    Set rst2 = Nothing
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
    Call ActionLogUpdate(12, "frmWSBordx2")
    Check = False
    Check2 = False
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    Call PayorComboListing
    ReDim CaseNoArray(0)
    
End Sub

'****************** Start Command Button *****************************
Private Sub CmdBordx_Click()
Const METHOD_NAME = "CmdBordx_Click"
On Error GoTo debugs
On Error GoTo errSave
Dim PAYCode As String
Dim ii As Integer
Dim Cnt As Integer
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim Update
Dim PayorIndex As Integer
Dim startTrans
    
    Cnt = 0
    RecordLeftOver = 0
    startTrans = False
    
    TotalAPP = 0
    TotCLM = 0
    
        
        If lstWorksheetTrackList.ListItems.Count = 0 Then
            MsgBox "Please select a record before proceeding!", vbOKOnly + vbCritical
        ElseIf UBound(CaseNoArray) = 0 Then
            MsgBox "Please select at least one record to be updated! ", vbCritical
        Else
            Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
             If Update = vbYes Then
                startTrans = True
                
             '   MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
             '   MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
             '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
             '   MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
                
                MediConnCISWS.beginTrans
                PAYCode = Trim(ChkStr(TxtTempPayor))
                CLMBordxNo = GenerateCLMBordxNo(PAYCode)
                
                For ii = 1 To UBound(CaseNoArray)
                 If Trim(CaseNoArray(ii)) <> "" Then
                    CASNo = Trim(Mid(CaseNoArray(ii), 1, IIf(InStr(1, CaseNoArray(ii), "-") = 0, 10, InStr(1, CaseNoArray(ii), "-") - 1)))
                                       
                    strsqlCLM = "Select CASNumber, CLMFinalApprove From Claim Where CASNumber ='" & Trim(CASNo) & "'"
                    CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
        
                    TotalAPP = TotalAPP + CLM!CLMFinalApprove
                    TotCLM = TotCLM + 1
                    CLM.Close
                    Set CLM = Nothing
                    Call UpdateRecord
                    recordFound = True
                 End If
               Next ii
               Erase CaseNoArray
               Arraycnt = 0
               ReDim CaseNoArray(0)

               Call SaveCLMBordx
                
                If startTrans = True Then
                    MediConnCISWS.commitTrans
                End If
                
                lstWorksheetTrackList.ListItems.Clear
                LblTotalAmt = 0
                lblTotal = 0
                lblCurrent = 0
                            
              '  MediConnCISDB.Close
              '  Set MediConnCISDB = Nothing
              '  medixmasterconn.Close
              '  Set medixmasterconn = Nothing
              '  MediConnCISWS.Close
              '  Set MediConnCISWS = Nothing
              '  MediConnCISPayment.Close
              '  Set MediConnCISPayment = Nothing
                
                MsgBox "Record Saved, Your Bordx No is : " & CLMBordxNo, vbOKOnly + vbInformation, DialogBoxTitle
                lblBordxNo = CLMBordxNo
                
            
            End If
        End If
    
    Exit Sub
errSave:
    MsgBox Err.Description
    If startTrans = True Then
        MediConnCISWS.RollbackTrans
        MediConnCISDB.Close
        Set MediConnCISDB = Nothing
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        MediConnCISWS.Close
        Set MediConnCISWS = Nothing
        MediConnCISPayment.Close
        Set MediConnCISPayment = Nothing
        
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    CboGrpCompany.Clear
    CboPayor.Clear
    lstWorksheetTrackList.ListItems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    lblBordxNo = ""
    Check = False
    Check2 = False
    ReDim CaseNoArray(0)
    Call PayorComboListing
End Sub

Private Sub CmdCheck_Click()
    SetCheckAllItems True
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    Check = False
    Check2 = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Public Sub CmdSearch_Click()
    
lblCurrent = 0
lblTotal = 0
LblTotalAmt = 0
lblBordxNo = ""
ReDim CaseNoArray(0)
    
'If CboGrpCompany = "" Then
    'MsgBox "Please select a Group Company Name", vbOKOnly + vbCritical, DialogBoxTitle
    'CboGrpCompany.SetFocus
'Else

    Screen.MousePointer = vbHourglass
      '  MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
      '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
      '  MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        
            Call SearchRecord
    
       ' MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
       ' MediConnCISWS.Close
       ' Set MediConnCISWS = Nothing
       ' MediConnCISPayment.Close
       ' Set MediConnCISPayment = Nothing
    Screen.MousePointer = Default
    
'End If

End Sub
'****************** End Command Button *****************************


'**********************Start Sorted ListView *******************************
Private Sub lstWorksheetTrackList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
    Case 5
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    End Select
End Sub
'**********************End Sorted ListView *******************************

'****************** Start Search Record *****************************

Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs
Dim strAppend As String
    
    strAppend = ""
    
    If Len(Trim(txtInvDOR1)) > 0 And IsDate(txtInvDOR1) = True Then
        strAppend = strAppend + " And RegisterDate >= '" & Format(txtInvDOR1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtInvDOR2)) > 0 And IsDate(txtInvDOR2) = True Then
        strAppend = strAppend + " And RegisterDate <= '" & Format(txtInvDOR2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(CboPayor)) Then
         strAppend = strAppend + " AND GRPCoPAYCode = '" & Mid(CboPayor, 1, 2) & "'"
    End If
    
    If Len(Trim(CboGrpCompany)) Then
         strAppend = strAppend + " AND GRPCoCode = '" & Mid(CboGrpCompany, 1, 3) & "'"
    End If
    
    Call WorksheetTrackingListing(strAppend)
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
'****************** End Search Record *****************************

'****************** Start WorksheetTracking Listing *****************************

Private Sub WorksheetTrackingListing(Optional strAppend As String, Optional Noprompt As Boolean)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant


    total = 0
    Current = 0
    
    lstWorksheetTrackList.ListItems.Clear

    Sql = " Select * From VIEW_CISINV_PendingBordx where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If

    rst2.Open Sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
             
    Do
        Do Until rst2.EOF
        'total = rst2.RecordCount
        lblTotal = total
        
            With rst2
                Set Record = lstWorksheetTrackList.ListItems.Add(, , "")
        
                If Len(!GRPCoPAYCode) Then
                    Record.SubItems(1) = !GRPCoPAYCode
                End If
                If Len(!InvoiceNo) Then
                    Record.SubItems(2) = Trim(!InvoiceNo)
                End If
                If Len(!InvoiceDate) Then
                   Record.SubItems(3) = Format(!InvoiceDate, "dd/mm/yyyy")
                End If
                If Len(!GrpCoName) Then
                    Record.SubItems(4) = Trim(!GrpCoName)
                End If
                If Len(!CLMTotalApproved) Then
                    Record.SubItems(5) = Format(!CLMTotalApproved, "#,##0.00")
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                If Len(!CLMTotalApproved) Then
                    TotalAmt = TotalAmt + !CLMTotalApproved
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
    Loop Until Check = False
    intExit = 0
    
    If lstWorksheetTrackList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdBordx.Enabled = True
        lblCurrent = 0
        lblTotal = 0
        LblTotalAmt = 0
    End If
    
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdBordx.Enabled = True
    
End Sub
'****************** End WorksheetTracking Listing *****************************

'****************** Start UpdateRecord *****************************

Private Sub UpdateRecord()
Const METHOD_NAME = "UpdateRecord"
On Error GoTo debugs
Dim Update
Dim strsqlCLMBordx As String
Dim ClaimBordx As New ADODB.Recordset
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim strsqlINV As String
Dim INV As New ADODB.Recordset
Dim Pending As New ADODB.Recordset
Dim strsql As String
        
    strsqlCLM = " Update Claim set " & _
                " BordxRefNo = '" & CLMBordxNo & "'," & _
                " ClaimBordxStatus = 1, CLMPending = 'N' " & _
                " where CASNumber ='" & Trim(CASNo) & "'"
    strsqlINV = " update ClaimInvoice set " & _
                " INVBordxRefNo = '" & CLMBordxNo & "'," & _
                " INVBordxed = 1 " & _
                " where CASNumber = '" & Trim(CASNo) & "'"
    MediConnCISWS.Execute strsqlCLM
    MediConnCISWS.Execute strsqlINV
    
    strsql = " Select CASNumber, Pending " & _
             " From CISCAS where CASNumber = '" & Trim(CASNo) & "'"
                 
    Pending.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
         
    If Pending.RecordCount > 0 Then
                  
        With Pending
           !Pending = "N"
        End With
        Pending.Update
        Pending.Close
        Set Pending = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'****************** End UpdateRecord *****************************


'****************** Start SaveCLMBordx *****************************

Private Sub SaveCLMBordx()
Const METHOD_NAME = "SaveCLMBordx"
On Error GoTo debugs
Dim strsqlCLMBordx As String
Dim ClaimBordx As New ADODB.Recordset
Dim GrpCoName As String
    
    GrpCoName = ""
    
    GrpCoName = Replace(CboGrpCompany, "'", "''")
    
    strsqlCLMBordx = " Select BordxPAYCode, BordxRefNo, " & _
                     " BordxDate, BordxBy, BordxClaimAmt, " & _
                     " BordxNoClaims, GroupCompany from Bordx " & _
                     " Where BordxRefNo = '" & Trim(CLMBordxNo) & "'"
    ClaimBordx.Open strsqlCLMBordx, MediConnCISWS, adOpenKeyset, adLockOptimistic

     ClaimBordx.AddNew
        With ClaimBordx
            !BordxPAYCode = Trim(ChkStr(TxtTempPayor))
            !GroupCompany = Trim(GrpCoName)
            !Bordxrefno = ChkStr(CLMBordxNo)
            !BordxDate = Date
            !BordxBy = Logon
            !BordxClaimAmt = TotalAPP
            !BordxNoClaims = TotCLM
        End With
    ClaimBordx.Update
    ClaimBordx.Close
    Set ClaimBordx = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'****************** End SaveCLMBordx *****************************

'****************** Start lstWorksheetTrackListItemCheck *****************************

Private Sub lstWorksheetTrackList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Static Arraycnt As Integer
Dim temp

    
'If check box is selected then
'call addtoArray

    If Item.Checked = True Then
    
            If Trim(Item.SubItems(3)) = "" Then
                MsgBox "Invoice Date cannot be empty.", vbCritical
                Item.Checked = False
            ElseIf Trim(Item.SubItems(4)) = "" Then
                MsgBox "Group Company cannot be empty!", vbCritical
                Item.Checked = False
            Else
                Arraycnt = IIf(UBound(CaseNoArray) = 0, 1, UBound(CaseNoArray) + 1)
                ReDim Preserve CaseNoArray(Arraycnt)
                CaseNoArray(Arraycnt) = Item.SubItems(2)
            End If
    Else
    
        ' Remove from the array
        newbound = 0
    
        If UBound(CaseNoArray) <> 0 Then
            For i = 0 To UBound(CaseNoArray)
                If (Item.SubItems(2)) <> CaseNoArray(i) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = CaseNoArray(i)
                    newbound = newbound + 1
                End If
            Next
            Erase CaseNoArray
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
            ReDim CaseNoArray(UBound(tempArray))
            For i = 0 To UBound(tempArray)
                CaseNoArray(i) = tempArray(i)
            Next
            Erase tempArray
        Else
            Erase CaseNoArray
            ReDim CaseNoArray(0)
            Arraycnt = 0
        End If
    End If

End Sub
'****************** End lstWorksheetTrackListItemCheck *****************************

Private Sub SetCheckAllItems(bState As Boolean)
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim R As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstWorksheetTrackList.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        R = SendMessageAny(lstWorksheetTrackList.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub

Public Sub SetCheck(hwnd As Long, lItemIndex As Long, bState As Boolean)
Dim LV As LVITEM
Dim R As Long
    With LV
    .mask = LVIF_STATE
    .state = IIf(bState, &H2000, &H1000)
    .stateMask = LVIS_STATEIMAGEMASK
    End With
    R = SendMessageAny(hwnd, LVM_SETITEMSTATE, lItemIndex, LV)
End Sub

'****************** Start ComboListing *****************************
Private Sub PayorComboListing()
Const METHOD_NAME = "PayorComboListing"
On Error GoTo debugs
Dim GrpCompany As New ADODB.Recordset
Dim PAYOR As New ADODB.Recordset


   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        
        GrpCompany.Open "Select GRPCoCode, GRPCoName from GRPCompany ", MediConnCISDB, adOpenKeyset, adLockOptimistic
        PAYOR.Open "Select PAYCode, PAYCompanyName from PAY ", MediConnCISDB, adOpenKeyset, adLockOptimistic
            
        Do Until GrpCompany.EOF
            CboGrpCompany.AddItem Trim(GrpCompany!GRPCoCode) & " - " & Trim(GrpCompany!GrpCoName)
            GrpCompany.MoveNext
        Loop
        GrpCompany.Close
        Set GrpCompany = Nothing
        
        Do Until PAYOR.EOF
            CboPayor.AddItem Trim(PAYOR!PAYCode) & " - " & Trim(PAYOR!PAYCompanyName)
            PAYOR.MoveNext
        Loop
        PAYOR.Close
        Set PAYOR = Nothing
        
  '  MediConnCISDB.Close
  '  Set MediConnCISDB = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'****************** End ComboListing *****************************


'****************** Start ComboBox Change/KeyPress*****************************

Private Sub CboGrpCompany_Change()
Const METHOD_NAME = "CboGrpCompany_Change"
On Error GoTo debugs
If CboGrpCompany.Text = "" Then
    If InStr(CboGrpCompany.Text, "null") Then
       CboGrpCompany.Enabled = False
    End If
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CboGrpCompany_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboGrpCompany_KeyPress"
On Error GoTo debugs
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
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboPayor_KeyPress"
On Error GoTo debugs
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayor.ListCount - 1
 
        If NewText = LCase(Left(CboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayor.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPayor.Text = ValidValue
                CboPayor.SelStart = 0
                CboPayor.SelLength = Len(ValidValue)
            End If
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

'****************** End ComboBox Change/KeyPress*****************************

Private Sub txtInvDOR1_LostFocus()
Const METHOD_NAME = "txtInvDOR1_LostFocus"
On Error GoTo debugs
    If IsDate(txtInvDOR1) Then
    Else
        If txtInvDOR1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtInvDOR1.SetFocus
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub txtInvDOR2_LostFocus()
Const METHOD_NAME = "txtInvDOR1_LostFocus"
On Error GoTo debugs
    If IsDate(txtInvDOR2) Then
    Else
        If txtInvDOR2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtInvDOR2.SetFocus
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
