VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOSBordx 
   Caption         =   "Outstanding Bordx"
   ClientHeight    =   7884
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7884
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   7080
      Width           =   11775
      Begin VB.TextBox LblTotalDue 
         Height          =   285
         Left            =   3840
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   7920
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10440
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9600
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Export"
         Height          =   375
         Left            =   8760
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtNoofDay1 
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1715
         _ExtentY        =   445
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNoofDay2 
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1715
         _ExtentY        =   445
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "Total Due"
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   410
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Outstanding Bordx"
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
         TabIndex        =   2
         Top             =   100
         Width           =   3285
      End
   End
   Begin MSComctlLib.ListView lstOSBordxNonSihat 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   11895
      _ExtentX        =   20976
      _ExtentY        =   9335
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Payor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Date sent to Payor"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "No of Days O/S"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Bordx <14 days "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Bordx >14 days"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   0
      TabIndex        =   11
      Top             =   360
      Width           =   11775
      Begin VB.ComboBox CboGrpCompany 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   450
         Width           =   5775
      End
      Begin MSMask.MaskEdBox txtOSDate 
         Height          =   255
         Left            =   9960
         TabIndex        =   13
         Top             =   450
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   445
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Non - Sihat Malaysia"
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
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Outstanding Bordereaux As At"
         Height          =   255
         Left            =   7680
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Group Company"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Bordx with Group Company"
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
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   2415
   End
End
Attribute VB_Name = "FrmOSBordx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private aa As Integer
Dim TotalAmt1 As Variant
Dim TotalAmt2 As Variant
Dim TotalBordx1 As Variant
Dim TotalBordx2 As Variant
Dim TotalBordx3 As Variant
Dim TotalBordx4 As Variant
Dim TotalBordx5 As Variant
Dim TotalBordx6 As Variant
Dim TotalDue As Variant
Private m_blnDirection(1 To 8) As Boolean

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

Private Sub cmdprint_Click()
        
    Call ExportToExcelNonSihat(lstOSBordxNonSihat)

End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "FrmOSBordx")
    txtOSDate = Date
    Call ClearAmt
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    
            Call ComboListing
            
       ' MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
End Sub

'************************ Start Command Button *****************************
Private Sub cmdSearch_Click()
    
    Call ClearAmt
    CmdClear.Enabled = False
    CmdPrint.Enabled = False
    CmdExit.Enabled = False
    
    Screen.MousePointer = vbHourglass
    '    MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
     '   MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
            Call SearchRecord
    
    '    MediConnCISDB.Close
    '    Set MediConnCISDB = Nothing
    '    medixmasterconn.Close
     '   Set medixmasterconn = Nothing
     '   MediConnCISWS.Close
     '   Set MediConnCISWS = Nothing
        
    Screen.MousePointer = Default
   
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdClear_Click()
    CboGrpCompany.Clear
    lstOSBordxNonSihat.ListItems.Clear
    Call ClearAmt
    Call ComboListing
End Sub

Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim StrAppend As String
Dim GrpCompanyName As String
    
    
    StrAppend = ""
    GrpCompanyName = ""
    
    Call ClearAmt
    
    If Len(Trim(CboGrpCompany)) Then
         GrpCompanyName = Replace(CboGrpCompany, "'", "''")
         StrAppend = StrAppend + " AND GroupCompany = '" & Trim(GrpCompanyName) & "'"
    End If
    Call NONOSBordxPAYListing(StrAppend)
 
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
'************************ End Command Button *****************************

'************************** Start ComboListing **************************
Private Sub ComboListing()
Dim Company As New ADODB.Recordset

Company.Open "Select GRPCoName from GRPCompany", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    
    Do Until Company.EOF
        CboGrpCompany.AddItem Trim(Company!GrpCoName)
        Company.MoveNext
    Loop
    Company.Close
    Set Company = Nothing
    
End Sub
'************************** End ComboListing **************************

'************************** Start ComboBox Change/ KeyPress **************************
Private Sub CboGrpCompany_Click()
    If CboGrpCompany.Text = "" Then
        If InStr(CboGrpCompany.Text, "null") Then
           CboGrpCompany.Enabled = False
        End If
    End If
End Sub
'************************** End ComboBox Change/ KeyPress **************************

'************************** Start NONOSBordxPAYListing **************************
Private Sub NONOSBordxPAYListing(Optional StrAppend As String)
Const METHOD_NAME = "NONOSBordxPAYListing"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim intloop2 As Integer
Dim rstCount As New ADODB.Recordset
'Dim strCount As String
Dim Total As Integer
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String


    lstOSBordxNonSihat.ListItems.Clear
            
    strAppendCase = "14"
    
    Sql = Sql + StrAppend
            
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimBordxModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, Sql)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    'Sql = " Select BordxRefNo, BordxPAYCode, GroupCompany, " & _
          " BordxPAYSentDate,BordxClaimAmt, BordxPAYToStatus " & _
          " From VIEW_OS_CISBordx where BordxPAYToStatus = 1 and (BordxPAYFRStatus = 'P' OR BordxPAYFRStatus IS NULL) "
     
    'strCount = "select count(*) as totalRecord From VIEW_OS_CISBordx where BordxPAYToStatus = 1 and (BordxPAYFRStatus = 'P' OR BordxPAYFRStatus IS NULL) "
           
    'If Len(Trim(StrAppend)) Then
        'Sql = Sql + StrAppend
        'strCount = strCount + StrAppend
    'End If
    
    'Set rstCount = MediConnCISWS.Execute(strCount)
    'rst2.Open Sql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    If rst2.EOF = True Then
        MsgBox "No record found for this Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
         
    'Do
        Do Until rst2.EOF
        
        'Total = rstCount!totalrecord
        
        txtNoofDay2 = ""
            If IsDate(txtOSDate) And Len(rst2!BordxPAYSentDate) Then
                txtNoofDay2 = DateValue(txtOSDate) - DateValue((rst2!BordxPAYSentDate)) + 1
            End If
            
            If txtNoofDay2 = "0" Then
                txtNoofDay2 = txtNoofDay2 + 1
            End If
                
            With rst2
                intloop2 = intloop2 + 1
                
                Set Record = lstOSBordxNonSihat.ListItems.Add(, , intloop2)
                If Len(rst2!Bordxrefno) Then
                   Record.SubItems(1) = rst2!Bordxrefno
                End If
                
                If Len(rst2!BordxPAYCode) Then
                   Record.SubItems(2) = rst2!BordxPAYCode
                End If
                If Len(rst2!BordxClaimAmt) Then
                   Record.SubItems(3) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                If Len(rst2!BordxPAYSentDate) Then
                   Record.SubItems(4) = Format(rst2!BordxPAYSentDate, "dd/mm/yyyy")
                End If
                If Len(txtNoofDay2) Then
                   Record.SubItems(5) = txtNoofDay2
                End If
                If Trim(txtNoofDay2) <= 14 Then
                    Record.SubItems(6) = Format(rst2!BordxClaimAmt, "#,##0.00")
                    TotalBordx5 = TotalBordx5 + rst2!BordxClaimAmt
                    'Lbl14Day3 = Format(TotalBordx5, "#,##0.00")
                Else
                    Record.SubItems(7) = Format(rst2!BordxClaimAmt, "#,##0.00")
                    TotalBordx6 = TotalBordx6 + rst2!BordxClaimAmt
                    'Lbl14Day4 = Format(TotalBordx6, "#,##0.00")
                End If
                
                TotalDue = TotalDue + rst2!BordxClaimAmt
                LblTotalDue = Format(TotalDue, "#,##0.00")
            End With
            rst2.MoveNext
        Loop
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdPrint.Enabled = True
    CmdExit.Enabled = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'************************** End NONOSBordxPAYListing **************************

'************************** Start ClearAmt **************************

Private Sub ClearAmt()
    
    LblTotalDue = 0
    TotalAmt1 = 0
    TotalAmt2 = 0
    TotalBordx1 = 0
    TotalBordx2 = 0
    TotalBordx3 = 0
    TotalBordx4 = 0
    TotalDue = 0

End Sub
'************************** End ClearAmt *****************************



'************************** Start Sorting listview *****************************


Private Sub lstOSBordxNonSihat_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
        
    'If Option1.Value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstOSBordxNonSihat, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
        Case 2
            SortListView lstOSBordxNonSihat, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView lstOSBordxNonSihat, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView lstOSBordxNonSihat, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView lstOSBordxNonSihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
        Case 6
            SortListView lstOSBordxNonSihat, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView lstOSBordxNonSihat, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView lstOSBordxNonSihat, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select
    'ElseIf Option3.Value = True Then
        'Select Case ColumnHeader.Index
        'Case 1
            'SortListView lstOSBordxMNRB, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
        'Case 2
            'SortListView lstOSBordxMNRB, ColumnHeader.Index, ldtString, m_blnDirection(2)
        'Case 3
            'SortListView lstOSBordxMNRB, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        'Case 4
            'SortListView lstOSBordxMNRB, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        'End Select
    'End If
End Sub
'************************** End Sorting listview *****************************




