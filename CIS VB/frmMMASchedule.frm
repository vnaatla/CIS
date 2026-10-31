VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMMASchedule 
   Caption         =   "MMA Schedule of Fees"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11796
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   11796
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   8400
      TabIndex        =   26
      Top             =   240
      Width           =   1695
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   255
         Left            =   860
         TabIndex        =   29
         Top             =   520
         Width           =   735
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   520
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   255
         Left            =   860
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Criteria"
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   8295
      Begin VB.TextBox txtMainBPMedix 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtProceduresMedix 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtSUGCat 
         Height          =   285
         Left            =   6600
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtAneCat 
         Height          =   285
         Left            =   6600
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtOCPS 
         Height          =   285
         Left            =   6600
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtMMAPart 
         Height          =   285
         Left            =   6600
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtPagefr 
         Height          =   285
         Left            =   6600
         TabIndex        =   7
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtPageTo 
         Height          =   285
         Left            =   7560
         TabIndex        =   6
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtMainBP 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtProcedures 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtCPTCode 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtCPTDesc 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label10 
         Caption         =   "CPT Code"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   7320
         TabIndex        =   24
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Page"
         Height          =   255
         Left            =   5760
         TabIndex        =   23
         Top             =   1220
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "MMA Part"
         Height          =   255
         Left            =   5760
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Ane 
         Caption         =   "Ane Cat"
         Height          =   255
         Left            =   5760
         TabIndex        =   21
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Body Part - Medix"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Medix Desc"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Surg Cat"
         Height          =   255
         Left            =   5760
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "OPCS"
         Height          =   255
         Left            =   5760
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "MMA Desc"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Body Part - MMA"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "CPT Desc"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lstMMAListing 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   9970
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Main Body Part Code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Main Body Part"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "OPCS "
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Procedures"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Surg's Cat"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Anae's Cat"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Surg's Fees"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Anae's Fees"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Part A Fees"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Page"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Part"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "CPTCode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "CPTDesc"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label38 
      Caption         =   "MMA Fee Listing"
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
      TabIndex        =   32
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "MMA Schedule of Fees"
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
      TabIndex        =   31
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmMMASchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MMACriteria As String
Private m_blnDirection(1 To 13) As Boolean

Private Sub cmdClear_Click()
    txtMainBP = ""
    txtProcedures = ""
    txtOCPS = ""
    lstMMAListing.ListItems.Clear
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelAnalysis1(lstMMAListing)
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()

Call ActionLogUpdate(12, "frmMMASchedule")

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

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSearch_Click()
    Dim strAppend As String
        
    Screen.MousePointer = vbHourglass
    cmdSearch.Enabled = False
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    strAppend = ""
           
        If Len(Trim(txtMainBP)) Then
            strAppend = strAppend + " And ProcMainBPCode like '%" & ChkStr(txtMainBP) & "%'"
        End If
    
        If Len(txtOCPS) Then
            strAppend = strAppend + " And OPCSCode like '%" & ChkStr(txtOCPS) & "%'"
        End If
        
        If Len(Trim(txtProcedures)) Then
            strAppend = strAppend + " And OPCSNarrative like '%" & ChkStr(txtProcedures) & "%'"
        End If
   
        If Len(Trim(txtSUGCat)) Then
            strAppend = strAppend + " And ProcSurgCat like '%" & Mid(txtSUGCat, 2, 8) & "%'"
        End If
    
        If Len(txtAneCat) Then
            strAppend = strAppend + " And ProcAnaesCat LIKE '%" & Mid(txtAneCat, 2, 8) & "%'"
        End If
        
        If Len(Trim(txtMMAPart)) Then
            strAppend = strAppend + " And ProcPart = '" & ChkStr(txtMMAPart) & "'"
        End If
        
        If Len(Trim(txtPagefr)) Then
            strAppend = strAppend + " And ProcPageNo >= '" & Trim(txtPagefr) & "'"
        End If
   
        If Len(Trim(txtPageTo)) Then
            strAppend = strAppend + " And ProcPageNo <= '" & Trim(txtPageTo) & "'"
        End If
   
        If Len(txtCPTCode) Then
            strAppend = strAppend + " And CPTCode = '" & Trim(txtCPTCode) & "'"
        End If
        
        If Len(Trim(txtCPTDesc)) Then
            strAppend = strAppend + " And CPTDesc LIKE '%" & Trim(txtCPTDesc) & "%'"
        End If
   
        MMACriteria = strAppend
        Call MMAListing(strAppend)
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
    cmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub MMAListing(Optional strAppend As String)
    Dim rst As New ADODB.Recordset
    Dim TotalPLNCount As Integer
    Dim MMAMaster As ListItem
    Dim StrSql As String
    Dim strsqlFees As String
    Dim rstFees As New ADODB.Recordset
    
    'TotalPLNCount = 0
    lstMMAListing.ListItems.Clear
    StrSql = "select ProcMainBPCode, OPCSCode, OPCSNarrative , PartACost, CPTCode,CPTDesc," & _
             "ProcSurgCat, ProcAnaesCat, ProcPageNo, ProcPart, SurgeonsFees, AnaesthetistsFees,Surgeon13thNew,Anesthetist13thNew " & _
             " From VIEW_MMA_Schedule where 0=0 "
    
    If Len(Trim(strAppend)) Then
        StrSql = StrSql + strAppend
    End If
    
    rst.Open StrSql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
    
    If rst.RecordCount = 0 Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
    Do Until rst.EOF
         Set MMAMaster = lstMMAListing.ListItems.Add(, , "")
             MMAMaster.SubItems(1) = IIf(Trim(rst!ProcMainBPCode) <> "", Trim(rst!ProcMainBPCode), "")
                          
             MMAMaster.SubItems(2) = IIf(Trim(rst!OPCSCode) <> "", Trim(rst!OPCSCode), "")
             
             MMAMaster.SubItems(3) = IIf(Trim(rst!OPCSNarrative) <> "", rst!OPCSNarrative, "")

             MMAMaster.SubItems(4) = IIf(Trim(rst!ProcSurgCat) <> "", Mid(rst!ProcSurgCat, 2, 8), "")

             MMAMaster.SubItems(5) = IIf(Trim(rst!ProcAnaesCat) <> "", Mid(rst!ProcAnaesCat, 2, 8), "")
             
             
             MMAMaster.SubItems(9) = IIf(Trim(rst!ProcPageNo) <> "", rst!ProcPageNo, "")

             MMAMaster.SubItems(10) = IIf(Trim(rst!ProcPart) <> "", rst!ProcPart, "")

            If Trim(rst!ProcPart) <> "A" Then
                MMAMaster.SubItems(6) = IIf(Trim(rst!Surgeon13thNew) <> "", rst!Surgeon13thNew, "")
                MMAMaster.SubItems(7) = IIf(Trim(rst!Anesthetist13thNew) <> "", rst!Anesthetist13thNew, "")
            Else
                MMAMaster.SubItems(8) = IIf(Trim(rst!PartACost) <> "", rst!PartACost, "")
            End If
            MMAMaster.SubItems(11) = IIf(Trim(rst!CPTCode) <> "", rst!CPTCode, "")

            MMAMaster.SubItems(12) = IIf(Trim(rst!CPTDesc) <> "", rst!CPTDesc, "")


             rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    End If
    
End Sub

Private Sub lstMMAListing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstMMAListing, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstMMAListing, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstMMAListing, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstMMAListing, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstMMAListing, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstMMAListing, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstMMAListing, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstMMAListing, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstMMAListing, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstMMAListing, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstMMAListing, ColumnHeader.Index, ldtString, m_blnDirection(11)
    
    End Select

End Sub

Private Sub lstMMAListing_DblClick()
    frmShowInvestigation.Source = 1
    frmShowInvestigation.Show vbModal
End Sub



