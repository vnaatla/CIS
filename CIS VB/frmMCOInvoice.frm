VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMCOInvoice 
   Caption         =   "MCO Invoice"
   ClientHeight    =   8490
   ClientLeft      =   -1560
   ClientTop       =   -1500
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      Caption         =   "Listing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5160
      Left            =   120
      TabIndex        =   35
      Top             =   2160
      Width           =   11055
      Begin MSComctlLib.ListView lvwInvoice 
         Height          =   4755
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8387
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Payor Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Membership No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Cover ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Relationship"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Bordereaux Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Policy No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Health Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Batch No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Group Company"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Printed"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Invoice No"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   7200
      Width           =   11055
      Begin VB.CommandButton cmdINVClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   9240
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVPrintPreview 
         Caption         =   "Pre&view"
         Height          =   300
         Left            =   8400
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   7560
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   300
         Left            =   6720
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtInvoiceYear 
         Height          =   285
         Left            =   5400
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtInvoiceNo 
         Height          =   285
         Left            =   5880
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5640
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdINVClose 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   300
         Left            =   10080
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label203 
         Alignment       =   2  'Center
         Caption         =   "Total Supp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   2880
         TabIndex        =   34
         Top             =   240
         Width           =   960
      End
      Begin VB.Label txtINVTotalSupplementary1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3840
         TabIndex        =   33
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label txtINVTotalPrincipal 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1440
         TabIndex        =   32
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label202 
         Alignment       =   2  'Center
         Caption         =   "Total Principal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "MCO Invoice"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
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
         Top             =   60
         Width           =   3285
      End
   End
   Begin VB.Frame FrameINV 
      Caption         =   "Enter Criteria For Searching"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11055
      Begin VB.TextBox txtINVInvoiceNoTo 
         Height          =   315
         Left            =   8520
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtINVBatchNo 
         Height          =   315
         Left            =   8520
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cboINVCriteriaSel 
         Height          =   315
         ItemData        =   "frmMCOInvoice.frx":0000
         Left            =   2160
         List            =   "frmMCOInvoice.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtINVInvoiceNo 
         Height          =   315
         Left            =   6000
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtINVInvoiceNoFr 
         Height          =   315
         Left            =   6000
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtINVBordxDate 
         Height          =   315
         Left            =   6000
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Crystal.CrystalReport cRptINV 
         Left            =   9720
         Top             =   1680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         UserName        =   "sa"
         PrintFileLinesPerPage=   60
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSMask.MaskEdBox txtINVBordxDateTo 
         Height          =   315
         Left            =   8520
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtINVBordxDateFr 
         Height          =   315
         Left            =   6000
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog_INV 
         Left            =   1560
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox cboINVListingVer 
         Height          =   315
         ItemData        =   "frmMCOInvoice.frx":0041
         Left            =   2160
         List            =   "frmMCOInvoice.frx":004E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   650
         Width           =   5415
      End
      Begin VB.ComboBox cboINVPayor 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   930
         Width           =   5415
      End
      Begin VB.ComboBox cboINVGrpCom 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   930
         Width           =   5415
      End
      Begin VB.ComboBox cboINVPrintOpt 
         Height          =   315
         ItemData        =   "frmMCOInvoice.frx":009B
         Left            =   2160
         List            =   "frmMCOInvoice.frx":00A5
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox CboDepartment 
         Height          =   315
         Left            =   6000
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Department"
         Height          =   255
         Left            =   5040
         TabIndex        =   37
         Top             =   1260
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl29 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   5040
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbl30 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   7560
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbl36 
         Caption         =   "Print Option"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lbl35 
         Caption         =   "Listing Versions"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label lbl34 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lbl27 
         Caption         =   "Single/Batch/Range"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmMCOInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const INV_PAYCODE = 0
Const INV_MEMBERNO = 1
Const INV_COVERID = 2
Const INV_RELATION = 3
Const INV_BORDXDATE = 4
Const INV_POLICYNO = 5
Const INV_HLTCODE = 6
Const INV_BATCHNO = 7
Const INV_NAME = 8
Const INV_GRPCO = 9
Const INV_PRINTED = 10
Const INV_INVOICE = 11

Dim sDate As String
Dim DateFrom As String
Dim DateTo As String
Dim Exportsql As String
Dim PAYCode As String
Dim BordxDat As String
Dim MBMNo As String
Dim Printed As String
Dim CancelPrinting As Boolean
Dim bPrintPass As Boolean
Dim TotalSupplementary As Integer
Dim InvNo As String
Dim GetSubSupp As String
Dim GetSuppMCO As String
Dim errorFound As Boolean

Private Sub cboINVGrpCom_Click()
Dim sql As String
Dim GRPCOMName As String
Dim rst2 As New ADODB.Recordset

    If cboINVListingVer.ListIndex = 2 Then
    
        sql = ""
        GRPCOMName = ""

        'GRPCOMName = Replace(cboMCOGrpCom, "'", "''")

        GRPCOMName = RemStr(Trim(Mid(cboINVGrpCom, 7, IIf(InStr(7, cboINVGrpCom, "-") = 0, 50, InStr(7, cboINVGrpCom, "-") - 7))))
        
        '=Run Sp=================================================
        sql = "Exec SearchDepartment '" & Trim(GRPCOMName) & "'"
        rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        '===================================================
        'sql = " Select MBMDepartment From MBMOthers Where MBMGroupCompany = '" & Trim(GRPCOMName) & "' GROUP BY MBMDepartment"
        'rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    
        CboDepartment.Clear
        Do Until rst2.EOF
            With rst2
                CboDepartment.AddItem Trim(!MBMDepartment)
                rst2.MoveNext
            End With
        Loop
                
        rst2.Close
        Set rst2 = Nothing
    
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

Public Function CreatePath(NewPath) As Boolean
    'Add a trailing slash if none
    If Right(NewPath, 1) <> "\" Then
        NewPath = NewPath & "\"
    End If
    
    'Call API
    If MakeSureDirectoryPathExists(NewPath) <> 0 Then
        'No errors, return True
        CreatePath = True
    End If

End Function

Private Sub RecordSearching_INV(objListView As Object)
    Screen.MousePointer = vbHourglass
        Call INVPrincipalListing
    Screen.MousePointer = Default
End Sub

Private Sub INVPrincipalListing()
Dim rsSearch As New ADODB.Recordset
Dim rsPrint As New ADODB.Recordset
Dim sql As String
Dim itemToPrint As Integer
Dim isPrinted As String
Dim CoverQty As Integer
Dim INVCriteria As String
Dim TotalPrincipal As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String


TotalSupplementary = 0
lvwInvoice.ListItems.Clear

If cboINVListingVer = "Invoice - by Payor" Then
    Select Case cboINVCriteriaSel.ListIndex
        Case 0 'Criteria Selection for Invoice on Bordx Batch
                sDate = CStr(Month(txtINVBordxDate)) & "/" & CStr(Day(txtINVBordxDate)) & "/" & CStr(Year(txtINVBordxDate))
                'sql = sql & " WHERE MBMBordxDate = '" & sDate & "'"
                sql = sql & " AND MBMBordxDate = '" & sDate & "'"
                sql = sql & " AND MBMBatchNo = '" & txtINVBatchNo & "'"
                sql = sql & " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "'"
        Case 1   'Criteria Selection for Invoice on Bordx Range
                 DateFrom = CStr(Month(txtINVBordxDateFr)) & "/" & CStr(Day(txtINVBordxDateFr)) & "/" & CStr(Year(txtINVBordxDateFr))
                 DateTo = CStr(Month(txtINVBordxDateTo)) & "/" & CStr(Day(txtINVBordxDateTo)) & "/" & CStr(Year(txtINVBordxDateTo))
                'sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                sql = sql & " AND MBMBordxDate >= '" & DateFrom & "'"
                sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
                sql = sql & " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "'"
        Case 2  'Criteria Selection for Invoice on Invoice Number
                'sql = sql & " WHERE MBMCInvoiceNo = '" & txtINVInvoiceNo & "'"
                sql = sql & " AND MBMCInvoiceNo = '" & txtINVInvoiceNo & "'"
    End Select
ElseIf cboINVListingVer = "Invoice - by Group Company" Then
    Select Case cboINVCriteriaSel.ListIndex
    Case 0 'Criteria Selection for Invoice on Bordx Batch
            sDate = CStr(Month(txtINVBordxDate)) & "/" & CStr(Day(txtINVBordxDate)) & "/" & CStr(Year(txtINVBordxDate))
            'sql = sql & " WHERE MBMBordxDate = '" & sDate & "'"
            sql = sql & " AND MBMBordxDate = '" & sDate & "'"
            sql = sql & " AND MBMBatchNo = '" & txtINVBatchNo & "'"
            sql = sql & " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "'"
    Case 1   'Criteria Selection for Invoice on Bordx Range
             DateFrom = CStr(Month(txtINVBordxDateFr)) & "/" & CStr(Day(txtINVBordxDateFr)) & "/" & CStr(Year(txtINVBordxDateFr))
             DateTo = CStr(Month(txtINVBordxDateTo)) & "/" & CStr(Day(txtINVBordxDateTo)) & "/" & CStr(Year(txtINVBordxDateTo))
            'sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
            sql = sql & " AND MBMBordxDate >= '" & DateFrom & "'"
            sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
            sql = sql & " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "'"
    Case 2  'Criteria Selection for Invoice on Invoice Number
            'sql = sql & " WHERE MBMCInvoiceNo = '" & txtINVInvoiceNo & "'"
            sql = sql & " AND MBMCInvoiceNo = '" & txtINVInvoiceNo & "'"
    End Select
ElseIf cboINVListingVer = "Invoice - by Department" Then
    Select Case cboINVCriteriaSel.ListIndex
    Case 0 'Criteria Selection for Invoice on Bordx Batch
            sDate = CStr(Month(txtINVBordxDate)) & "/" & CStr(Day(txtINVBordxDate)) & "/" & CStr(Year(txtINVBordxDate))
            'sql = sql & " WHERE MBMBordxDate = '" & sDate & "'"
            sql = sql & " AND MBMBordxDate = '" & sDate & "'"
            sql = sql & " AND MBMBatchNo = '" & txtINVBatchNo & "'"
            sql = sql & " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "'"
            sql = sql & " AND MBMDepartment = '" & Trim(CboDepartment) & "'"
    Case 1  'Criteria Selection for Invoice on Bordx Range
            DateFrom = CStr(Month(txtINVBordxDateFr)) & "/" & CStr(Day(txtINVBordxDateFr)) & "/" & CStr(Year(txtINVBordxDateFr))
            DateTo = CStr(Month(txtINVBordxDateTo)) & "/" & CStr(Day(txtINVBordxDateTo)) & "/" & CStr(Year(txtINVBordxDateTo))
            'sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
            sql = sql & " AND MBMBordxDate >= '" & DateFrom & "'"
            sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
            sql = sql & " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "'"
            sql = sql & " AND MBMDepartment = '" & Trim(CboDepartment) & "'"
    Case 2  'Criteria Selection for Invoice on Invoice Number
            'sql = sql & " AND MBMCInvoiceNo = '" & txtINVInvoiceNo & "'"
            sql = sql & " WHERE MBMCInvoiceNo = '" & txtINVInvoiceNo & "'"
    End Select
End If

If cboINVPrintOpt.ListIndex = 0 Then
    sql = sql & " AND (MBMCInvoiced = 'N' OR MBMCInvoiced = '' OR MBMCInvoiced IS NULL )"
Else
    sql = sql & " AND MBMCInvoiced = 'Y' "
End If

sql = sql & " ORDER BY MBMNumber"


'rsSearch.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic, adCmdText

    strAppendCase = "1"
            
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "MCOFeeModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rsSearch.CursorLocation = adUseClient
    rsSearch.CursorType = adOpenForwardOnly
    rsSearch.CacheSize = 100
    Set rsSearch = cmd.Execute

    
    
    'Do Until rsSearch.EOF
    'If rsSearch.RecordCount = 0 Then
        'MsgBox "Record not found !", vbInformation + vbOKOnly
        'cmdINVPrint.Enabled = False
        'cmdINVPrintPreview.Enabled = False
        
    'Else
        cmdINVPrint.Enabled = True
        cmdINVPrintPreview.Enabled = True

            Do Until rsSearch.EOF
            
                Set ListMaster = lvwInvoice.ListItems.Add(, , rsSearch!PAYCode)
                    ListMaster.SubItems(1) = Trim(rsSearch!MBMNumber)
                    ListMaster.SubItems(2) = "00"
                    ListMaster.SubItems(3) = rsSearch!MBMMemberType
                    ListMaster.SubItems(4) = rsSearch!MBMBordxDate
                    If Len(rsSearch!MBMPolicyNo) Then
                        ListMaster.SubItems(5) = Trim(rsSearch!MBMPolicyNo)
                    End If
                    ListMaster.SubItems(6) = Trim(rsSearch!HLTCode)
                    If Len(rsSearch!MBMBatchNo) Then
                        ListMaster.SubItems(7) = rsSearch!MBMBatchNo
                    End If
                    ListMaster.SubItems(8) = Trim(rsSearch!MBMName)
                                        
                    If Len(Trim(rsSearch!MBMGroupCompany)) = 0 Or IsNull(rsSearch!MBMGroupCompany) = True Then
                        ListMaster.SubItems(9) = ""
                    Else
                        ListMaster.SubItems(9) = Trim(rsSearch!MBMGroupCompany)
                    End If
                     
                    If Len(rsSearch!MBMCInvoiceNo) Then
                        ListMaster.SubItems(11) = Trim(rsSearch!MBMCInvoiceNo)
                    End If
                    
                        TotalPrincipal = TotalPrincipal + 1
                        txtINVTotalPrincipal = TotalPrincipal
                        PAYCode = rsSearch!PAYCode
                        MBMNo = rsSearch!MBMNumber
                        BordxDat = rsSearch!MBMBordxDate
                        If Len(rsSearch!MBMCInvoiceNo) Then
                            InvNo = Trim(rsSearch!MBMCInvoiceNo)
                        End If
                        
                        itemToPrint = INV_PRINTED
                        cmdINVPrint.Enabled = True
                        cmdINVPrintPreview.Enabled = True
                            
                            If rsSearch.RecordCount > 0 Then
                                If Len(rsSearch!MBMCInvoiced) Then
                                    isPrinted = Trim(rsSearch!MBMCInvoiced)
                                Else
                                    isPrinted = "N"
                                End If
                                                        
                                If isPrinted = "Y" Then
                                    ListMaster.SubItems(itemToPrint) = "YES"
                                Else
                                    ListMaster.SubItems(itemToPrint) = "NO"
                                End If
                                Printed = Trim(ListMaster.SubItems(itemToPrint))
                            Else
                                ListMaster.SubItems(itemToPrint) = "NO"
                                Printed = ListMaster.SubItems(itemToPrint)
                            End If
                                                        
                            If Trim(rsSearch!INSCode) = "F" Or Trim(rsSearch!INSCode) = "H" Then
                                Call INVSuppListing
                            End If
            rsSearch.MoveNext
            Loop
    'End If
    
    If lvwInvoice.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
    End If
        
    rsSearch.Close
    Set rsSearch = Nothing
End Sub

Private Sub cboINVCriteriaSel_Change()
Select Case cboINVCriteriaSel.ListIndex
    Case 0
        lbl29.Caption = "Bordx Date"
        lbl30.Caption = "Batch Num"
        txtINVBordxDate.Visible = True
        txtINVBatchNo.Visible = True
        txtINVBordxDateFr.Visible = False
        txtINVBordxDateTo.Visible = False
        txtINVInvoiceNo.Visible = False
        txtINVInvoiceNoFr.Visible = False
        txtINVInvoiceNoTo.Visible = False
        cboINVPayor.Enabled = True
        cboINVGrpCom.Enabled = True
        txtINVBatchNo = ""
        txtINVBordxDate = "__/__/____"
        txtINVBordxDateFr = "__/__/____"
        txtINVBordxDateTo = "__/__/____"
        cboINVPrintOpt.ListIndex = 0
        cboINVPrintOpt.Enabled = True
        cboINVListingVer.Enabled = True
        cboINVListingVer.ListIndex = 0
        txtINVInvoiceNo = ""
        txtINVInvoiceNoFr = ""
        txtINVInvoiceNoTo = ""
        lvwInvoice.ListItems.Clear
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        txtINVTotalPrincipal = 0
        txtINVTotalSupplementary1 = 0
        
    Case 1
        lbl29.Caption = "Bordx From"
        lbl30.Caption = "To"
        txtINVBordxDate.Visible = False
        txtINVBatchNo.Visible = False
        txtINVBordxDateFr.Visible = True
        txtINVBordxDateTo.Visible = True
        txtINVInvoiceNo.Visible = False
        txtINVInvoiceNoFr.Visible = False
        txtINVInvoiceNoTo.Visible = False
        cboINVPayor.Enabled = True
        cboINVGrpCom.Enabled = True
        txtINVBatchNo = ""
        txtINVBordxDate = "__/__/____"
        txtINVBordxDateFr = "__/__/____"
        txtINVBordxDateTo = "__/__/____"
        cboINVPrintOpt.ListIndex = 0
        cboINVPrintOpt.Enabled = False
        cboINVListingVer.Enabled = True
        cboINVListingVer.ListIndex = 0
        txtINVInvoiceNo = ""
        txtINVInvoiceNoFr = ""
        txtINVInvoiceNoTo = ""
        lvwInvoice.ListItems.Clear
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        txtINVTotalPrincipal = 0
        txtINVTotalSupplementary1 = 0
        
    Case 2
        lbl29.Caption = "Invoice Num"
        lbl30.Caption = ""
        txtINVBordxDate.Visible = False
        txtINVBatchNo.Visible = False
        txtINVBordxDateFr.Visible = False
        txtINVBordxDateTo.Visible = False
        txtINVInvoiceNo.Visible = True
        txtINVInvoiceNoFr.Visible = False
        txtINVInvoiceNoTo.Visible = False
        cboINVPayor.Enabled = False
        cboINVGrpCom.Enabled = False
        txtINVBatchNo = ""
        txtINVBordxDate = "__/__/____"
        txtINVBordxDateFr = "__/__/____"
        txtINVBordxDateTo = "__/__/____"
        cboINVPrintOpt.ListIndex = 1
        cboINVPrintOpt.Enabled = False
        cboINVListingVer.ListIndex = 1
        'cboINVPrintSeq.ListIndex = 1
        cboINVListingVer.Enabled = False
        txtINVInvoiceNo = ""
        txtINVInvoiceNoFr = ""
        txtINVInvoiceNoTo = ""
        lvwInvoice.ListItems.Clear
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        txtINVTotalPrincipal = 0
        txtINVTotalSupplementary1 = 0
End Select
End Sub

Private Sub cboINVCriteriaSel_Click()
Select Case cboINVCriteriaSel.ListIndex
    Case 0
        lbl29.Caption = "Bordx Date"
        lbl30.Caption = "Batch Num"
        txtINVBordxDate.Visible = True
        txtINVBatchNo.Visible = True
        txtINVBordxDateFr.Visible = False
        txtINVBordxDateTo.Visible = False
        txtINVInvoiceNo.Visible = False
        txtINVInvoiceNoFr.Visible = False
        txtINVInvoiceNoTo.Visible = False
        cboINVPayor.Enabled = True
        cboINVGrpCom.Enabled = True
        txtINVBatchNo = ""
        txtINVBordxDate = "__/__/____"
        txtINVBordxDateFr = "__/__/____"
        txtINVBordxDateTo = "__/__/____"
        cboINVPrintOpt.ListIndex = 0
        cboINVPrintOpt.Enabled = True
        cboINVListingVer.Enabled = True
        cboINVListingVer.ListIndex = 0
        txtINVInvoiceNo = ""
        txtINVInvoiceNoFr = ""
        txtINVInvoiceNoTo = ""
        lvwInvoice.ListItems.Clear
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        txtINVTotalPrincipal = 0
        txtINVTotalSupplementary1 = 0
        
    Case 1
        lbl29.Caption = "Bordx From"
        lbl30.Caption = "To"
        txtINVBordxDate.Visible = False
        txtINVBatchNo.Visible = False
        txtINVBordxDateFr.Visible = True
        txtINVBordxDateTo.Visible = True
        txtINVInvoiceNo.Visible = False
        txtINVInvoiceNoFr.Visible = False
        txtINVInvoiceNoTo.Visible = False
        cboINVPayor.Enabled = True
        cboINVGrpCom.Enabled = True
        txtINVBatchNo = ""
        txtINVBordxDate = "__/__/____"
        txtINVBordxDateFr = "__/__/____"
        txtINVBordxDateTo = "__/__/____"
        cboINVPrintOpt.ListIndex = 0
        cboINVPrintOpt.Enabled = False
        cboINVListingVer.Enabled = True
        cboINVListingVer.ListIndex = 0
        txtINVInvoiceNo = ""
        txtINVInvoiceNoFr = ""
        txtINVInvoiceNoTo = ""
        lvwInvoice.ListItems.Clear
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        txtINVTotalPrincipal = 0
        txtINVTotalSupplementary1 = 0
        
    Case 2
        lbl29.Caption = "Invoice Num"
        lbl30.Caption = ""
        txtINVBordxDate.Visible = False
        txtINVBatchNo.Visible = False
        txtINVBordxDateFr.Visible = False
        txtINVBordxDateTo.Visible = False
        txtINVInvoiceNo.Visible = True
        txtINVInvoiceNoFr.Visible = False
        txtINVInvoiceNoTo.Visible = False
        cboINVPayor.Enabled = False
        cboINVGrpCom.Enabled = False
        txtINVBatchNo = ""
        txtINVBordxDate = "__/__/____"
        txtINVBordxDateFr = "__/__/____"
        txtINVBordxDateTo = "__/__/____"
        cboINVPrintOpt.ListIndex = 1
        cboINVPrintOpt.Enabled = False
        cboINVListingVer.ListIndex = 1
        'cboINVPrintSeq.ListIndex = 1
        cboINVListingVer.Enabled = False
        txtINVInvoiceNo = ""
        txtINVInvoiceNoFr = ""
        txtINVInvoiceNoTo = ""
        lvwInvoice.ListItems.Clear
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        txtINVTotalPrincipal = 0
        txtINVTotalSupplementary1 = 0
End Select
End Sub

Private Sub cboINVCriteriaSel_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboINVCriteriaSel.Text, cboINVCriteriaSel.SelStart) + Chr(KeyAscii) + Mid(cboINVCriteriaSel.Text, cboINVCriteriaSel.SelStart + cboINVCriteriaSel.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboINVCriteriaSel.ListCount - 1
        
        If NewText = LCase(Left(cboINVCriteriaSel.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINVCriteriaSel.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboINVCriteriaSel.Text = ValidValue
               cboINVCriteriaSel.SelStart = 0
               cboINVCriteriaSel.SelLength = Len(ValidValue)
             End If
        End If

End Sub

Private Sub cboINVListingVer_Click()
Select Case cboINVListingVer.ListIndex
    Case 0
        lbl34.Caption = "Payor"
        cboINVPayor.Visible = True
        cboINVGrpCom.Visible = False
        Label1.Visible = False
        CboDepartment.Visible = False
    Case 1
        lbl34.Caption = "Group Company"
        cboINVGrpCom.Visible = True
        cboINVPayor.Visible = False
        Label1.Visible = False
        CboDepartment.Visible = False
    Case 2
        lbl34.Caption = "Group Company"
        cboINVGrpCom.Visible = True
        cboINVPayor.Visible = False
        Label1.Visible = True
        CboDepartment.Visible = True
    End Select
End Sub

Private Sub cboINVListingVer_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboINVListingVer.Text, cboINVListingVer.SelStart) + Chr(KeyAscii) + Mid(cboINVListingVer.Text, cboINVListingVer.SelStart + cboINVListingVer.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboINVListingVer.ListCount - 1
        
        If NewText = LCase(Left(cboINVListingVer.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINVListingVer.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboINVListingVer.Text = ValidValue
               cboINVListingVer.SelStart = 0
               cboINVListingVer.SelLength = Len(ValidValue)
             End If
        End If
End Sub


Private Sub cboINVPayor_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboINVPayor.Text, cboINVPayor.SelStart) + Chr(KeyAscii) + Mid(cboINVPayor.Text, cboINVPayor.SelStart + cboINVPayor.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboINVPayor.ListCount - 1
        
        If NewText = LCase(Left(cboINVPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINVPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboINVPayor.Text = ValidValue
               cboINVPayor.SelStart = 0
               cboINVPayor.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboINVPrintOpt_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboINVPrintOpt.Text, cboINVPrintOpt.SelStart) + Chr(KeyAscii) + Mid(cboINVPrintOpt.Text, cboINVPrintOpt.SelStart + cboINVPrintOpt.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboINVPrintOpt.ListCount - 1
        
        If NewText = LCase(Left(cboINVPrintOpt.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINVPrintOpt.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboINVPrintOpt.Text = ValidValue
               cboINVPrintOpt.SelStart = 0
               cboINVPrintOpt.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cmdINVClear_Click()
    txtINVBatchNo = ""
    txtINVBordxDate = "__/__/____"
    txtINVBordxDateFr = "__/__/____"
    txtINVBordxDateTo = "__/__/____"
    txtINVInvoiceNo = ""
    txtINVInvoiceNoFr = ""
    txtINVInvoiceNoTo = ""
    lvwInvoice.ListItems.Clear
    cmdINVPrint.Enabled = False
    cmdINVPrintPreview.Enabled = False
    txtINVTotalPrincipal = 0
    txtINVTotalSupplementary1 = 0

End Sub

Private Sub cmdINVClose_Click()
    Unload Me
End Sub

Private Sub cmdINVPrint_Click()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim strsql As String
Dim strsql2 As String
Dim strsql3 As String
Dim strsql4 As String
Dim MCO As New ADODB.Recordset
Dim MCO3 As New ADODB.Recordset
Dim TotalPrin As String
Dim TotalSupp As String
Dim SubSupp As String
Dim SuppMCO As String
Dim TotalSuppMCO As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim strAppendCase As String
Dim strAppendCase2 As String
Dim MemberNo As String

TotalPrin = "0"
SubSupp = "0"
TotalSupp = "0"
SuppMCO = "0"
TotalSuppMCO = "0"
strCount = 0
i = 1
    
    Screen.MousePointer = vbHourglass
    
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
        End If
        
        If Err.Number > 0 Then
            MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
            GoTo ExitFunction
        End If
        
        On Error GoTo HANDLE_ERROR
        objExcel.Interactive = False
    
        strsql4 = "ORDER BY MBMBordxDate"
        
        If cboINVListingVer = "Invoice - by Payor" Then
        
            If cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 0 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report_Batch Where 0 = 0 "
                strsql2 = " AND MBMBordxDate = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                          " AND MBMBatchNo = " & Trim(txtINVBatchNo) & "" & _
                          " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "' AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "3"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
            ElseIf cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 1 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report_Batch Where 0 = 0 "
                strsql2 = " AND MBMBordxDate = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                          " AND MBMBatchNo = " & Trim(txtINVBatchNo) & "" & _
                          " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "' AND (MBMCInvoiceNo IS NOT NULL OR MBMCInvoiceNo <> '')"
                
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "3"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
                
            ElseIf cboINVCriteriaSel.ListIndex = 1 And cboINVPrintOpt.ListIndex = 0 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report Where 0 = 0 "
                strsql2 = " AND MBMBordxDate >= '" & Format(txtINVBordxDateFr, "mm/dd/yyyy") & "'" & _
                          " AND MBMBordxDate <= '" & Format(txtINVBordxDateTo, "mm/dd/yyyy") & "'" & _
                          " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "' AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
                          
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "4"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
            ElseIf cboINVCriteriaSel.ListIndex = 2 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report Where 0 = 0 "
                strsql2 = "AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
                
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "4"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
                
            End If
            
        ElseIf cboINVListingVer = "Invoice - by Group Company" Then
        
            If cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 0 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report_Batch Where 0 = 0 "
                strsql2 = " AND MBMBordxDate = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                          " AND MBMBatchNo = " & Trim(txtINVBatchNo) & "" & _
                          " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "' AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "3"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
            ElseIf cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 1 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report_Batch Where 0 = 0 "
                strsql2 = " AND MBMBordxDate = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                          " AND MBMBatchNo = " & Trim(txtINVBatchNo) & "" & _
                          " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "' AND (MBMCInvoiceNo IS NOT NULL OR MBMCInvoiceNo <> '')"
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "3"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                          
            ElseIf cboINVCriteriaSel.ListIndex = 1 And cboINVPrintOpt.ListIndex = 0 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report Where 0 = 0 "
                strsql2 = " AND MBMBordxDate >= '" & Format(txtINVBordxDateFr, "mm/dd/yyyy") & "'" & _
                          " AND MBMBordxDate <= '" & Format(txtINVBordxDateTo, "mm/dd/yyyy") & "'" & _
                          " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "' AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "4"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
            ElseIf cboINVCriteriaSel.ListIndex = 2 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report Where 0 = 0 "
                strsql2 = "AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
                
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "4"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
            End If
            
        ElseIf cboINVListingVer = "Invoice - by Department" Then
        
            If cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 0 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report_Batch Where 0 = 0 "
                strsql2 = " AND MBMBordxDate = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                          " AND MBMBatchNo = " & Trim(txtINVBatchNo) & "" & _
                          " AND MBMDepartment = '" & Trim(CboDepartment) & "'" & _
                          " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "' AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "3"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                          
            ElseIf cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 1 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report_Batch Where 0 = 0 "
                strsql2 = " AND MBMBordxDate = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                          " AND MBMBatchNo = " & Trim(txtINVBatchNo) & "" & _
                          " AND MBMDepartment = '" & Trim(CboDepartment) & "'" & _
                          " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "' AND (MBMCInvoiceNo IS NOT NULL OR MBMCInvoiceNo <> '')"
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "3"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                          
            ElseIf cboINVCriteriaSel.ListIndex = 1 And cboINVPrintOpt.ListIndex = 0 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report Where 0 = 0 "
                strsql2 = " AND MBMBordxDate >= '" & Format(txtINVBordxDateFr, "mm/dd/yyyy") & "'" & _
                          " AND MBMBordxDate <= '" & Format(txtINVBordxDateTo, "mm/dd/yyyy") & "'" & _
                          " AND MBMDepartment = '" & Trim(CboDepartment) & "'" & _
                          " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "' AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "4"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
            ElseIf cboINVCriteriaSel.ListIndex = 2 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report Where 0 = 0 "
                strsql2 = "AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
                
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "4"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
            End If
        
        End If
        
        MCO.CursorLocation = adUseClient
        MCO.CursorType = adOpenForwardOnly
        MCO.CacheSize = 100
        Set MCO = cmd.Execute
        
        Set cmd2.ActiveConnection = MediConnCISDB
        cmd2.CommandText = "MCOFeeModule"
        cmd2.CommandType = adCmdStoredProc
        cmd2.CommandTimeout = 300
        Set Prm1 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
        cmd2.Parameters.Append Prm1
        Set Prm2 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase2)
        cmd2.Parameters.Append Prm2
        MCO3.CursorLocation = adUseClient
        MCO3.CursorType = adOpenForwardOnly
        MCO3.CacheSize = 100
        Set MCO3 = cmd2.Execute
        
        'MCO.Open strsql, MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic
        Set objExcel = New Excel.Application
        
        'If MCO.EOF = True Then
        If Not MCO.EOF Then
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
        
            objExcel.WindowState = xlMaximized
    
        Set objWorkbook = objExcel.Workbooks.Add
        Set objWorksheet = objWorkbook.Sheets(1)
        objExcel.DisplayAlerts = False
        
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
            
            If cboINVListingVer = "Invoice - by Payor" Then
                objWorksheet.Cells(8, "A") = Trim(MCO3!PAYCompanyName)
                objWorksheet.Cells(9, "A") = Trim(MCO3!PAYAddress1)
                objWorksheet.Cells(10, "A") = Trim(MCO3!PAYAddress2)
                objWorksheet.Cells(11, "A") = Trim(MCO3!PAYAddress3)
                objWorksheet.Cells(13, "A") = "ATTENTION: " & " " & Trim(MCO3!PAYMAINContact)
                objWorksheet.Cells(15, "A") = "TEL :" & " " & Trim(MCO3!PAYTELno)
                objWorksheet.Cells(15, "C") = "FAX :" & " " & Trim(MCO3!PAYFAXNo)
            ElseIf cboINVListingVer = "Invoice - by Group Company" Then
                objWorksheet.Cells(8, "A") = Trim(MCO3!MBMGroupCompany)
                objWorksheet.Cells(9, "A") = Trim(MCO3!GRPCoAdd1)
                objWorksheet.Cells(10, "A") = Trim(MCO3!GRPCoAdd2)
                objWorksheet.Cells(11, "A") = Trim(MCO3!GRPCoAdd3)
                objWorksheet.Cells(13, "A") = "ATTENTION: " & " " & Trim(MCO3!GRPCoContact)
                objWorksheet.Cells(15, "A") = "TEL :" & " " & Trim(MCO3!GRPCoPhone)
                objWorksheet.Cells(15, "C") = "FAX :" & " " & Trim(MCO3!GRPCoFax)
            ElseIf cboINVListingVer = "Invoice - by Department" Then
                objWorksheet.Cells(8, "A") = Trim(MCO3!MBMGroupCompany)
                objWorksheet.Cells(9, "A") = Trim(MCO3!GRPCoAdd1)
                objWorksheet.Cells(10, "A") = Trim(MCO3!GRPCoAdd2)
                objWorksheet.Cells(11, "A") = Trim(MCO3!GRPCoAdd3)
                objWorksheet.Cells(13, "A") = "ATTENTION: " & " " & Trim(MCO3!GRPCoContact)
                objWorksheet.Cells(15, "A") = "TEL :" & " " & Trim(MCO3!GRPCoPhone)
                objWorksheet.Cells(15, "C") = "FAX :" & " " & Trim(MCO3!GRPCoFax)
                If Len(MCO3!MBMDepartment) Then
                    objWorksheet.Cells(7, "A") = Trim(MCO3!MBMDepartment)
                'Else
                    'objWorksheet.Cells(16, "A") = "DEPT :"
                End If
            End If
            
            Set objRange = objWorksheet.Rows(5)
            objRange.Font.Size = "13"
            objRange.Font.Bold = True
            objWorksheet.Cells(5, "G") = "Invoice"
            
            If cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 0 Then
                Call IncInvoiceNo
                objWorksheet.Cells(8, "H") = Trim(txtInvoiceNo)
            ElseIf cboINVCriteriaSel.ListIndex = 1 Then
                Call IncInvoiceNo
                objWorksheet.Cells(8, "H") = Trim(txtInvoiceNo)
            Else
                objWorksheet.Cells(8, "H") = Trim(MCO!MBMCInvoiceNo)
            End If
            
            objWorksheet.Cells(8, "G") = "Number :"
            objWorksheet.Cells(10, "G") = "Date :"
            objWorksheet.Cells(10, "H") = Format(Date, "dd/mm/yyyy")
            objWorksheet.Cells(12, "G") = "Terms :"
            objWorksheet.Cells(12, "H") = "30 DAYS"
            objWorksheet.Cells(19, "A") = "ITEM"
            objWorksheet.Cells(18, "B") = "BORDEREAUX"
            objWorksheet.Cells(19, "B") = "DATE"
            objWorksheet.Cells(19, "D") = "PLAN TYPE"
            objWorksheet.Cells(19, "E") = "ADULT/CHILD"
            objWorksheet.Cells(19, "F") = "QTY"
            objWorksheet.Cells(18, "G") = "COVERED"
            objWorksheet.Cells(19, "G") = "PERSONS"
            objWorksheet.Cells(19, "H") = "AMOUNT"
            
            Do Until MCO.EOF
                
                objWorksheet.Cells(21 + CDbl(i), "A") = i
                objWorksheet.Cells(21 + CDbl(i), "B") = Format(MCO!MBMBordxDate, "mm/dd/yyyy")
                If Trim(MCO!InsType) = "I" Then
                    objWorksheet.Cells(21 + CDbl(i), "D") = "I - INDIVIDUAL"
                ElseIf Trim(MCO!InsType) = "F" Then
                    objWorksheet.Cells(21 + CDbl(i), "D") = "F - FAMILY"
                ElseIf Trim(MCO!InsType) = "G" Then
                    objWorksheet.Cells(21 + CDbl(i), "D") = "G - GROUP INDIVIDUAL"
                ElseIf Trim(MCO!InsType) = "H" Then
                    objWorksheet.Cells(21 + CDbl(i), "D") = "H - GROUP FAMILY"
                End If
                
                If Trim(MCO!MBMAdultChild) = "A" Then
                    objWorksheet.Cells(21 + CDbl(i), "E") = "Adult"
                ElseIf Trim(MCO!MBMAdultChild) = "C" Then
                    objWorksheet.Cells(21 + CDbl(i), "E") = "Child"
                End If
                objWorksheet.Cells(21 + CDbl(i), "F") = IIf(IsNumeric(MCO!TotalPrincipal) = True, MCO!TotalPrincipal, 0)
                TotalPrin = CDbl(TotalPrin) + IIf(IsNumeric(MCO!TotalPrincipal) = True, MCO!TotalPrincipal, 0)
                objWorksheet.Cells(21 + CDbl(i), "H") = Format(MCO!MBMCLNMCOFee, "#,##0.00")
                
                If Trim(MCO!InsType) = "F" Or Trim(MCO!InsType) = "H" Then
                    
                    If cboINVListingVer = "Invoice - by Payor" Then
                    
                        If cboINVCriteriaSel.ListIndex = 0 Then
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report_Batch Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND MBMBatchNo = " & Trim(MCO!MBMBatchNo) & "" & _
                                      " AND PAYCode = '" & Trim(MCO3!PAYCode) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "5"
                            
                        ElseIf cboINVCriteriaSel.ListIndex = 1 Then
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate >= '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND MBMBordxdate <= '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND PAYCode = '" & Trim(MCO3!PAYCode) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "6"
                           
                        Else
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND PAYCode = '" & Trim(MCO3!PAYCode) & "' AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "6"
                                                        
                        End If
                        
                    ElseIf cboINVListingVer = "Invoice - by Group Company" Then
                    
                        If cboINVCriteriaSel.ListIndex = 0 Then
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report_Batch Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND MBMBatchNo = " & Trim(MCO3!MBMBatchNo) & "" & _
                                      " AND GrpCoCode = '" & Trim(MCO3!GRPCoCode) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "5"
                                      
                        ElseIf cboINVCriteriaSel.ListIndex = 1 Then
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND GrpCoCode = '" & Trim(MCO3!GRPCoCode) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "6"
                                            
                        Else
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND GrpCoCode = '" & Trim(MCO3!GRPCoCode) & "' AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "6"
                
                        End If
                        
                    ElseIf cboINVListingVer = "Invoice - by Department" Then
                    
                        If cboINVCriteriaSel.ListIndex = 0 Then
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report_Batch Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND MBMBatchNo = " & Trim(MCO3!MBMBatchNo) & "" & _
                                      " AND GrpCoCode = '" & Trim(MCO3!GRPCoCode) & "' AND MBMDepartment = '" & Trim(MCO3!MBMDepartment) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "5"
                
                        ElseIf cboINVCriteriaSel.ListIndex = 1 Then
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND GrpCoCode = '" & Trim(MCO3!GRPCoCode) & "' AND MBMDepartment = '" & Trim(MCO3!MBMDepartment) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "6"
                
                        Else
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND GrpCoCode = '" & Trim(MCO3!GRPCoCode) & "' AND MBMDepartment = '" & Trim(MCO3!MBMDepartment) & "' AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "6"
                        End If
                    End If
                                        
                    'MCO2.Open strsql3, MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic
                                
                        Call GetSuppMCOFee(strsql3, strAppendCase, TotalPrin, MCO!MBMCLNMCOFee)
                        objWorksheet.Cells(21 + CDbl(i), "G") = GetSubSupp
                        TotalSupp = CDbl(TotalSupp) + objWorksheet.Cells(21 + CDbl(i), "G")
                        
                        objWorksheet.Cells(21 + CDbl(i), "H") = GetSuppMCO
                        TotalSuppMCO = CDbl(TotalSuppMCO) + objWorksheet.Cells(21 + CDbl(i), "H")
                Else
                    
                    SubSupp = CDbl(SubSupp) + CDbl(TotalPrin)
                    objWorksheet.Cells(21 + CDbl(i), "G") = SubSupp
                    TotalSupp = CDbl(TotalSupp) + objWorksheet.Cells(21 + CDbl(i), "G")
                    
                End If
                            
            MCO.MoveNext
            MCO3.MoveNext
            
            i = i + 1
            TotalPrin = "0"
            SubSupp = "0"
            TotalSupp = "0"
            SuppMCO = "0"
            TotalSuppMCO = "0"
            
            Loop
            MCO.Close
            Set MCO = Nothing
            MCO3.Close
            Set MCO3 = Nothing
            
            
            Set objRange = objWorksheet.Rows(42)
            objRange.Font.Bold = True
                        
            objWorksheet.Cells(42, "F") = "'-------------------------------------------------------"
            objWorksheet.Cells(43, "D") = "TOTAL"
            objWorksheet.Range("F43").Formula = "=SUM(F22:F41)"
            objWorksheet.Range("G43").Formula = "=SUM(G22:G41)"
            objWorksheet.Range("H43").Formula = "=SUM(H22:H41)"
            objWorksheet.Cells(44, "F") = "'======================================="
            
            objWorksheet.Cells(49, "A") = "_________________________________________"
            objWorksheet.Cells(50, "A") = "For United Health Connnect Sdn. Bhd."
            objWorksheet.Cells(57, "A") = "Interest of 1% per annum will be levied after expiry of terms of payment."
            
                For i = 1 To lvwInvoice.ListItems.Count
                    If Trim(lvwInvoice.ListItems(i).ListSubItems(10).Text) = "NO" Then
                        MemberNo = Trim(lvwInvoice.ListItems(i).ListSubItems(1).Text)
                        Call UpdateINVPrintingFile(MemberNo)
                    End If
                Next i
                
            End If
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            Screen.MousePointer = vbDefault
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
End Sub

Private Sub cmdINVPrintPreview_Click()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim strsql As String
Dim strsql2 As String
Dim strsql3 As String
Dim strsql4 As String
Dim MCO As New ADODB.Recordset
Dim MCO2 As New ADODB.Recordset
Dim MCO3 As New ADODB.Recordset
Dim strCount
Dim TotalPrin As String
Dim TotalSupp As String
Dim SubSupp As String
Dim SuppMCO As String
Dim TotalSuppMCO As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim strAppendCase As String
Dim strAppendCase2 As String

TotalPrin = "0"
SubSupp = "0"
TotalSupp = "0"
SuppMCO = "0"
TotalSuppMCO = "0"
strCount = 0
i = 1
    
    Screen.MousePointer = vbHourglass

    
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
        End If
        
        If Err.Number > 0 Then
            MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
            GoTo ExitFunction
        End If
        
        On Error GoTo HANDLE_ERROR
        objExcel.Interactive = False
    
        strsql4 = "ORDER BY MBMBordxDate"
        
        If cboINVListingVer = "Invoice - by Payor" Then
        
            If cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 0 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report_Batch Where 0 = 0 "
                'strCount = " Select count(*) as totalRecord From VIEW_ClinicalInvoice_Prin_Report_Batch where 0 = 0 "
                strsql2 = " AND MBMBordxDate = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                          " AND MBMBatchNo = " & Trim(txtINVBatchNo) & "" & _
                          " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "' AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "3"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
            ElseIf cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 1 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report_Batch Where 0 = 0 "
                'strCount = " Select count(*) as totalRecord From VIEW_ClinicalInvoice_Prin_Report_Batch where 0 = 0 "
                strsql2 = " AND MBMBordxDate = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                          " AND MBMBatchNo = " & Trim(txtINVBatchNo) & "" & _
                          " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "' AND (MBMCInvoiceNo IS NOT NULL OR MBMCInvoiceNo <> '')"
                
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "3"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
                
            ElseIf cboINVCriteriaSel.ListIndex = 1 And cboINVPrintOpt.ListIndex = 0 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report Where 0 = 0 "
                'strCount = " Select count(*) as totalRecord From VIEW_ClinicalInvoice_Prin_Report where 0 = 0 "
                strsql2 = " AND MBMBordxDate >= '" & Format(txtINVBordxDateFr, "mm/dd/yyyy") & "'" & _
                          " AND MBMBordxDate <= '" & Format(txtINVBordxDateTo, "mm/dd/yyyy") & "'" & _
                          " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "' AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
                          
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "4"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
            ElseIf cboINVCriteriaSel.ListIndex = 2 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report Where 0 = 0 "
                'strCount = " Select count(*) as totalRecord From VIEW_ClinicalInvoice_Prin_Report where 0 = 0 "
                strsql2 = "AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
                
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "4"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
                
            End If
            
        ElseIf cboINVListingVer = "Invoice - by Group Company" Then
        
            If cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 0 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report_Batch Where 0 = 0 "
                'strCount = " Select count(*) as totalRecord From VIEW_ClinicalInvoice_Prin_Report_Batch where 0 = 0 "
                strsql2 = " AND MBMBordxDate = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                          " AND MBMBatchNo = " & Trim(txtINVBatchNo) & "" & _
                          " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "' AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "3"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
            ElseIf cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 1 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report_Batch Where 0 = 0 "
                'strCount = " Select count(*) as totalRecord From VIEW_ClinicalInvoice_Prin_Report_Batch where 0 = 0 "
                strsql2 = " AND MBMBordxDate = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                          " AND MBMBatchNo = " & Trim(txtINVBatchNo) & "" & _
                          " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "' AND (MBMCInvoiceNo IS NOT NULL OR MBMCInvoiceNo <> '')"
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "3"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                          
            ElseIf cboINVCriteriaSel.ListIndex = 1 And cboINVPrintOpt.ListIndex = 0 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report Where 0 = 0 "
                'strCount = " Select count(*) as totalRecord From VIEW_ClinicalInvoice_Prin_Report where 0 = 0 "
                strsql2 = " AND MBMBordxDate >= '" & Format(txtINVBordxDateFr, "mm/dd/yyyy") & "'" & _
                          " AND MBMBordxDate <= '" & Format(txtINVBordxDateTo, "mm/dd/yyyy") & "'" & _
                          " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "' AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "4"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
            ElseIf cboINVCriteriaSel.ListIndex = 2 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report Where 0 = 0 "
                'strCount = " Select count(*) as totalRecord From VIEW_ClinicalInvoice_Prin_Report where 0 = 0 "
                strsql2 = "AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
                
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "4"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
            End If
            
        ElseIf cboINVListingVer = "Invoice - by Department" Then
        
            If cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 0 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report_Batch Where 0 = 0 "
                'strCount = " Select count(*) as totalRecord From VIEW_ClinicalInvoice_Prin_Report_Batch where 0 = 0 "
                strsql2 = " AND MBMBordxDate = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                          " AND MBMBatchNo = " & Trim(txtINVBatchNo) & "" & _
                          " AND MBMDepartment = '" & Trim(CboDepartment) & "'" & _
                          " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "' AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "3"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                          
            ElseIf cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 1 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report_Batch Where 0 = 0 "
                'strCount = " Select count(*) as totalRecord From VIEW_ClinicalInvoice_Prin_Report_Batch where 0 = 0 "
                strsql2 = " AND MBMBordxDate = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                          " AND MBMBatchNo = " & Trim(txtINVBatchNo) & "" & _
                          " AND MBMDepartment = '" & Trim(CboDepartment) & "'" & _
                          " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "' AND (MBMCInvoiceNo IS NOT NULL OR MBMCInvoiceNo <> '')"
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "3"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                          
            ElseIf cboINVCriteriaSel.ListIndex = 1 And cboINVPrintOpt.ListIndex = 0 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report Where 0 = 0 "
                'strCount = " Select count(*) as totalRecord From VIEW_ClinicalInvoice_Prin_Report where 0 = 0 "
                strsql2 = " AND MBMBordxDate >= '" & Format(txtINVBordxDateFr, "mm/dd/yyyy") & "'" & _
                          " AND MBMBordxDate <= '" & Format(txtINVBordxDateTo, "mm/dd/yyyy") & "'" & _
                          " AND MBMDepartment = '" & Trim(CboDepartment) & "'" & _
                          " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "' AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
                          
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "4"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
            ElseIf cboINVCriteriaSel.ListIndex = 2 Then
                'strsql = " Select * From VIEW_ClinicalInvoice_Prin_Report Where 0 = 0 "
                'strCount = " Select count(*) as totalRecord From VIEW_ClinicalInvoice_Prin_Report where 0 = 0 "
                strsql2 = "AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
                
                strsql = strsql + strsql2 '+ strsql4
                
                strAppendCase = "4"
                strAppendCase2 = "13"
                
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append Prm2
                
            End If
        
        End If
        
        'strsql = strsql + strsql2 '+ strsql4
        'strCount = strCount + strsql2
        
        'Set strCount = MediConnCISDB.Execute(strCount)
        
        MCO.CursorLocation = adUseClient
        MCO.CursorType = adOpenForwardOnly
        MCO.CacheSize = 100
        Set MCO = cmd.Execute
        
        Set cmd2.ActiveConnection = MediConnCISDB
        cmd2.CommandText = "MCOFeeModule"
        cmd2.CommandType = adCmdStoredProc
        cmd2.CommandTimeout = 300
        Set Prm1 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
        cmd2.Parameters.Append Prm1
        Set Prm2 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase2)
        cmd2.Parameters.Append Prm2
        MCO3.CursorLocation = adUseClient
        MCO3.CursorType = adOpenForwardOnly
        MCO3.CacheSize = 100
        Set MCO3 = cmd2.Execute
        
        'Do Until MCO.EOF
            'strCount = strCount + 1
            'MCO.MoveNext
        'Loop
        
        'MCO.Open strsql, MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic
        Set objExcel = New Excel.Application
        
        'If MCO.EOF = True Then
        If Not MCO.EOF Then
            'RecordSelected = strCount
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
        
            objExcel.WindowState = xlMaximized
    
        Set objWorkbook = objExcel.Workbooks.Add
        Set objWorksheet = objWorkbook.Sheets(1)
        objExcel.DisplayAlerts = False
        
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
            
            If cboINVListingVer = "Invoice - by Payor" Then
                objWorksheet.Cells(8, "A") = Trim(MCO3!PAYCompanyName)
                objWorksheet.Cells(9, "A") = Trim(MCO3!PAYAddress1)
                objWorksheet.Cells(10, "A") = Trim(MCO3!PAYAddress2)
                objWorksheet.Cells(11, "A") = Trim(MCO3!PAYAddress3)
                objWorksheet.Cells(13, "A") = "ATTENTION: " & " " & Trim(MCO3!PAYMAINContact)
                objWorksheet.Cells(15, "A") = "TEL :" & " " & Trim(MCO3!PAYTELno)
                objWorksheet.Cells(15, "C") = "FAX :" & " " & Trim(MCO3!PAYFAXNo)
            ElseIf cboINVListingVer = "Invoice - by Group Company" Then
                objWorksheet.Cells(8, "A") = Trim(MCO3!MBMGroupCompany)
                objWorksheet.Cells(9, "A") = Trim(MCO3!GRPCoAdd1)
                objWorksheet.Cells(10, "A") = Trim(MCO3!GRPCoAdd2)
                objWorksheet.Cells(11, "A") = Trim(MCO3!GRPCoAdd3)
                objWorksheet.Cells(13, "A") = "ATTENTION: " & " " & Trim(MCO3!GRPCoContact)
                objWorksheet.Cells(15, "A") = "TEL :" & " " & Trim(MCO3!GRPCoPhone)
                objWorksheet.Cells(15, "C") = "FAX :" & " " & Trim(MCO3!GRPCoFax)
            ElseIf cboINVListingVer = "Invoice - by Department" Then
                objWorksheet.Cells(8, "A") = Trim(MCO3!MBMGroupCompany)
                objWorksheet.Cells(9, "A") = Trim(MCO3!GRPCoAdd1)
                objWorksheet.Cells(10, "A") = Trim(MCO3!GRPCoAdd2)
                objWorksheet.Cells(11, "A") = Trim(MCO3!GRPCoAdd3)
                objWorksheet.Cells(13, "A") = "ATTENTION: " & " " & Trim(MCO3!GRPCoContact)
                objWorksheet.Cells(15, "A") = "TEL :" & " " & Trim(MCO3!GRPCoPhone)
                objWorksheet.Cells(15, "C") = "FAX :" & " " & Trim(MCO3!GRPCoFax)
                If Len(MCO3!MBMDepartment) Then
                    objWorksheet.Cells(7, "A") = Trim(MCO3!MBMDepartment)
                'Else
                    'objWorksheet.Cells(16, "A") = "DEPT :"
                End If
            End If
            
            Set objRange = objWorksheet.Rows(5)
            objRange.Font.Size = "13"
            objRange.Font.Bold = True
            objWorksheet.Cells(5, "G") = "Invoice"
            
            objWorksheet.Cells(8, "G") = "Number :"
            objWorksheet.Cells(8, "H") = Trim(MCO3!MBMCInvoiceNo)
            objWorksheet.Cells(10, "G") = "Date :"
            objWorksheet.Cells(10, "H") = Format(Date, "dd/mm/yyyy")
            objWorksheet.Cells(12, "G") = "Terms :"
            objWorksheet.Cells(12, "H") = "30 DAYS"
            objWorksheet.Cells(19, "A") = "ITEM"
            objWorksheet.Cells(18, "B") = "BORDEREAUX"
            objWorksheet.Cells(19, "B") = "DATE"
            objWorksheet.Cells(19, "D") = "PLAN TYPE"
            objWorksheet.Cells(19, "E") = "ADULT/CHILD"
            objWorksheet.Cells(19, "F") = "QTY"
            objWorksheet.Cells(18, "G") = "COVERED"
            objWorksheet.Cells(19, "G") = "PERSONS"
            objWorksheet.Cells(19, "H") = "AMOUNT"
            
            Do Until MCO.EOF
                
                objWorksheet.Cells(21 + CDbl(i), "A") = i
                objWorksheet.Cells(21 + CDbl(i), "B") = Format(MCO!MBMBordxDate, "mm/dd/yyyy")
                If Trim(MCO!InsType) = "I" Then
                    objWorksheet.Cells(21 + CDbl(i), "D") = "I - INDIVIDUAL"
                ElseIf Trim(MCO!InsType) = "F" Then
                    objWorksheet.Cells(21 + CDbl(i), "D") = "F - FAMILY"
                ElseIf Trim(MCO!InsType) = "G" Then
                    objWorksheet.Cells(21 + CDbl(i), "D") = "G - GROUP INDIVIDUAL"
                ElseIf Trim(MCO!InsType) = "H" Then
                    objWorksheet.Cells(21 + CDbl(i), "D") = "H - GROUP FAMILY"
                End If
                
                If Trim(MCO!MBMAdultChild) = "A" Then
                    objWorksheet.Cells(21 + CDbl(i), "E") = "Adult"
                ElseIf Trim(MCO!MBMAdultChild) = "C" Then
                    objWorksheet.Cells(21 + CDbl(i), "E") = "Child"
                End If
                objWorksheet.Cells(21 + CDbl(i), "F") = IIf(IsNumeric(MCO!TotalPrincipal) = True, MCO!TotalPrincipal, 0)
                TotalPrin = CDbl(TotalPrin) + IIf(IsNumeric(MCO!TotalPrincipal) = True, MCO!TotalPrincipal, 0)
                objWorksheet.Cells(21 + CDbl(i), "H") = Format(MCO!MBMCLNMCOFee, "#,##0.00")
                
                If Trim(MCO!InsType) = "F" Or Trim(MCO!InsType) = "H" Then
                    
                    If cboINVListingVer = "Invoice - by Payor" Then
                    
                        If cboINVCriteriaSel.ListIndex = 0 Then
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report_Batch Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND MBMBatchNo = " & Trim(MCO!MBMBatchNo) & "" & _
                                      " AND PAYCode = '" & Trim(MCO3!PAYCode) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "5"
                            
                        ElseIf cboINVCriteriaSel.ListIndex = 1 Then
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate >= '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND MBMBordxdate <= '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND PAYCode = '" & Trim(MCO3!PAYCode) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "6"
                           
                        Else
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND PAYCode = '" & Trim(MCO3!PAYCode) & "' AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "6"
                                                        
                        End If
                        
                    ElseIf cboINVListingVer = "Invoice - by Group Company" Then
                    
                        If cboINVCriteriaSel.ListIndex = 0 Then
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report_Batch Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND MBMBatchNo = " & Trim(MCO3!MBMBatchNo) & "" & _
                                      " AND GrpCoCode = '" & Trim(MCO3!GRPCoCode) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "5"
                                      
                        ElseIf cboINVCriteriaSel.ListIndex = 1 Then
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND GrpCoCode = '" & Trim(MCO3!GRPCoCode) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "6"
                                            
                        Else
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND GrpCoCode = '" & Trim(MCO3!GRPCoCode) & "' AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "6"
                
                        End If
                        
                    ElseIf cboINVListingVer = "Invoice - by Department" Then
                    
                        If cboINVCriteriaSel.ListIndex = 0 Then
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report_Batch Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND MBMBatchNo = " & Trim(MCO3!MBMBatchNo) & "" & _
                                      " AND GrpCoCode = '" & Trim(MCO3!GRPCoCode) & "' AND MBMDepartment = '" & Trim(MCO3!MBMDepartment) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "5"
                
                        ElseIf cboINVCriteriaSel.ListIndex = 1 Then
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND GrpCoCode = '" & Trim(MCO3!GRPCoCode) & "' AND MBMDepartment = '" & Trim(MCO3!MBMDepartment) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "6"
                
                        Else
                            'strsql3 = " Select * From VIEW_ClinicalInvoice_Supp_Report Where INSType = '" & Trim(MCO!InsType) & "'" & _

                            strsql3 = " AND INSType = '" & Trim(MCO!InsType) & "'" & _
                                      " AND MBMBordxdate = '" & Format(MCO!MBMBordxDate, "mm/dd/yyyy") & "'" & _
                                      " AND GrpCoCode = '" & Trim(MCO3!GRPCoCode) & "' AND MBMDepartment = '" & Trim(MCO3!MBMDepartment) & "' AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
                                      
                            strsql3 = strsql3 '+ strsql4
                            strAppendCase = "6"
                        End If
                    End If
                                        
                    'MCO2.Open strsql3, MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic
                                
                        Call GetSuppMCOFee(strsql3, strAppendCase, TotalPrin, MCO!MBMCLNMCOFee)
                        'SubSupp = CDbl(TotalPrin) + IIf(IsNumeric(MCO2!CQty) = True, MCO2!CQty, 0)
                        objWorksheet.Cells(21 + CDbl(i), "G") = GetSubSupp
                        TotalSupp = CDbl(TotalSupp) + objWorksheet.Cells(21 + CDbl(i), "G")
                        
                        'SuppMCO = CDbl(MCO!MBMCLNMCOFee) + IIf(IsNumeric(MCO2!MBMCMCO) = True, MCO2!MBMCMCO, 0)
                        objWorksheet.Cells(21 + CDbl(i), "H") = GetSuppMCO
                        TotalSuppMCO = CDbl(TotalSuppMCO) + objWorksheet.Cells(21 + CDbl(i), "H")
                Else
                    
                    SubSupp = CDbl(SubSupp) + CDbl(TotalPrin)
                    objWorksheet.Cells(21 + CDbl(i), "G") = SubSupp
                    TotalSupp = CDbl(TotalSupp) + objWorksheet.Cells(21 + CDbl(i), "G")
                    
                End If
                            
            MCO.MoveNext
            MCO3.MoveNext
            
            i = i + 1
            TotalPrin = "0"
            SubSupp = "0"
            TotalSupp = "0"
            SuppMCO = "0"
            TotalSuppMCO = "0"
            
            Loop
            MCO.Close
            Set MCO = Nothing
            MCO3.Close
            Set MCO3 = Nothing
            
            
            Set objRange = objWorksheet.Rows(42)
            objRange.Font.Bold = True
                        
            objWorksheet.Cells(42, "F") = "'-------------------------------------------------------"
            objWorksheet.Cells(43, "D") = "TOTAL"
            objWorksheet.Range("F43").Formula = "=SUM(F22:F41)"
            objWorksheet.Range("G43").Formula = "=SUM(G22:G41)"
            objWorksheet.Range("H43").Formula = "=SUM(H22:H41)"
            objWorksheet.Cells(44, "F") = "'======================================="
            
            objWorksheet.Cells(49, "A") = "_________________________________________"
            objWorksheet.Cells(50, "A") = "For United Health Connnect Sdn. Bhd."
            objWorksheet.Cells(57, "A") = "Interest of 1% per annum will be levied after expiry of terms of payment."
            
            
            End If
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            Screen.MousePointer = vbDefault
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
End Sub

Private Sub cmdINVSearch_Click()
Dim errorFound As Boolean
errorFound = True

txtINVTotalPrincipal = ""
txtINVTotalSupplementary1 = ""

If cboINVCriteriaSel.ListIndex = 0 Then
    If cboINVListingVer.ListIndex = 0 Then
        If txtINVBordxDate = "__/__/____" Then
            MsgBox "Please Enter Bordereaux Date!", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf txtINVBatchNo = "" Then
            MsgBox "Please Enter Batch Number!", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf cboINVPayor = "" Then
            MsgBox "Please Select a Payor!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            errorFound = False
        End If
    ElseIf cboINVListingVer.ListIndex = 1 Then
        If txtINVBordxDate = "__/__/____" Then
            MsgBox "Please Enter Bordereaux Date!", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf txtINVBatchNo = "" Then
            MsgBox "Please Enter Batch Number!", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf cboINVGrpCom = "" Then
            MsgBox "Please Select a Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            errorFound = False
        End If
    ElseIf cboINVListingVer.ListIndex = 2 Then
        If txtINVBordxDate = "__/__/____" Then
            MsgBox "Please Enter Bordereaux Date!", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf txtINVBatchNo = "" Then
            MsgBox "Please Enter Batch Number!", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf cboINVGrpCom = "" Then
            MsgBox "Please Select a Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf CboDepartment = "" Then
            MsgBox "Please Select a Department!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            errorFound = False
        End If
    End If
ElseIf cboINVCriteriaSel.ListIndex = 1 Then
    If cboINVListingVer.ListIndex = 0 Then
        If txtINVBordxDateFr = "__/__/____" Or txtINVBordxDateTo = "__/__/____" Then
            MsgBox "Please enter valid Bordereaux Range !", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf cboINVPayor = "" Then
            MsgBox "Please Select a Payor!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            errorFound = False
        End If
    ElseIf cboINVListingVer.ListIndex = 1 Then
        If txtINVBordxDateFr = "__/__/____" Or txtINVBordxDateTo = "__/__/____" Then
            MsgBox "Please enter valid Bordereaux Range !", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf cboINVGrpCom = "" Then
            MsgBox "Please Select a Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            errorFound = False
        End If
    ElseIf cboINVListingVer.ListIndex = 2 Then
        If txtINVBordxDateFr = "__/__/____" Or txtINVBordxDateTo = "__/__/____" Then
            MsgBox "Please enter valid Bordereaux Range !", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf cboINVGrpCom = "" Then
            MsgBox "Please Select a Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf CboDepartment = "" Then
            MsgBox "Please Select a Department!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            errorFound = False
        End If
    End If
ElseIf cboINVCriteriaSel.ListIndex = 2 Then
    If txtINVInvoiceNo = "" Then
        MsgBox "Please enter Invoice Number!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        errorFound = False
    End If
End If

If errorFound = False Then
    TotalSupplementary = 0
    txtINVTotalPrincipal = 0
    cmdINVClose.Enabled = False
    RecordSearching_INV lvwInvoice
    cmdINVClose.Enabled = True
End If

End Sub

Private Sub INVSuppListing()
'On Error Resume Next
Dim strsql As String
Dim sql As String
Dim strsqls As String
Dim rsPrint As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim ListMaster As ListItem
Dim isPrinted As Boolean
Dim Total As Integer
Dim Current As Integer
Dim tempMBMCNumber, tempCoverID
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    If cboINVListingVer = "Invoice - by Payor" Then
        Select Case cboINVCriteriaSel.ListIndex
            Case 0
                    sDate = CStr(Month(txtINVBordxDate)) & "/" & CStr(Day(txtINVBordxDate)) & "/" & CStr(Year(txtINVBordxDate))
                    'strsql = "SELECT * FROM VIEW_Clinic_Invoice_ReportSearchSupp "
                    'strsql = strsql & " WHERE MBMBordxDate = '" & sDate & "'"
                    strsql = strsql & " AND MBMBordxDate >= '" & DateFrom & "'"
                    strsql = strsql & " AND MBMBatchNo = '" & Trim(txtINVBatchNo) & "'"
                    strsql = strsql & " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    strsql = strsql & " ORDER BY MBMCCoverID"
                    
                    'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
            Case 1
                    DateFrom = CStr(Month(txtINVBordxDateFr)) & "/" & CStr(Day(txtINVBordxDateFr)) & "/" & CStr(Year(txtINVBordxDateFr))
                    DateTo = CStr(Month(txtINVBordxDateTo)) & "/" & CStr(Day(txtINVBordxDateTo)) & "/" & CStr(Year(txtINVBordxDateTo))
                    'strsql = "SELECT * FROM VIEW_Clinic_Invoice_ReportSearchSupp "
                    'strsql = strsql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                    strsql = strsql & " AND MBMBordxDate >= '" & DateFrom & "'"
                    strsql = strsql & " AND MBMBordxDate <= '" & DateTo & "'"
                    strsql = strsql & " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    strsql = strsql & " ORDER BY MBMCCoverID"
                    
                    'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        
            Case 2
                    'strsql = "SELECT * FROM VIEW_Clinic_Invoice_ReportSearchSupp "
                    'strsql = strsql & " WHERE MBMCInvoiceNo = '" & Trim(INVNo) & "'"
                    strsql = strsql & " AND MBMCInvoiceNo = '" & Trim(InvNo) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    
                    'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                          
        End Select
        
    ElseIf cboINVListingVer = "Invoice - by Group Company" Then
        Select Case cboINVCriteriaSel.ListIndex
            Case 0
                    sDate = CStr(Month(txtINVBordxDate)) & "/" & CStr(Day(txtINVBordxDate)) & "/" & CStr(Year(txtINVBordxDate))
                    'strsql = "SELECT * FROM VIEW_Clinic_Invoice_ReportSearchSupp "
                    'strsql = strsql & " WHERE MBMBordxDate = '" & sDate & "'"
                    strsql = strsql & " AND MBMBordxDate = '" & sDate & "'"
                    strsql = strsql & " AND MBMBatchNo = '" & Trim(txtINVBatchNo) & "'"
                    strsql = strsql & " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    strsql = strsql & " ORDER BY MBMCCoverID"
                    
                    'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
            Case 1
                    DateFrom = CStr(Month(txtINVBordxDateFr)) & "/" & CStr(Day(txtINVBordxDateFr)) & "/" & CStr(Year(txtINVBordxDateFr))
                    DateTo = CStr(Month(txtINVBordxDateTo)) & "/" & CStr(Day(txtINVBordxDateTo)) & "/" & CStr(Year(txtINVBordxDateTo))
                    'strsql = "SELECT * FROM VIEW_Clinic_Invoice_ReportSearchSupp "
                    'strsql = strsql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                    strsql = strsql & " AND MBMBordxDate >= '" & DateFrom & "'"
                    strsql = strsql & " AND MBMBordxDate <= '" & DateTo & "'"
                    strsql = strsql & " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    strsql = strsql & " ORDER BY MBMCCoverID"
                    
                    'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
            Case 2
                    'strsql = "SELECT * FROM VIEW_Clinic_Invoice_ReportSearchSupp "
                    'strsql = strsql & " WHERE MBMCInvoiceNo = '" & Trim(INVNo) & "'"
                    strsql = strsql & " AND MBMCInvoiceNo = '" & Trim(InvNo) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    
                    'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                          
        End Select
    
    ElseIf cboINVListingVer = "Invoice - by Department" Then
        Select Case cboINVCriteriaSel.ListIndex
            Case 0
                    sDate = CStr(Month(txtINVBordxDate)) & "/" & CStr(Day(txtINVBordxDate)) & "/" & CStr(Year(txtINVBordxDate))
                    'strsql = "SELECT * FROM VIEW_Clinic_Invoice_ReportSearchSupp "
                    'strsql = strsql & " WHERE MBMBordxDate = '" & sDate & "'"
                    strsql = strsql & " AND MBMBordxDate = '" & sDate & "'"
                    strsql = strsql & " AND MBMBatchNo = '" & Trim(txtINVBatchNo) & "'"
                    strsql = strsql & " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "'"
                    strsql = strsql & " AND MBMDepartment = '" & Trim(CboDepartment) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    strsql = strsql & " ORDER BY MBMCCoverID"
                    
                    'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
            Case 1
                    DateFrom = CStr(Month(txtINVBordxDateFr)) & "/" & CStr(Day(txtINVBordxDateFr)) & "/" & CStr(Year(txtINVBordxDateFr))
                    DateTo = CStr(Month(txtINVBordxDateTo)) & "/" & CStr(Day(txtINVBordxDateTo)) & "/" & CStr(Year(txtINVBordxDateTo))
                    'strsql = "SELECT * FROM VIEW_Clinic_Invoice_ReportSearchSupp "
                    'strsql = strsql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                    strsql = strsql & " AND MBMBordxDate >= '" & DateFrom & "'"
                    strsql = strsql & " AND MBMBordxDate <= '" & DateTo & "'"
                    strsql = strsql & " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "'"
                    strsql = strsql & " AND MBMDepartment = '" & Trim(CboDepartment) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    strsql = strsql & " order by MBMCCoverID"
                    
                    'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
            Case 2
                    'strsql = "SELECT * FROM VIEW_Clinic_Invoice_ReportSearchSupp "
                    'strsql = strsql & " WHERE MBMCInvoiceNo = '" & Trim(INVNo) & "'"
                    strsql = strsql & " AND MBMCInvoiceNo = '" & Trim(InvNo) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    
                    'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                          
        End Select
    
    End If
    
    strAppendCase = "2"
            
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "MCOFeeModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    MBMCov.CursorLocation = adUseClient
    MBMCov.CursorType = adOpenForwardOnly
    MBMCov.CacheSize = 100
    Set MBMCov = cmd.Execute
    

               Do Until MBMCov.EOF
                   If tempMBMCNumber <> MBMCov!MBMCNUmber Then
                        Set ListMaster = lvwInvoice.ListItems.Add(, , PAYCode)
                         ListMaster.SubItems(1) = Trim(MBMCov!MBMCNUmber)
                         ListMaster.SubItems(2) = MBMCov!MBMCCoverID
                         ListMaster.SubItems(3) = MBMCov!MBMCRELCode
                         ListMaster.SubItems(4) = lvwInvoice.SelectedItem.SubItems(4)
                         ListMaster.SubItems(5) = Trim(lvwInvoice.SelectedItem.SubItems(5))
                         ListMaster.SubItems(6) = lvwInvoice.SelectedItem.SubItems(6)
                         ListMaster.SubItems(7) = lvwInvoice.SelectedItem.SubItems(7)
                         ListMaster.SubItems(8) = Trim(MBMCov!MBMCName)
                     
                        If MBMCov!MBMGroupCompany = "" Then
                             ListMaster.SubItems(9) = ""
                         Else
                             If Len(MBMCov!MBMGroupCompany) Then
                                 ListMaster.SubItems(9) = Trim(MBMCov!MBMGroupCompany)
                             End If
                         End If
                     
                        If Len(MBMCov!MBMCInvoiceNo) Then
                             ListMaster.SubItems(11) = Trim(MBMCov!MBMCInvoiceNo)
                         End If
                     
                        TotalSupplementary = TotalSupplementary + 1
                        txtINVTotalSupplementary1 = TotalSupplementary
 
                            itemToPrint = INV_PRINTED
                            cmdINVPrint.Enabled = True
                            cmdINVPrintPreview.Enabled = True
 
                        
                        If Printed = "YES" Then
                            ListMaster.SubItems(itemToPrint) = "YES"
                        Else
                            ListMaster.SubItems(itemToPrint) = "NO"
                        End If
                        
                        tempMBMCNumber = MBMCov!MBMCNUmber
                        tempCoverID = MBMCov!MBMCCoverID
                     Else
                        If tempCoverID <> MBMCov!MBMCCoverID Then
                         Set ListMaster = lvwInvoice.ListItems.Add(, , PAYCode)
                         ListMaster.SubItems(1) = Trim(MBMCov!MBMCNUmber)
                         ListMaster.SubItems(2) = MBMCov!MBMCCoverID
                         ListMaster.SubItems(3) = MBMCov!MBMCRELCode
                         ListMaster.SubItems(4) = lvwInvoice.SelectedItem.SubItems(4)
                         ListMaster.SubItems(5) = Trim(lvwInvoice.SelectedItem.SubItems(5))
                         ListMaster.SubItems(6) = lvwInvoice.SelectedItem.SubItems(6)
                         ListMaster.SubItems(7) = lvwInvoice.SelectedItem.SubItems(7)
                         ListMaster.SubItems(8) = Trim(MBMCov!MBMCName)
                     
                        If MBMCov!MBMGroupCompany = "" Then
                             ListMaster.SubItems(9) = ""
                         Else
                             If Len(MBMCov!MBMGroupCompany) Then
                                 ListMaster.SubItems(9) = Trim(MBMCov!MBMGroupCompany)
                             End If
                         End If
                     
                        If Len(MBMCov!MBMCInvoiceNo) Then
                             ListMaster.SubItems(11) = Trim(MBMCov!MBMCInvoiceNo)
                         End If
                     
                        TotalSupplementary = TotalSupplementary + 1
                        txtINVTotalSupplementary1 = TotalSupplementary
                        
                            itemToPrint = INV_PRINTED
                            cmdINVPrint.Enabled = True
                            cmdINVPrintPreview.Enabled = True
                                                
                        If Printed = "YES" Then
                            ListMaster.SubItems(itemToPrint) = "YES"
                        Else
                            ListMaster.SubItems(itemToPrint) = "NO"
                        End If
                        
                        tempMBMCNumber = Trim(MBMCov!MBMCNUmber)
                        tempCoverID = MBMCov!MBMCCoverID
                        End If
                    End If
                    MBMCov.MoveNext
                   Loop
            
                MBMCov.Close
                Set MBMCov = Nothing

End Sub

Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim sqlstr As String
Dim GrpCom As New ADODB.Recordset
Dim sqlstr2 As String
   
    sqlstr = "select PAYCode, PAYCompanyName from PAY "
    
    PAY.Open sqlstr, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    Do While Not PAY.EOF
        cboINVPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    
    PAY.Close
    Set PAY = Nothing
    
    sqlstr = "select GRPCoCode , GRPCoName from GRPCompany"
    
    GrpCom.Open sqlstr, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    Do While Not GrpCom.EOF
    
        cboINVGrpCom.AddItem Trim((GrpCom!GRPCoCode)) & " - " & Trim((GrpCom!GrpCoName))
        GrpCom.MoveNext
    
    Loop
    
    GrpCom.Close
    Set GrpCom = Nothing
    
End Sub


'Private Sub cmdMCOClose_Click()
'Dim RecordExit
    'RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
    'If RecordExit = vbYes Then
        'bExit = True
        'Unload Me
    'End If
'End Sub

Private Sub Form_Load()
    Call ComboListing
    
       'X = CreatePath("c:\cardPrinting")
        'If X = False Then
            'MsgBox "Error Creating Folder c:\CardPrinting !" & vbNewLine & _
                    " Please create that folder manually to store the output of leaflet in sofe copy!"
        'End If
End Sub

Private Sub lvwInvoice_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwInvoice.SortKey = ColumnHeader.Index - 1
    lvwInvoice.Sorted = True
    If lvwInvoice.SortOrder = lvwAscending Then
       lvwInvoice.SortOrder = lvwDescending
    Else
       lvwInvoice.SortOrder = lvwAscending
    End If
End Sub

Private Sub UpdateINVPrintingFile(MemberNo As String)
Dim sql As String
Dim strsql As String
Dim InvoiceDate As String

    
    'Update file for MBMOthersCLN Table
    InvoiceDate = Format(Date, "YYYY/MM/DD")
    
    sql = " update MBMCLNOthers set " & _
             " MBMCInvoiceDate = '" & InvoiceDate & "', " & _
             " MBMCInvoiceNo = '" & Trim(txtInvoiceNo) & "', " & _
             " MBMCInvoiceBy = '" & Logon & "', MBMCInvoiced = 'Y' " & _
             " Where MBMNumber = '" & Trim(MemberNo) & "'"
    
    medixmasterconn.Execute sql
    
    'update file for MBMCoveredPersonsCLN Table
    strsql = " update MBMCoveredPersonsCLN set " & _
             " MBMCInvoiceDate = '" & InvoiceDate & "', " & _
             " MBMCInvoiceNo = '" & Trim(txtInvoiceNo) & "', " & _
             " MBMCInvoiceBy = '" & Logon & "' " & _
             " Where MBMCNumber = '" & Trim(MemberNo) & "'"
    
    medixmasterconn.Execute strsql
    
End Sub

Private Sub IncInvoiceNo()
Dim MBMMCOINVQ As New ADODB.Recordset
Dim strsql As String
Dim InvoiceNo As String
    
    strsql = "Select * from CLNMCOINVQ "
    MBMMCOINVQ.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If MBMMCOINVQ.RecordCount = 0 Then
        With MBMMCOINVQ
            .AddNew
                !MCOInvoice_No = "00001"
                InvoiceNo = "00001"
            .Update
        End With
    Else
    
        With MBMMCOINVQ
                !MCOInvoice_No = Format(!MCOInvoice_No + 1, "00000")
                InvoiceNo = Format(!MCOInvoice_No, "00000")
        .Update
        End With
    End If

    txtInvoiceNo = "I" & InvoiceNo
    
    MBMMCOINVQ.Close
    Set MBMMCOINVQ = Nothing
            
End Sub

Private Sub GetSuppMCOFee(strsql, strAppendCase, TotalPrin, MBMCLNMCOFee)
Dim cmd As New ADODB.Command
Dim MCO2 As New ADODB.Recordset
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter

    GetSubSupp = 0
    GetSuppMCO = 0
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "MCOFeeModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    MCO2.CursorLocation = adUseClient
    MCO2.CursorType = adOpenForwardOnly
    MCO2.CacheSize = 100
    Set MCO2 = cmd.Execute
    
    If Not MCO2.EOF Then
        GetSubSupp = CDbl(TotalPrin) + IIf(IsNumeric(MCO2!CQty) = True, MCO2!CQty, 0)
        GetSuppMCO = CDbl(MBMCLNMCOFee) + IIf(IsNumeric(MCO2!MBMCMCO) = True, MCO2!MBMCMCO, 0)
    End If
    MCO2.Close
    Set MCO2 = Nothing
    
End Sub

