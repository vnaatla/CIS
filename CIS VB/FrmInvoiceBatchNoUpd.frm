VERSION 5.00
Begin VB.Form FrmInvoiceBatchNoUpd 
   Caption         =   "Invoice Batch Number Transfer"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8088
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3240
   ScaleWidth      =   8088
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   600
      Left            =   0
      TabIndex        =   19
      Top             =   2640
      Width           =   8055
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   6240
         TabIndex        =   12
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   300
         Left            =   7080
         TabIndex        =   13
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Transfer"
         Height          =   300
         Left            =   5280
         TabIndex        =   11
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   0
      TabIndex        =   18
      Top             =   960
      Width           =   8055
      Begin VB.TextBox txtNOofCas1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtNOofCas2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtInvAmt1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtInvAmt2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6840
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtInvNo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtClinic1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   26
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtInvNo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   6
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtClinic2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   28
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Invoice Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5880
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Clinic Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   27
         Top             =   760
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Final Batch No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Invoice Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblFinalBatchNo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   1335
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         X1              =   0
         X2              =   8040
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label4 
         Caption         =   "Total Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   5520
         TabIndex        =   23
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Total Case(s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblFinalTotAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Left            =   6480
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblFinalCases 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   4320
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Invoice Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   525
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "No. of Case(s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   260
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   360
      Width           =   8055
      Begin VB.TextBox txtBatchNo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5520
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtBatchNo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Transfer --->"
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
         Left            =   3960
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Batch Number "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Invoice Batch Number Transfer"
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
      TabIndex        =   14
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "FrmInvoiceBatchNoUpd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClear_Click()
    Call ClearEntries
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If

End Sub
Private Sub Form_Load()
Const METHOD_NAME = "Form_Load"
On Error GoTo debugs

Call ActionLogUpdate(12, "FrmInvoiceBatchNoUpd")


debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdSave_Click()
Const METHOD_NAME = "CmdSave_Click"
On Error GoTo debugs
Dim RecordSaved
    If Trim(txtBatchNo1) = "" Or Trim(txtBatchNo2) = "" Then
        MsgBox "Please confirm your Batch Number ! ", vbCritical + vbOKOnly
    ElseIf Trim(txtInvNo1) <> Trim(txtInvNo2) Then
        MsgBox "Different Invoice No. Must be SAME Invoice No. ! ", vbCritical + vbOKOnly
    'ElseIf Trim(txtClinic1) <> Trim(txtClinic2) Then
        'MsgBox "Different Clinic. Must be SAME Clinic. ! ", vbCritical + vbOKOnly
    ElseIf Trim(txtBatchNo1) = Trim(txtBatchNo2) Then
        MsgBox "Same Batch Number ! ", vbCritical + vbOKOnly
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"

            Call BatchTransProcess
            
            'MediConnCISDB.Close
            'Set MediConnCISDB = Nothing
            'medixmasterconn.Close
            'Set medixmasterconn = Nothing
            'MediConnCISWS.Close
            'Set MediConnCISWS = Nothing
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub txtBatchNo1_LostFocus()
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"

        txtBatchNo1 = ChkStr(txtBatchNo1)
        
        If Trim(txtBatchNo1) <> "" Then
            Call ShowInfo(txtBatchNo1, txtNOofCas1, txtInvNo1, txtInvAmt1, txtClinic1)
        End If
        Call FinalBatchCal

    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
    

End Sub

Private Sub ShowInfo(BatchNo As TextBox, SNoOfCase As TextBox, SInvNo As TextBox, SInvAmt As TextBox, SCLNName As TextBox)
Const METHOD_NAME = "ShowInfo"
On Error GoTo debugs
Dim tmpNoofCAS As Integer
Dim tmpINVAmt As Double
Dim cmd2 As New ADODB.Command
Dim CASRec As New ADODB.Recordset
Dim prm2 As New ADODB.Parameter
Dim tmpCnt As Integer
Dim tmpTOtCas As Integer
Dim tmpInvNo As String
Dim FirstInvNo As String

    Set cmd2.ActiveConnection = MediConnCISDB
    cmd2.CommandText = "SP_SeachCasByOPBatch"
    cmd2.CommandType = adCmdStoredProc
    cmd2.CommandTimeout = 300
    Set prm2 = cmd2.CreateParameter("tmpBatchNo", adVarChar, adParamInput, 20, Trim(BatchNo))
    cmd2.Parameters.Append prm2
    CASRec.CursorLocation = adUseClient
    CASRec.CursorType = adOpenForwardOnly
    CASRec.CacheSize = 100
    Set CASRec = cmd2.Execute
    tmpCnt = 0
    tmpTOtCas = 0
    SInvAmt = 0
    If CASRec.EOF Then
        MsgBox "Invalid OP Batch Number ! ", vbCritical + vbOKOnly
        BatchNo = ""
        BatchNo.SetFocus
    Else
        Do While Not CASRec.EOF
            tmpCnt = tmpCnt + 1
            tmpInvNo = IIf(Len(CASRec!InvoiceNo), Trim(CASRec!InvoiceNo), "")
            If tmpCnt = 1 Then
                FirstInvNo = tmpInvNo
                SInvNo = tmpInvNo
            End If
            If Trim(FirstInvNo) <> Trim(tmpInvNo) Then
                SInvNo = SInvNo & "," & tmpInvNo
            End If
            
            SInvAmt = CDbl(SInvAmt) + Format(IIf(IsNumeric(CASRec!InvoiceStmtAmt), CDbl(CASRec!InvoiceStmtAmt), 0), "#,##0.00")
            SCLNName = IIf(Len(CASRec!CLNName), Trim(CASRec!CLNName), "")
            tmpTOtCas = tmpTOtCas + CASRec!TotNoOfCase
            SNoOfCase = tmpTOtCas
            CASRec.MoveNext
        Loop
    End If
    CASRec.Close
    Set CASRec = Nothing
    If tmpCnt > 1 Then
        MsgBox "Invoice No : " & SInvNo & ", Different Clinic Found !", vbOKOnly
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub txtBatchNo2_LostFocus()

   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
        txtBatchNo2 = ChkStr(txtBatchNo2)
        If Trim(txtBatchNo2) <> "" Then
            Call ShowInfo(txtBatchNo2, txtNOofCas2, txtInvNo2, txtInvAmt2, txtClinic2)
        End If
        Call FinalBatchCal

  '  MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  MediConnCISWS.Close
  '  Set MediConnCISWS = Nothing

End Sub

Private Sub FinalBatchCal()
Dim TotCas1, TotCas2 As Integer
Dim totInvAmt1, totInvAmt2 As Double
    TotCas1 = IIf(IsNumeric(txtNOofCas1), txtNOofCas1, 0)
    TotCas2 = IIf(IsNumeric(txtNOofCas2), txtNOofCas2, 0)
    totInvAmt1 = IIf(IsNumeric(txtInvAmt1), txtInvAmt1, 0)
    totInvAmt2 = IIf(IsNumeric(txtInvAmt2), txtInvAmt2, 0)
    lblFinalBatchNo = txtBatchNo2
    lblFinalCases = TotCas1 + TotCas2
    lblFinalTotAmt = Format(totInvAmt1 + totInvAmt2, "#,##0.00")
End Sub

Private Sub ClearEntries()
Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Name <> "txtMBMNo" And ctl.Name <> "txtInvNo" And ctl.Name <> "TxtClinicCode" _
            And ctl.Name <> "txtNPClinic" And ctl.Name <> "TxtTotInvoiceAmt" And ctl.Name <> "txtClinicName" Then
               ctl.Text = ""
            End If
       ElseIf TypeOf ctl Is MaskEdBox Then
            If ctl.Name <> "txtInvDate" And ctl.Name <> "txtInvRecDate" Then
                ctl.mask = ""
                ctl.Text = ""
                ctl.mask = "##/##/####"
            End If
            If ctl.Name = "txtVisitTime" Then
                ctl.mask = "##:##"
            End If
        End If
    Next ctl
    lblFinalBatchNo = ""
    lblFinalCases = ""
    lblFinalTotAmt = ""
    txtBatchNo1.SetFocus
End Sub

Private Sub BatchTransProcess()
Const METHOD_NAME = "BatchTransProcess"
On Error GoTo debugs
Dim CompTrans As Boolean
Dim UpdBatch As String
Dim FinalAmt As Double
Dim UpdInvAmt As String
Dim TransRemark As String
    
    FinalAmt = IIf(IsNumeric(lblFinalTotAmt), lblFinalTotAmt, 0)
    On Err GoTo ErrorSave
    Screen.MousePointer = vbHourglass
    CompTrans = False
    MediConnCISDB.beginTrans
    'BatchTransRemarks
    TransRemark = "Transfer By : " & Logon & " " & Trim(txtBatchNo1) & "->" & Trim(txtBatchNo2) & " , " & Format(Date, "yyyymmdd")
    UpdBatch = "UPDATE CISCAS SET OPBatchNo = '" & Trim(txtBatchNo2) & _
               "', BatchTransRemarks = '" & Trim(TransRemark) & _
               "' WHERE OPBatchNo ='" & Trim(txtBatchNo1) & "'"
    MediConnCISDB.Execute UpdBatch
    UpdInvAmt = "UPDATE CISCAS SET InvoiceStmtAmt = " & FinalAmt & _
            " WHERE OPBatchNo ='" & Trim(txtBatchNo2) & "'"
    MediConnCISDB.Execute UpdInvAmt
    'Call ReSequence
    MediConnCISDB.commitTrans
    CompTrans = True
    Screen.MousePointer = Default
    MsgBox "Records Transfered Successfully ! ", vbOKOnly
    Call ClearEntries
    Exit Sub
ErrorSave:
    If CompTrans = False Then
        MediConnCISDB.RollbackTrans
    End If
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ReSequence()
Const METHOD_NAME = "ReSequence"
On Error GoTo debugs
Dim StrBatchSql As String
Dim BATSeqRec As New ADODB.Recordset
Dim NewSeqNo As String

    StrBatchSql = "Select OPBatchSeq From OPBatch WHERE OPBatchSeq='" & Right(txtBatchNo1, 5) & "'"
    BATSeqRec.Open StrBatchSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do While Not BATSeqRec.EOF
        NewSeqNo = BATSeqRec!OPBAtchSeq - 1
        BATSeqRec!OPBAtchSeq = Format(NewSeqNo, "00000")
        BATSeqRec.update
        BATSeqRec.MoveNext
    Loop
    BATSeqRec.Close
    Set BATSeqRec = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
