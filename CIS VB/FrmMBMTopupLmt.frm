VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmMBMTopupLmt 
   Caption         =   "Member Topup Limit"
   ClientHeight    =   3528
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   7032
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3528
   ScaleWidth      =   7032
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   50
      TabIndex        =   19
      Top             =   240
      Width           =   6855
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
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
         Left            =   6000
         TabIndex        =   5
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
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
         Left            =   5280
         TabIndex        =   4
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdPreview 
         Caption         =   "&Prev"
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
         Left            =   4560
         TabIndex        =   3
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "&Go"
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
         Left            =   3840
         TabIndex        =   2
         Top             =   180
         Width           =   735
      End
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   1
         Top             =   160
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Member No"
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
         Top             =   200
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2040
      Left            =   50
      TabIndex        =   20
      Top             =   720
      Width           =   6855
      Begin VB.OptionButton optSPlmt 
         Caption         =   "SP"
         Height          =   195
         Left            =   960
         TabIndex        =   40
         Top             =   1440
         Width           =   615
      End
      Begin VB.OptionButton optGPlmt 
         Caption         =   "GP"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   200
         Width           =   5655
      End
      Begin VB.TextBox txtGrpComp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   435
         Width           =   5655
      End
      Begin VB.TextBox txtPlan 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   675
         Width           =   1095
      End
      Begin VB.TextBox txtInsType 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   33
         Top             =   675
         Width           =   1095
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   915
         Width           =   3375
      End
      Begin VB.TextBox txtICNo 
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   12
         Top             =   1155
         Width           =   3375
      End
      Begin VB.TextBox txtTopupAmt 
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtBalLmt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtStatus 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   36
         Top             =   675
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtEffDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   10
         Top             =   915
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   508
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtExpDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   11
         Top             =   1155
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   508
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtNewBal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   15
         Top             =   1395
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Status"
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
         Index           =   14
         Left            =   4680
         TabIndex        =   37
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ins Type"
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
         Index           =   12
         Left            =   2400
         TabIndex        =   34
         Top             =   700
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "IC Number"
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
         Index           =   11
         Left            =   120
         TabIndex        =   32
         Top             =   1215
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Plan Code"
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
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "New Bal."
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
         Index           =   10
         Left            =   4680
         TabIndex        =   30
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Topup Amt"
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
         Index           =   9
         Left            =   2400
         TabIndex        =   29
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Exp Date"
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
         Index           =   8
         Left            =   4680
         TabIndex        =   28
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Balance Lmt"
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
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Eff Date"
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
         Index           =   5
         Left            =   4680
         TabIndex        =   26
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Department"
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
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Group Comp."
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
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   470
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "MBM Name"
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
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   230
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   660
      Left            =   0
      TabIndex        =   22
      Top             =   2760
      Width           =   6855
      Begin VB.TextBox lblLimitType 
         Height          =   285
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
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
         Left            =   6000
         TabIndex        =   18
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
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
         Left            =   5280
         TabIndex        =   17
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
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
         Left            =   4560
         TabIndex        =   16
         Top             =   180
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Plan Code"
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
      Index           =   13
      Left            =   4680
      TabIndex        =   35
      Top             =   1485
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Member Topup Limit and Amendment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "FrmMBMTopupLmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MembershipNo As String
Dim CovID As String
Dim tmpPrevICNo As String
Dim tmpSptLmtInd As String

Private Sub CmdClear_Click()
    Call ClearEntries
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Public Sub CmdGo_Click()
Const METHOD_NAME = "CmdGo_Click"
On Error GoTo debugs

Dim MBMNumber As String
Dim ZZ As Integer
Dim tmpRecCnt As Integer
    
    MBMNumber = txtMBMNumber
    Call ClearEntries
    txtMBMNumber = MBMNumber
    If Trim(txtMBMNumber) <> "" Then
        ZZ = InStr(1, txtMBMNumber, "-")
        MembershipNo = Trim(Left(txtMBMNumber, IIf(ZZ - 1 <= 0, Len(Trim(txtMBMNumber)), ZZ - 1)))
        If ZZ = 0 Then
            CovID = "00"
        Else
            CovID = Trim(Right(txtMBMNumber, IIf(ZZ - 1 <= 0, 0, Len(Trim(txtMBMNumber)) - ZZ)))
        End If
        tmpRecCnt = 1
        If Len(Trim(txtMBMNumber)) = 8 Or Len(Trim(txtMBMNumber)) = 9 Then
            frmSearchMBM05.Source = 6
            frmSearchMBM05.Show
            tmpRecCnt = frmSearchMBM05.SearchRecCnt
            If frmSearchMBM05.SearchRecCnt < 1 Then
                Unload frmSearchMBM05
            ElseIf frmSearchMBM05.SearchRecCnt = 1 Then
                txtMBMNumber = frmSearchMBM05.tmpMBMNo
                Unload frmSearchMBM05
            End If
        End If
        ZZ = InStr(1, txtMBMNumber, "-")
        MembershipNo = Trim(Left(txtMBMNumber, IIf(ZZ - 1 <= 0, Len(Trim(txtMBMNumber)), ZZ - 1)))
        If ZZ = 0 Then
            CovID = "00"
        Else
            CovID = Trim(Right(txtMBMNumber, IIf(ZZ - 1 <= 0, 0, Len(Trim(txtMBMNumber)) - ZZ)))
        End If
        If tmpRecCnt < 2 Then
            Screen.MousePointer = vbHourglass
            optGPlmt.Value = True
            optSPlmt.Value = False
            Call FindMBM
            Screen.MousePointer = Default
        End If
    Else
        MsgBox "Please Enter Membership No", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMNumber.SetFocus
    End If
    
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If


End Sub

Private Sub CmdNext_Click()
Const METHOD_NAME = "CmdNext_Click"
On Error GoTo debugs

    Dim MBM As New ADODB.Recordset
    Dim sql As String
    sql = " Select MBMNumber from MBM Where HSCLNInd <> 'H' AND MBMNumber > '" & RemStr(MembershipNo) & "' " & _
        " Order by MBMnumber "
    MBM.MaxRecords = 1
    MBM.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If MBM.EOF = False Then
        txtMBMNumber.Text = MBM!MBMNumber & "- 00"
        MBM.Close
        Set MBM = Nothing
        CmdGo_Click
    Else
        MBM.Close
        Set MBM = Nothing
        MsgBox " Last Record Reached! ", vbCritical + vbOKOnly
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdPreview_Click()
Const METHOD_NAME = "CmdPreview_Click"
On Error GoTo debugs

Dim MBM As New ADODB.Recordset
Dim sql As String
            
    sql = " select MBMNumber from MBM WHERE HSCLNInd <> 'H' AND MBMNumber < '" & RemStr(MembershipNo) & "' " & _
          " Order by MBMnumber DESC "
    MBM.MaxRecords = 1
    MBM.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    If MBM.EOF = False Then
        txtMBMNumber.Text = MBM!MBMNumber & "- 00"
        MBM.Close
        Set MBM = Nothing
        CmdGo_Click
    Else
        MBM.Close
        Set MBM = Nothing
        MsgBox "First Record Reached! " & vbNewLine & "Not a valid Membership Number Format! ", vbCritical + vbOKOnly
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdSave_Click()
Const METHOD_NAME = "CmdSave_Click"
On Error GoTo debugs

Dim tmpTopupAmt As Double
Dim SqlTopup As String
Dim tmpPrevLmt As Double
Dim tmpNewLmt As Double
Dim MBMSql, MBMXRef, MBMCOVSql, MBMCISO, MBMCISCov As String
Dim tmpCovID As String
Dim StartTrans As Boolean

    StartTrans = False
    tmpPrevLmt = IIf(IsNumeric(txtBalLmt), txtBalLmt, 0)
    tmpTopupAmt = IIf(IsNumeric(txtTopupAmt), txtTopupAmt, 0)
    tmpNewLmt = IIf(IsNumeric(txtNewBal), txtNewBal, 0)
    If tmpTopupAmt = 0 And Trim(tmpPrevICNo) = Trim(txtICNo) Then
        MsgBox "Nothing to Adjust ! ", vbOKOnly + vbCritical
        txtICNo.SetFocus
    Else
        Screen.MousePointer = vbHourglass
        MediConnCISDB.beginTrans
        medixmasterconn.beginTrans
        medixmasterconnMaint.beginTrans
        StartTrans = True
            If optGPlmt.Value = True Then
                SqlTopup = "Insert Into MBMCISTopupHistory(MBMNumber, MBMCovID, MBMICNo, " & _
                           "PrevICNo, PrevLimit, TopupAmt, NewLimit," & _
                           "Updatedate, UpdateBy)" & _
                           " Values ('" & MembershipNo & "','" & CovID & "','" & Trim(txtICNo) & _
                           "','" & tmpPrevICNo & "'," & tmpPrevLmt & "," & tmpTopupAmt & "," & tmpNewLmt & _
                           ",'" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "')"
            ElseIf optSPlmt.Value = True Then
                SqlTopup = "Insert Into MBMCISTopupHistory(MBMNumber, MBMCovID,  " & _
                           "PrevLimit, SP_TopupAmt, NewLimit," & _
                           "Updatedate, UpdateBy)" & _
                           " Values ('" & MembershipNo & "'," & CovID & _
                           "," & tmpPrevLmt & "," & tmpTopupAmt & "," & tmpNewLmt & _
                           ",'" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "')"
            End If
            MediConnCISDB.Execute SqlTopup
            If optGPlmt.Value = True Then
                If Trim(tmpPrevICNo) <> Trim(txtICNo) Then
                    If CovID = "00" Then
                        MBMSql = "Update MBM SET MBMIcBcPp = '" & Trim(txtICNo) & "' WHERE MBMNumber = '" & Trim(MembershipNo) & "'"
                        MBMXRef = "Update MBMCrossReference SET MBMIcBcPp = '" & Trim(txtICNo) & "' WHERE MBMNumber = '" & Trim(MembershipNo) & "'"
                        medixmasterconn.Execute MBMSql
                        medixmasterconnMaint.Execute MBMXRef
                    Else
                        MBMCOVSql = "Update MBMCoveredPersons SET MBMCICBCPPNo = '" & Trim(txtICNo) & "' WHERE MBMCNumber = '" & Trim(MembershipNo) & _
                                  "' AND MBMCCoverID='" & Trim(CovID) & "'"
                        medixmasterconn.Execute MBMCOVSql
                    End If
                End If
                If tmpNewLmt <> 0 Then
                    tmpCovID = CovID
                    If tmpSptLmtInd = "A" Or tmpCovID = "00" Then
                        MBMCISO = "Update MBMCLNOthers Set MBMCLNBalLmt = " & tmpNewLmt & " WHERE MBMNumber = '" & Trim(MembershipNo) & "'"
                        medixmasterconn.Execute MBMCISO
                    Else
                        If tmpSptLmtInd <> "N" Then
                            If tmpSptLmtInd = "B" Then tmpCovID = "01"
                            MBMCISCov = "Update MBMCoveredPersonsCLN Set MBMCCLNBalLmt = " & tmpNewLmt & " WHERE MBMCNumber = '" & Trim(MembershipNo) & _
                                        "' AND MBMCCoverID='" & Trim(tmpCovID) & "'"
                            medixmasterconn.Execute MBMCISCov
                        End If
                    End If
                End If
            End If
        MediConnCISDB.commitTrans
        medixmasterconn.commitTrans
        medixmasterconnMaint.commitTrans
        Screen.MousePointer = vbDefault
        MsgBox "Updated Successfully", vbOKOnly + vbInformation, DialogBoxTitle
        Exit Sub
ErrorSave:
        If StartTrans = True Then
            MediConnCISDB.RollbackTrans
            medixmasterconn.RollbackTrans
            medixmasterconnMaint.RollbackTrans
            Screen.MousePointer = vbDefault
        End If
        Screen.MousePointer = vbDefault
        If Err.Description <> "" Then
            Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
        End If
        MsgBox Err.Description
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub cmdSearch_Click()
    Call ClearEntries
    frmSearchMBM.Source = 7
    frmSearchMBM.Show
End Sub

Private Sub ClearEntries()
Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is ComboBox Then
        ElseIf TypeOf ctl Is ListView Then
            ctl.ListItems.Clear
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.mask = ""
            ctl.Text = ""
            ctl.mask = "##/##/####"
        ElseIf TypeOf ctl Is Label Then
        End If
    Next ctl
End Sub

Private Sub FindMBM()
Const METHOD_NAME = "FindMBM"
On Error GoTo debugs

Dim MBMRec As New ADODB.Recordset
Dim sql As String
Dim SqlOthers As String
Dim MBMORec As New ADODB.Recordset

    SqlOthers = "Select MBMGroupCompany, MBMDepartment From MBMOthers where MBMNumber='" & MembershipNo & "'"
    MBMORec.Open SqlOthers, medixmasterconn, adOpenForwardOnly, adLockOptimistic
    Do While Not MBMORec.EOF
        txtGrpComp = IIf(Len(MBMORec!MBMGroupCompany), MBMORec!MBMGroupCompany, "")
        txtDept = IIf(Len(MBMORec!MBMDepartment), MBMORec!MBMDepartment, "")
        MBMORec.MoveNext
    Loop
    MBMORec.Close
    Set MBMORec = Nothing
    
    sql = "Exec SearchCISMemberProc '" & MembershipNo & "'"
    MBMRec.Open sql, medixmasterconn, adOpenForwardOnly, adLockOptimistic
    
    Do While Not MBMRec.EOF
        CmdSave.Enabled = True
        txtInsType = MBMRec!INSType
        txtName = IIf(Len(MBMRec!MBMName), MBMRec!MBMName, "")
        txtICNo = IIf(Len(MBMRec!MBMIcBcPp), MBMRec!MBMIcBcPp, "")
        txtPlan = IIf(Len(MBMRec!PLNType), MBMRec!PLNType, "")
        txtEffDate = MBMRec!MBMPolEffDate
        txtExpDate = MBMRec!MBMPolExpDate
        'txtBalLmt = IIf(IsNumeric(MBMRec!MBMCLNBalLmt), MBMRec!MBMCLNBalLmt, 0)
        txtStatus = IIf(Len(MBMRec!MBMCLNStatus), MBMRec!MBMCLNStatus, "")
        If CovID = "00" Then Exit Do
        If CovID = MBMRec!MBMCCoverID Then
            If Len(Trim(MBMRec!MBMCName)) Then txtName = MBMRec!MBMCName
            If Len(Trim(MBMRec!MBMCICBCPPNo)) Then txtICNo = MBMRec!MBMCICBCPPNo
            If Len(Trim(MBMRec!MBMCCLNPLNCode)) Then txtPlan = MBMRec!MBMCCLNPLNCode
            If IsNumeric(MBMRec!MBMCCLNBalLmt) Then txtBalLmt = MBMRec!MBMCCLNBalLmt
            If Len(Trim(MBMRec!MBMCCLNStatus)) Then txtStatus = MBMRec!MBMCCLNStatus
            Exit Do
        End If
        MBMRec.MoveNext
    Loop
    MBMRec.Close
    Set MBMRec = Nothing
    tmpPrevICNo = Trim(txtICNo)
    If txtStatus = "A" Then txtStatus = "A - Active"
    If txtStatus = "C" Then txtStatus = "C - Cancelled"
    If txtStatus = "D" Then txtStatus = "Delete - Delete"
    If Trim(txtName) <> "" Then
        Call GetSptLmtInd
    End If
    'New Condition
    sql = ""
    sql = "Exec spUtilisationOP_test1 NULL,'" & MembershipNo & "',NULL,NULL,NULL,NULL"
    MBMRec.Open sql, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
    Do While Not MBMRec.EOF
        If optGPlmt.Value = True And CovID = MBMRec!CovID Then
            txtBalLmt = IIf(IsNumeric(MBMRec!GPBalance), MBMRec!GPBalance, 0)
        ElseIf optSPlmt.Value = True And CovID = MBMRec!CovID Then
            txtBalLmt = IIf(IsNumeric(MBMRec!SPBalance), MBMRec!SPBalance, 0)
        End If
        MBMRec.MoveNext
    Loop
    MBMRec.Close
    Set MBMRec = Nothing
    
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub Form_Load()
Call ActionLogUpdate(12, "FrmMBMTopupLmt")
optGPlmt.Value = True
optSPlmt.Value = False
End Sub

Private Sub optGPlmt_Click()
Dim sql As String
Dim MBMRec As New ADODB.Recordset
sql = ""
    sql = "Exec spUtilisationOP_test1 NULL,'" & MembershipNo & "',NULL,NULL,NULL,NULL"
    MBMRec.Open sql, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
    Do While Not MBMRec.EOF
        If optGPlmt.Value = True And CovID = MBMRec!CovID Then
            txtBalLmt = IIf(IsNumeric(MBMRec!GPBalance), MBMRec!GPBalance, 0)
        ElseIf optSPlmt.Value = True And CovID = MBMRec!CovID Then
            txtBalLmt = IIf(IsNumeric(MBMRec!SPBalance), MBMRec!SPBalance, 0)
        End If
        MBMRec.MoveNext
    Loop
    MBMRec.Close
    Set MBMRec = Nothing
End Sub

Private Sub optSPlmt_Click()
Dim sql As String
Dim MBMRec As New ADODB.Recordset
    sql = ""
    sql = "Exec spUtilisationOP_test1 NULL,'" & MembershipNo & "',NULL,NULL,NULL,NULL"
    MBMRec.Open sql, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
    Do While Not MBMRec.EOF
        If optGPlmt.Value = True And CovID = MBMRec!CovID Then
            txtBalLmt = IIf(IsNumeric(MBMRec!GPBalance), MBMRec!GPBalance, 0)
        ElseIf optSPlmt.Value = True And CovID = MBMRec!CovID Then
            txtBalLmt = IIf(IsNumeric(MBMRec!SPBalance), MBMRec!SPBalance, 0)
        End If
        MBMRec.MoveNext
    Loop
    MBMRec.Close
    Set MBMRec = Nothing
End Sub

Private Sub txtICNo_LostFocus()
    txtICNo = ChkStr(txtICNo)
End Sub

Private Sub txtMBMNumber_GotFocus()
    CmdSave.Enabled = False
    CmdGo.Default = True
End Sub

Private Sub txtMBMNumber_LostFocus()
    txtMBMNumber = ChkStr(txtMBMNumber)
    CmdGo.Default = False
End Sub

Private Sub txtTopupAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtTopupAmt_LostFocus()
Dim tmpTopupAmt As Double
Dim tmpBalLmt As Double
    tmpTopupAmt = IIf(IsNumeric(txtTopupAmt), txtTopupAmt, 0)
    tmpBalLmt = IIf(IsNumeric(txtBalLmt), txtBalLmt, 0)
    txtNewBal = tmpBalLmt + tmpTopupAmt
End Sub

Private Sub GetSptLmtInd()
Const METHOD_NAME = "GetSptLmtInd"
On Error GoTo debugs

Dim AnnLmt As New ADODB.Recordset
Dim StrsqlAnnLmt, strsql As String
Dim AnnLmtType As New ADODB.Recordset
    tmpSptLmtInd = "N"
    StrsqlAnnLmt = "Select SuppLimitStatus from AnnualLimit where PLNCode = '" & Trim(txtPlan) & _
               "' AND INSCode='" & Trim(txtInsType) & _
               "' AND AnnEffDate <= '" & Format(txtEffDate, "mm/dd/yyyy") & "' order by AnnEffDate  desc"
    AnnLmt.Open StrsqlAnnLmt, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    If AnnLmt.EOF = False Then
        tmpSptLmtInd = IIf(Len(Trim(AnnLmt!SuppLimitStatus)), Trim(AnnLmt!SuppLimitStatus), "A")
    End If
    AnnLmt.Close
    Set AnnLmt = Nothing
    If Trim(txtInsType) = "F" Or Trim(txtInsType) = "H" Then
        lblLimitType = tmpSptLmtInd
        strsql = "Select LimitDesc from AnnualLimitType where LimitCode='" & Trim(tmpSptLmtInd) & " '"
        AnnLmtType.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        If AnnLmtType.EOF = False Then
            lblLimitType = tmpSptLmtInd & " - " & AnnLmtType!LimitDesc
        End If
    End If
    txtTopupAmt.Enabled = True
    If tmpSptLmtInd = "N" Then
        lblLimitType = "N - Unlimited Limit"
        txtTopupAmt.Enabled = False
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If


End Sub
