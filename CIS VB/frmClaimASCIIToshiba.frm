VERSION 5.00
Begin VB.Form frmClaimASCIIToshiba 
   Caption         =   "Generate Claim ASCII (Toshiba)"
   ClientHeight    =   2472
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   7188
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2472
   ScaleWidth      =   7188
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Height          =   1575
      Left            =   80
      TabIndex        =   4
      Top             =   240
      Width           =   7095
      Begin VB.TextBox txtBordxNoTo 
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
         Left            =   3240
         TabIndex        =   1
         Top             =   1170
         Width           =   1455
      End
      Begin VB.ComboBox CboExportStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox txtBordxNoFr 
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
         Left            =   1440
         TabIndex        =   0
         Top             =   1170
         Width           =   1455
      End
      Begin VB.ComboBox cboCLNList 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5280
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "To"
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
         Left            =   3000
         TabIndex        =   22
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label txtFilName 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1440
         TabIndex        =   15
         Top             =   840
         Width           =   3225
      End
      Begin VB.Label Label1 
         Caption         =   "Export Status"
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
         TabIndex        =   14
         Top             =   1920
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label txtLocalFileLocation 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1440
         TabIndex        =   13
         Top             =   540
         Width           =   3225
      End
      Begin VB.Label Label26 
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
         Index           =   0
         Left            =   5160
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label24 
         Caption         =   "File on Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label txtPath 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   3225
      End
      Begin VB.Label Label28 
         Caption         =   "File on PC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   615
         Width           =   765
      End
      Begin VB.Label Label25 
         Caption         =   "Export File Name"
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
         TabIndex        =   8
         Top             =   900
         Width           =   1245
      End
      Begin VB.Label Label26 
         Caption         =   "Bordx No"
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
         TabIndex        =   7
         Top             =   1180
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   80
      TabIndex        =   16
      Top             =   1800
      Width           =   7095
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6240
         TabIndex        =   3
         Top             =   240
         Width           =   705
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         TabIndex        =   2
         Top             =   240
         Width           =   720
      End
      Begin VB.Label txtTotalVisits 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1150
         TabIndex        =   17
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label17 
         Caption         =   "Total Visit(s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   270
         Width           =   960
      End
      Begin VB.Label lblTotalAmt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3480
         TabIndex        =   19
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label Label17 
         Caption         =   "Total Bordx Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   1
         Left            =   2280
         TabIndex        =   18
         Top             =   270
         Width           =   1200
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Generate Claim ASCII (Toshiba)"
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
      TabIndex        =   21
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmClaimASCIIToshiba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpCLNClientCode As String

Private Sub cboCLNList_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNList.Text, cboCLNList.SelStart) + Chr(KeyAscii) + Mid(cboCLNList.Text, cboCLNList.SelStart + cboCLNList.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCLNList.ListCount - 1
 
        If NewText = LCase(Left(cboCLNList.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNList.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCLNList.Text = ValidValue
                cboCLNList.SelStart = 0
                cboCLNList.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
Const METHOD_NAME = "cmdGenerate_Click"
On Error GoTo debugs
Dim RecordSaved
    txtTotalVisits = 0
    lblTotalAmt = 0
    
    If txtBordxNoFr = "" Then
        MsgBox "Please Enter Bordx No Fr ", vbOKOnly + vbCritical, DialogBoxTitle
        txtBordxNoFr.SetFocus
    ElseIf txtBordxNoTo = "__/__/____" Then
        MsgBox "Please Enter Bordx No To ", vbOKOnly + vbCritical, DialogBoxTitle
        txtBordxNoTo.SetFocus
    'ElseIf cboCLNList = "" Then
    '    MsgBox "Please Select Clinic Name ", vbOKOnly + vbCritical, DialogBoxTitle
    '    CboExportStatus.SetFocus
    Else
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation, DialogBoxTitle)
        If RecordSaved = vbYes Then
            Screen.MousePointer = vbHourglass
                Call ClaimASCII
            Screen.MousePointer = vbDefault
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ClaimASCII()
Const METHOD_NAME = "ClaimASCII"
On Error GoTo debugs
Dim strsql As String
Dim CLMBordx As New ADODB.Recordset
Dim PAYCode As String
Dim sql As String
Dim ClaimNo As String
Dim VersionNo As String
Dim strAppend As String
Dim DataString As String
Dim tmpMCDays As Integer
Dim tmpCLNCode As String
Dim tmpClaimActual As Double
Dim tmpTotalAmount As Double
Dim tmpBordxDay As String
Dim tmpBordxMonth As String
Dim tmpBordxYear As String
Dim tmpRelationship As String
Dim tmpVisitDay As String
Dim tmpVisitMonth As String
Dim tmpVisitYear As String
Dim tmpbordxamt As Double
Dim tmpPostingTime As String
Dim tmpPostinghh As String
Dim tmpPostingmm As String
Dim tmpPostingss As String
Dim tmpPostingID As String
Dim tmpToshibaICD As String
        tmpPostingID = "001000"
        'StrAppend = ""
        'sql = ""
        'StrAppend = " AND BordxRefNo >= '" & Trim(txtBordxNoFr) & "' AND BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
        
        'StrAppend = StrAppend & " AND ClinicCode = '" & Trim(Left(cboCLNList, IIf(InStr(1, cboCLNList, "-") - 1 < 0, Len(cboCLNList), InStr(1, cboCLNList, "-") - 1))) & "'"
        
        'sql = "SELECT * From View_ClaimDetails_ASCII where 0 = 0 "
        
        'StrSql = sql + StrAppend
        'CLMBordx.Open StrSql, MediConnCISWS, adOpenStatic, adLockOptimistic


    Dim cmd As New ADODB.Command
    Dim Prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim total As Integer
    
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
            
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimDetails_ASCII"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("BordxRefNoFrom", adVarChar, adParamInput, 20, Trim(txtBordxNoFr))
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("BordxRefNoTo", adVarChar, adParamInput, 20, Trim(txtBordxNoTo))
    cmd.Parameters.Append prm2
    CLMBordx.CursorLocation = adUseClient
    CLMBordx.CursorType = adOpenForwardOnly
    CLMBordx.CacheSize = 100
    Set CLMBordx = cmd.Execute
    
    
        If CLMBordx.EOF Then
            MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            tmpMCDays = 0
            tmpClaimActual = 0
            tmpTotalAmount = 0
        

            txtFilName = txtBordxNoFr & "(M-MEDIX)" & ".txt"
            txtFilName = txtBordxNoFr & "(M-MEDIX)" & ".txt"
            
            Open txtPath & "\" & txtFilName For Output As #1
            Open txtLocalFileLocation & txtFilName For Output As #2
            
        Do Until CLMBordx.EOF
            tmpBordxDay = Format(Day(CLMBordx!BordxDate), "00")
            tmpBordxMonth = Format(Month(CLMBordx!BordxDate), "00")
            tmpBordxYear = Format(Year(CLMBordx!BordxDate), "0000")
            
            tmpVisitDay = Format(Day(CLMBordx!DateVisit), "00")
            tmpVisitMonth = Format(Month(CLMBordx!DateVisit), "00")
            tmpVisitYear = Format(Year(CLMBordx!DateVisit), "0000")
            tmpMCDays = IIf(IsNumeric(CLMBordx!MCDays), Trim(CLMBordx!MCDays), "0")
            tmpCLNCode = Trim(Left(cboCLNList, IIf(InStr(1, cboCLNList, "-") - 1 < 0, Len(cboCLNList), InStr(1, cboCLNList, "-") - 1)))
            tmpToshibaICD = IIf(Len(CLMBordx!ToshibaICD), CLMBordx!ToshibaICD, "")
            tmpPostingID = Format(tmpPostingID + 1, "000000")
            DataString = ""
            
            DataString = l("M-MEDIX", 8)
            DataString = DataString & l(CLMBordx!Bordxrefno, 15)
            DataString = DataString & l(tmpBordxDay, 2)
            DataString = DataString & l(tmpBordxMonth, 2)
            DataString = DataString & l(tmpBordxYear, 4)
            'DataString = DataString & l(CLMBordx!MBMEmployeeNo, 7)
            DataString = DataString & Left(CLMBordx!MBMEmployeeNo & Space(7), 7)
            If Trim(CLMBordx!coverid) <> "00" Then
                'If Len(CLMBordx!PatientName) Then
                '    DataString = DataString & l(CLMBordx!PatientName, 35)
                'Else
                '    DataString = DataString & l("", 35)
                'End If
                
                If Len(CLMBordx!PatientName) Then
                    DataString = DataString & l(CLMBordx!PatientName, 60)
                Else
                    DataString = DataString & l("", 60)
                End If
            Else
                'DataString = DataString & l("", 35)
                DataString = DataString & l("", 60)
            End If
            DataString = DataString & l("NO", 2)
            DataString = DataString & l(tmpVisitDay, 2)
            DataString = DataString & l(tmpVisitMonth, 2)
            DataString = DataString & l(tmpVisitYear, 4)
            If Len(tmpToshibaICD) Then
                DataString = DataString & l(tmpToshibaICD, 2)
                DataString = DataString & l(CLMBordx!ReasonDesc, 60)
            Else
                DataString = DataString & l("52", 2)
                DataString = DataString & l("Medication & Treatment", 60)
            End If
            DataString = DataString & l(CLMBordx!MCDays, 3)
            DataString = DataString & l(Replace(Format(CLMBordx!CLMBordxAmt, "#,##,0.00"), ".", ""), 9)
            DataString = DataString & l("P", 1)
            DataString = DataString & l(tmpBordxDay, 2)
            DataString = DataString & l(tmpBordxMonth, 2)
            DataString = DataString & l(tmpBordxYear, 4)
            DataString = DataString & l(tmpBordxDay, 2)
            DataString = DataString & l(tmpBordxMonth, 2)
            DataString = DataString & l(tmpBordxYear, 4)
            DataString = DataString & l(tmpPostingID, 6)
            tmpPostinghh = Format(Time, "hh")
            tmpPostingmm = Format(Time, "mm")
            tmpPostingss = Format(Time, "ss")
            tmpPostingTime = tmpPostinghh & tmpPostingmm & tmpPostingss
            DataString = DataString & l(tmpPostingTime, 6)
            DataString = DataString & l("", 10)
            DataString = DataString & l("000000", 6)
            
            txtTotalVisits = txtTotalVisits + 1
            If IsNumeric(CLMBordx!CLMBordxAmt) Then
                lblTotalAmt = CDbl(lblTotalAmt) + IIf(IsNumeric(CLMBordx!CLMBordxAmt), CDbl(CLMBordx!CLMBordxAmt), 0)
            End If
            CLMBordx.MoveNext
            
            Print #1, DataString
            Print #2, DataString
        Loop
            
                cmdGenerate.Enabled = False
                Close #1
                Close #2
                
        MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
        cmdGenerate.Enabled = True
    End If
    CLMBordx.Close
    Set CLMBordx = Nothing
    
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_Load()
Const METHOD_NAME = "Form_Load"
On Error GoTo debugs
Dim CLN As New ADODB.Recordset
Dim strsql As String
Call ActionLogUpdate(12, "frmClaimASCIIToshiba")
    
    txtPath = SRVPath & "ClinicalClaimASCII"
    txtLocalFileLocation = "C:\ClinicalClaimsASCII\"
    
    
    cboCLNList.Clear
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
    strsql = "EXEC SearchCLN"
    CLN.Open strsql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic

    Do Until CLN.EOF
        cboCLNList.AddItem CLN!CLNCode & " - " & CLN!CLNName
        CLN.MoveNext
    Loop
    CLN.Close
    Set CLN = Nothing
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
