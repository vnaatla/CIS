VERSION 5.00
Begin VB.Form frmClaimDetailsASCII 
   Caption         =   "Generate Claim ASCII (Bintulu Port)"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2700
   ScaleWidth      =   7200
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Height          =   1815
      Left            =   80
      TabIndex        =   5
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
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1430
         Width           =   3255
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
      Begin VB.Label Label26 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
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
         Top             =   1185
         Width           =   315
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
         TabIndex        =   21
         Top             =   1180
         Width           =   915
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
         TabIndex        =   14
         Top             =   900
         Width           =   1245
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
         TabIndex        =   13
         Top             =   615
         Width           =   765
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
         TabIndex        =   12
         Top             =   240
         Width           =   3225
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
         Left            =   120
         TabIndex        =   10
         Top             =   1450
         Width           =   915
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
         TabIndex        =   9
         Top             =   540
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
         TabIndex        =   8
         Top             =   1920
         Visible         =   0   'False
         Width           =   1155
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
         TabIndex        =   7
         Top             =   840
         Width           =   3225
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   80
      TabIndex        =   15
      Top             =   2040
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   960
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
         Left            =   2160
         TabIndex        =   20
         Top             =   270
         Width           =   1200
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
         Left            =   3360
         TabIndex        =   19
         Top             =   240
         Width           =   1920
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
         TabIndex        =   17
         Top             =   270
         Width           =   1080
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Generate Claim ASCII (Bintulu Port)"
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
      TabIndex        =   18
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmClaimDetailsASCII"
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
    ElseIf cboCLNList = "" Then
        MsgBox "Please Select Clinic Name ", vbOKOnly + vbCritical, DialogBoxTitle
        cboCLNList.SetFocus
    Else
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation, DialogBoxTitle)
        If RecordSaved = vbYes Then
            Screen.MousePointer = vbHourglass
                'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
                    Call ClaimASCII
                'MediConnCISWS.Close
                'Set MediConnCISWS = Nothing
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
Dim StrAppend As String
Dim DataString As String
Dim tmpMCDays As Integer
Dim tmpCLNCode As String
Dim tmpClaimActual As Double
Dim tmpTotalAmount As Double
Dim tmpBordxMonth As String
Dim tmpBordxYear As String
Dim tmpRelationship As String

        StrAppend = " AND BordxRefNo >= '" & Trim(txtBordxNoFr) & "'" & _
                    "AND BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
        
        StrAppend = StrAppend & " AND ClinicCode = '" & Trim(Left(cboCLNList, IIf(InStr(1, cboCLNList, "-") - 1 < 0, Len(cboCLNList), InStr(1, cboCLNList, "-") - 1))) & "'"
        
        sql = "SELECT * From View_ClaimDetails_ASCII where 0 = 0 "
        
        strsql = sql + StrAppend
        CLMBordx.Open strsql, MediConnCISWS, adOpenStatic, adLockOptimistic

    
        If CLMBordx.EOF Then
            MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
                  
            tmpBordxMonth = Month(CLMBordx!BordxDate)
            tmpBordxYear = Year(CLMBordx!BordxDate)
            
            tmpCLNCode = Trim(Left(cboCLNList, IIf(InStr(1, cboCLNList, "-") - 1 < 0, Len(cboCLNList), InStr(1, cboCLNList, "-") - 1)))
            txtFilName = txtBordxNoTo & "(" & tmpCLNCode & ")" & tmpBordxMonth & tmpBordxYear
            Open txtPath & "\" & txtFilName For Output As #1
            Open txtLocalFileLocation & txtFilName For Output As #2
            
        Do Until CLMBordx.EOF
            DataString = ""
            tmpMCDays = 0
            tmpClaimActual = 0
            tmpTotalAmount = 0
            
            If Trim(CLMBordx!Relation) = "P" Then
                tmpRelationship = "E"
            ElseIf Trim(CLMBordx!Relation) = "H" Then
                tmpRelationship = "U"
            ElseIf Trim(CLMBordx!Relation) = "W" Then
                tmpRelationship = "U"
            Else
                tmpRelationship = CLMBordx!Relation
            End If
            
            DataString = "BPSB" & Chr(124)
            DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "") & Chr(124)
            DataString = DataString & IIf(Len(CLMBordx!PatientName), Trim(CLMBordx!PatientName), "") & Chr(124)
            DataString = DataString & IIf(Len(tmpRelationship), Trim(tmpRelationship), "") & Chr(124)
            DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & Chr(124)
            DataString = DataString & IIf(Len(CLMBordx!VisitTime), Trim(CLMBordx!VisitTime), "") & Chr(124)
            DataString = DataString & "" & Chr(124)
            DataString = DataString & "M" & Chr(124)
            
            DataString = DataString & "0" & Chr(124)
            'DataString = DataString & IIf(IsNumeric(CLMBordx!ConsultFee), Trim(CLMBordx!ConsultFee), "0") & Chr(124)
            'tmpTotalAmount = CDbl(tmpTotalAmount) + IIf(IsNumeric(CLMBordx!ConsultFee), CDbl(CLMBordx!ConsultFee), "0")
            
            DataString = DataString & "0" & Chr(124)
            'DataString = DataString & IIf(IsNumeric(CLMBordx!MedicalFee), Trim(CLMBordx!MedicalFee), "0") & Chr(124)
            'tmpTotalAmount = CDbl(tmpTotalAmount) + IIf(IsNumeric(CLMBordx!MedicalFee), CDbl(CLMBordx!MedicalFee), "0")
            
            DataString = DataString & "0" & Chr(124)
            'DataString = DataString & IIf(IsNumeric(CLMBordx!InjectionFee), Trim(CLMBordx!InjectionFee), "0") & Chr(124)
            'tmpTotalAmount = CDbl(tmpTotalAmount) + IIf(IsNumeric(CLMBordx!InjectionFee), CDbl(CLMBordx!InjectionFee), "0")
            
            DataString = DataString & "0" & Chr(124)
            'DataString = DataString & IIf(IsNumeric(CLMBordx!CLMXRay), Trim(CLMBordx!CLMXRay), "0") & Chr(124)
            'tmpTotalAmount = CDbl(tmpTotalAmount) + IIf(IsNumeric(CLMBordx!CLMXRay), CDbl(CLMBordx!CLMXRay), "0")
            
            DataString = DataString & "0" & Chr(124)
            'DataString = DataString & IIf(IsNumeric(CLMBordx!CLMECG), Trim(CLMBordx!CLMECG), "0") & Chr(124)
            'tmpTotalAmount = CDbl(tmpTotalAmount) + IIf(IsNumeric(CLMBordx!CLMECG), CDbl(CLMBordx!CLMECG), "0")
            
            DataString = DataString & "0" & Chr(124)
            'DataString = DataString & IIf(IsNumeric(CLMBordx!CLMUrine), Trim(CLMBordx!CLMUrine), "0") & Chr(124)
            'tmpTotalAmount = CDbl(tmpTotalAmount) + IIf(IsNumeric(CLMBordx!CLMUrine), CDbl(CLMBordx!CLMUrine), "0")
            
            DataString = DataString & "0" & Chr(124)
            'DataString = DataString & IIf(IsNumeric(CLMBordx!Lab), Trim(CLMBordx!Lab), "0") & Chr(124)
            'tmpTotalAmount = CDbl(tmpTotalAmount) + IIf(IsNumeric(CLMBordx!Lab), CDbl(CLMBordx!Lab), "0")
            
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMTotalActual), Trim(CLMBordx!CLMTotalActual), "0") & Chr(124)
            
            DataString = DataString & IIf(Len(CLMBordx!CASTreatReason1), Trim(CLMBordx!CASTreatReason1), "199") & Chr(124)
            DataString = DataString & IIf(IsNumeric(CLMBordx!MCDays), Trim(CLMBordx!MCDays), "0") & Chr(124)
            tmpMCDays = IIf(IsNumeric(CLMBordx!MCDays), Trim(CLMBordx!MCDays), "0")
            If tmpMCDays > 0 Then
                DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & Chr(124)
            Else
                DataString = DataString & "" & Chr(124)
            End If
            DataString = DataString & "" & Chr(124)
            DataString = DataString & "" & Chr(13)
            
            'If CDbl(tmpTotalAmount) <> CDbl(CLMBordx!CLMTotalActual) Then
            '    MsgBox "Error in Bill Amount! Case Number - " & CLMBordx!CasNumber, vbOKOnly
            'End If
            txtTotalVisits = txtTotalVisits + 1
            lblTotalAmt = CDbl(lblTotalAmt) + IIf(IsNumeric(CLMBordx!CLMTotalActual), CDbl(CLMBordx!CLMTotalActual), 0)

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

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_Load()
Dim CLN As New ADODB.Recordset
Dim strsql As String
Call ActionLogUpdate(12, "frmClaimDetailsASCII")
    txtPath = SRVPath & "ClinicalClaimASCII"
    txtLocalFileLocation = "C:\ClinicalClaimsASCII\"
    
    
    cboCLNList.Clear
    strsql = "EXEC SearchGrpCLN"
    CLN.Open strsql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic

    Do Until CLN.EOF
        cboCLNList.AddItem CLN!GrpComClnCode & " - " & CLN!CLNName
        CLN.MoveNext
    Loop
    CLN.Close
    Set CLN = Nothing
    
End Sub

