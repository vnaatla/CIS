VERSION 5.00
Begin VB.Form frmClaimASCIISinger 
   Caption         =   "Generate Claim ASCII (Singer)"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   7212
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   7212
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Height          =   1575
      Left            =   80
      TabIndex        =   4
      Top             =   240
      Width           =   7095
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
         TabIndex        =   5
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   540
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
         TabIndex        =   6
         Top             =   1185
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   80
      TabIndex        =   15
      Top             =   1800
      Width           =   7095
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
         TabIndex        =   19
         Top             =   270
         Width           =   1080
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
         TabIndex        =   18
         Top             =   240
         Width           =   1920
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
         TabIndex        =   17
         Top             =   270
         Width           =   1200
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Generate Claim ASCII (Singer)"
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
      TabIndex        =   20
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmClaimASCIISinger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpCLNClientCode As String

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
Dim RecordSaved
    txtTotalVisits = 0
    lblTotalAmt = 0
    
    If txtBordxNoFr = "" Then
        MsgBox "Please Enter Bordx No Fr ", vbOKOnly + vbCritical, DialogBoxTitle
        txtBordxNoFr.SetFocus
    ElseIf txtBordxNoTo = "__/__/____" Then
        MsgBox "Please Enter Bordx No To ", vbOKOnly + vbCritical, DialogBoxTitle
        txtBordxNoTo.SetFocus
    Else
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation, DialogBoxTitle)
        If RecordSaved = vbYes Then
            Screen.MousePointer = vbHourglass
                Call ClaimASCII
            Screen.MousePointer = vbDefault
        End If
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
Dim tmpClaimActual As Double
Dim tmpTotalAmount As Double
Dim tmpBordxDay As String
Dim tmpBordxMonth As String
Dim tmpBordxYear As String
Dim tmpRelationship As String
Dim tmpbordxdate As String
Dim z As Boolean
Dim Header As String
Dim Trailer As String
Dim tmpstring As String
        tmpstring = ""
        z = False
        
tmpBordxDay = ""
tmpBordxMonth = ""
tmpBordxYear = ""
        'StrAppend = " AND BordxRefNo >= '" & Trim(txtBordxNoFr) & "'" & _
        '            "AND BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
        
        
        'sql = "SELECT * From View_ClaimDetails_ASCII where 0 = 0 "
        
        'strsql = sql + StrAppend
        'CLMBordx.Open strsql, MediConnCISWS, adOpenStatic, adLockOptimistic

    Dim cmd As New ADODB.Command
    Dim Prm1 As New ADODB.Parameter
    Dim Prm2 As New ADODB.Parameter
    Dim total As Integer

    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimDetails_ASCII"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("BordxRefNoFrom", adVarChar, adParamInput, 20, Trim(txtBordxNoFr))
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("BordxRefNoTo", adVarChar, adParamInput, 20, Trim(txtBordxNoTo))
    cmd.Parameters.Append Prm2
    CLMBordx.CursorLocation = adUseClient
    CLMBordx.CursorType = adOpenForwardOnly
    CLMBordx.CacheSize = 100
    Set CLMBordx = cmd.Execute
    
        If CLMBordx.EOF Then
            MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
           
            tmpBordxDay = Day(CLMBordx!BordxDate)
            If Len(Trim(tmpBordxDay)) = 1 Then
                tmpBordxDay = "0" + tmpBordxDay
            End If
            tmpBordxMonth = Month(CLMBordx!BordxDate)
            If Len(Trim(tmpBordxMonth)) = 1 Then
                tmpBordxMonth = "0" + tmpBordxMonth
            End If
            'tmpBordxYear = Format(Year(CLMBordx!BordxDate), "yy")
            tmpBordxYear = Mid$(CLMBordx!BordxDate, 9, 2)
            
            txtFilName = "CC" & tmpBordxYear & tmpBordxMonth & tmpBordxDay & "01"
            z = CreatePath(txtLocalFileLocation)
            
            Open txtPath & "\" & txtFilName For Output As #1
            Open txtLocalFileLocation & txtFilName For Output As #2
            
            Header = txtFilName
            
            Print #1, Trim(Header)
            Print #2, Trim(Header)

        Do Until CLMBordx.EOF
            DataString = ""
            tmpClaimActual = 0
            tmpTotalAmount = 0
            tmpbordxdate = CLMBordx!BordxDate
            DataString = "1" & ","
            DataString = DataString & IIf(Len(CLMBordx!BordxDate), Trim(CLMBordx!BordxDate), "") & ","
            DataString = DataString & "01" & ","
            DataString = DataString & """"
            DataString = DataString & IIf(Len(CLMBordx!MBMIcBcPp), Trim(CLMBordx!MBMIcBcPp), "")
            DataString = DataString & """" & ","
            DataString = DataString & """"
            DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "")
            DataString = DataString & """" & ","
            DataString = DataString & """"
            DataString = DataString & IIf(Len(CLMBordx!PatientName), Trim(CLMBordx!PatientName), "")
            DataString = DataString & """" & ","
            DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & ","
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMTotalActual), Trim(CLMBordx!CLMTotalActual), "0") & ","
            DataString = DataString & """"
            DataString = DataString & IIf(Len(CLMBordx!ClinicCode), Trim(CLMBordx!ClinicCode), "")
            DataString = DataString & """" & ","
            DataString = DataString & """"
            DataString = DataString & IIf(Len(CLMBordx!CLNName), Trim(CLMBordx!CLNName), "")
            
            DataString = DataString & """"
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
            
                Trailer = "2" & "," & tmpbordxdate & "," & "01" & ",,,,," & lblTotalAmt
                'Debug.Print Trailer
                Print #1, Trailer
                Print #2, Trailer

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
Call ActionLogUpdate(12, "frmClaimASCIISinger")
Dim CLN As New ADODB.Recordset
Dim strsql As String
    
    txtPath = SRVPath & "ClinicalClaimASCII"
    txtLocalFileLocation = "C:\ClinicalClaimsASCII\"
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Function CreatePath(NewPath) As Boolean
Const METHOD_NAME = "CreatePath"
On Error GoTo debugs
    'Add a trailing slash if none
    If Right(NewPath, 1) <> "\" Then
        NewPath = NewPath & "\"
    End If
    
    'Call API
    If MakeSureDirectoryPathExists(NewPath) <> 0 Then
        'No errors, return True
        CreatePath = True
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

