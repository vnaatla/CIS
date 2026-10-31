VERSION 5.00
Begin VB.Form frmClaimASCIIUtusan 
   Caption         =   "Generate Claim ASCII (Utusan)"
   ClientHeight    =   2484
   ClientLeft      =   60
   ClientTop       =   516
   ClientWidth     =   7212
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2484
   ScaleWidth      =   7212
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   80
      TabIndex        =   13
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   240
         Width           =   720
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
         TabIndex        =   19
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
         TabIndex        =   18
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
   End
   Begin VB.Frame Frame8 
      Height          =   1575
      Left            =   80
      TabIndex        =   0
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   1
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   315
         Width           =   1125
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   840
         Width           =   3225
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Generate Claim ASCII (Utusan)"
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
Attribute VB_Name = "frmClaimASCIIUtusan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpCLNClientCode As String
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

    Dim cmd As New ADODB.Command
    Dim Prm1 As New ADODB.Parameter
    Dim Prm2 As New ADODB.Parameter
    
    Dim cmd2 As New ADODB.Command
    Dim Prm3 As New ADODB.Parameter
    Dim Prm4 As New ADODB.Parameter
    
    Dim cmd3 As New ADODB.Command
    Dim Prm5 As New ADODB.Parameter
    Dim Prm6 As New ADODB.Parameter
    
    Dim cmd4 As New ADODB.Command
    Dim Prm7 As New ADODB.Parameter
    Dim Prm8 As New ADODB.Parameter
    
    Dim cmd5 As New ADODB.Command
    Dim Prm9 As New ADODB.Parameter
    Dim Prm10 As New ADODB.Parameter
    
    Dim cmd6 As New ADODB.Command
    Dim Prm11 As New ADODB.Parameter
    Dim Prm12 As New ADODB.Parameter
    
    Dim cmd7 As New ADODB.Command
    Dim Prm13 As New ADODB.Parameter
    Dim Prm14 As New ADODB.Parameter
    
    Dim cmd8 As New ADODB.Command
    Dim Prm15 As New ADODB.Parameter
    Dim Prm16 As New ADODB.Parameter
    
    Dim cmd9 As New ADODB.Command
    Dim Prm17 As New ADODB.Parameter
    Dim Prm18 As New ADODB.Parameter
    
    Dim cmd10 As New ADODB.Command
    Dim Prm19 As New ADODB.Parameter
    Dim Prm20 As New ADODB.Parameter
    
    Dim cmd11 As New ADODB.Command
    Dim Prm21 As New ADODB.Parameter
    Dim Prm22 As New ADODB.Parameter
    
    Dim cmd12 As New ADODB.Command
    Dim Prm23 As New ADODB.Parameter
    Dim Prm24 As New ADODB.Parameter
    
    Dim cmd13 As New ADODB.Command
    Dim Prm25 As New ADODB.Parameter
    Dim Prm26 As New ADODB.Parameter
    
    Dim total As Integer


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
                'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
                    Call ClaimASCII
                    Call ClaimASCIITwo
                'MediConnCISWS.Close
                'Set MediConnCISWS = Nothing
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub ClaimASCIITwo()
Const METHOD_NAME = "ClaimASCIITwo"
On Error GoTo debugs
        'tmpstring = ""
        'z = False
        
        'tmpBordxDay = ""
        'tmpBordxMonth = ""
        'tmpBordxYear = ""
        'StrAppend = " AND BordxRefNo >= '" & Trim(txtBordxNoFr) & "'" & _
        '            "AND BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
        
        

            Set cmd7.ActiveConnection = MediConnCISWS
            cmd7.CommandText = "ClaimDetails_ASCII_Utusan"
            cmd7.CommandType = adCmdStoredProc
            cmd7.CommandTimeout = 300
            Set Prm13 = cmd7.CreateParameter("@strAppend", adVarChar, adParamInput, 200, Trim(StrAppend))
            cmd7.Parameters.Append Prm13
            Set Prm14 = cmd7.CreateParameter("@strAppendCase", adVarChar, adParamInput, 1, 7)
            cmd7.Parameters.Append Prm14
            CLMBordx.CursorLocation = adUseClient
            CLMBordx.CursorType = adOpenForwardOnly
            CLMBordx.CacheSize = 100
            Set CLMBordx = cmd7.Execute
            
                'If CLMBordx.EOF Then
        
            Do Until CLMBordx.EOF
                DataString = ""
                DataString = "UMMB" & ","
                DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "") & ","
                DataString = DataString & "03B" & ","
                DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & "FINANCE" & ","
                DataString = DataString & IIf(Len(CLMBordx!ClinicCode), Trim(CLMBordx!ClinicCode), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceNo), Trim(CLMBordx!InvoiceNo), "") & ", "
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceDate), Trim(CLMBordx!InvoiceDate), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "Upload From MediExpress" & ","
                
                If Trim(CLMBordx!coverid) <> "00" Then
                    If Trim(CLMBordx!MBMCRELCode) = "W" Then
                        DataString = DataString & "I" & ","
                    ElseIf Trim(CLMBordx!MBMCRELCode) = "H" Then
                        DataString = DataString & "S" & ","
                    Else
                        DataString = DataString & "A" & ","
                    End If
                Else
                    DataString = DataString & "" & ","
                End If
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                If Trim(CLMBordx!coverid) <> "00" Then
                    DataString = DataString & IIf(Len(CLMBordx!MBMCName), Trim(CLMBordx!MBMCName), "") & ","
                End If
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & Chr(13)
                    
                    CLMBordx.MoveNext
                    
                    Print #1, DataString
                    Print #2, DataString
                Loop
                
                'End If
                
                CLMBordx.Close
                Set CLMBordx = Nothing
            
            Set cmd8.ActiveConnection = MediConnCISWS
            cmd8.CommandText = "ClaimDetails_ASCII_Utusan"
            cmd8.CommandType = adCmdStoredProc
            cmd8.CommandTimeout = 300
            Set Prm15 = cmd8.CreateParameter("@strAppend", adVarChar, adParamInput, 200, Trim(StrAppend))
            cmd8.Parameters.Append Prm15
            Set Prm16 = cmd8.CreateParameter("@strAppendCase", adVarChar, adParamInput, 1, 8)
            cmd8.Parameters.Append Prm16
            CLMBordx.CursorLocation = adUseClient
            CLMBordx.CursorType = adOpenForwardOnly
            CLMBordx.CacheSize = 100
            Set CLMBordx = cmd8.Execute
            
            Do Until CLMBordx.EOF
                DataString = ""
                DataString = "UMMB" & ","
                DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "") & ","
                DataString = DataString & "04A" & ","
                DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & "FINANCE" & ","
                DataString = DataString & IIf(Len(CLMBordx!ClinicCode), Trim(CLMBordx!ClinicCode), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceNo), Trim(CLMBordx!InvoiceNo), "") & ", "
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceDate), Trim(CLMBordx!InvoiceDate), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "Upload From MediExpress" & ","
                
                If Trim(CLMBordx!coverid) <> "00" Then
                    If Trim(CLMBordx!MBMCRELCode) = "W" Then
                        DataString = DataString & "I" & ","
                    ElseIf Trim(CLMBordx!MBMCRELCode) = "H" Then
                        DataString = DataString & "S" & ","
                    Else
                        DataString = DataString & "A" & ","
                    End If
                Else
                    DataString = DataString & "" & ","
                End If
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                If Trim(CLMBordx!coverid) <> "00" Then
                    DataString = DataString & IIf(Len(CLMBordx!MBMCName), Trim(CLMBordx!MBMCName), "") & ","
                End If
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & Chr(13)
                    
                    CLMBordx.MoveNext
                    
                    Print #1, DataString
                    Print #2, DataString
                Loop
                
                'End If
                
                CLMBordx.Close
                Set CLMBordx = Nothing
            
            
            Set cmd9.ActiveConnection = MediConnCISWS
            cmd9.CommandText = "ClaimDetails_ASCII_Utusan"
            cmd9.CommandType = adCmdStoredProc
            cmd9.CommandTimeout = 300
            Set Prm17 = cmd9.CreateParameter("@strAppend", adVarChar, adParamInput, 200, Trim(StrAppend))
            cmd9.Parameters.Append Prm17
            Set Prm18 = cmd9.CreateParameter("@strAppendCase", adVarChar, adParamInput, 1, 9)
            cmd9.Parameters.Append Prm18
            CLMBordx.CursorLocation = adUseClient
            CLMBordx.CursorType = adOpenForwardOnly
            CLMBordx.CacheSize = 100
            Set CLMBordx = cmd9.Execute
            
            Do Until CLMBordx.EOF
                DataString = ""
                DataString = "UMMB" & ","
                DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "") & ","
                DataString = DataString & "04B" & ","
                DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & "FINANCE" & ","
                DataString = DataString & IIf(Len(CLMBordx!ClinicCode), Trim(CLMBordx!ClinicCode), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceNo), Trim(CLMBordx!InvoiceNo), "") & ", "
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceDate), Trim(CLMBordx!InvoiceDate), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "Upload From MediExpress" & ","
                
                If Trim(CLMBordx!coverid) <> "00" Then
                    If Trim(CLMBordx!MBMCRELCode) = "W" Then
                        DataString = DataString & "I" & ","
                    ElseIf Trim(CLMBordx!MBMCRELCode) = "H" Then
                        DataString = DataString & "S" & ","
                    Else
                        DataString = DataString & "A" & ","
                    End If
                Else
                    DataString = DataString & "" & ","
                End If
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                If Trim(CLMBordx!coverid) <> "00" Then
                    DataString = DataString & IIf(Len(CLMBordx!MBMCName), Trim(CLMBordx!MBMCName), "") & ","
                End If
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & Chr(13)
                    
                    CLMBordx.MoveNext
                    
                    Print #1, DataString
                    Print #2, DataString
                Loop
                
                'End If
                
                CLMBordx.Close
                Set CLMBordx = Nothing
            
            
            Set cmd10.ActiveConnection = MediConnCISWS
            cmd10.CommandText = "ClaimDetails_ASCII_Utusan"
            cmd10.CommandType = adCmdStoredProc
            cmd10.CommandTimeout = 300
            Set Prm19 = cmd10.CreateParameter("@strAppend", adVarChar, adParamInput, 200, Trim(StrAppend))
            cmd10.Parameters.Append Prm19
            Set Prm20 = cmd10.CreateParameter("@strAppendCase", adVarChar, adParamInput, 10, 10)
            cmd10.Parameters.Append Prm20
            CLMBordx.CursorLocation = adUseClient
            CLMBordx.CursorType = adOpenForwardOnly
            CLMBordx.CacheSize = 100
            Set CLMBordx = cmd10.Execute
            
            Do Until CLMBordx.EOF
                DataString = ""
                DataString = "UMMB" & ","
                DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "") & ","
                DataString = DataString & "09" & ","
                DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & "FINANCE" & ","
                DataString = DataString & IIf(Len(CLMBordx!ClinicCode), Trim(CLMBordx!ClinicCode), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceNo), Trim(CLMBordx!InvoiceNo), "") & ", "
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceDate), Trim(CLMBordx!InvoiceDate), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "Upload From MediExpress" & ","
                
                If Trim(CLMBordx!coverid) <> "00" Then
                    If Trim(CLMBordx!MBMCRELCode) = "W" Then
                        DataString = DataString & "I" & ","
                    ElseIf Trim(CLMBordx!MBMCRELCode) = "H" Then
                        DataString = DataString & "S" & ","
                    Else
                        DataString = DataString & "A" & ","
                    End If
                Else
                    DataString = DataString & "" & ","
                End If
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                If Trim(CLMBordx!coverid) <> "00" Then
                    DataString = DataString & IIf(Len(CLMBordx!MBMCName), Trim(CLMBordx!MBMCName), "") & ","
                End If
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & Chr(13)
                    
                    CLMBordx.MoveNext
                    
                    Print #1, DataString
                    Print #2, DataString
                Loop
                
                'End If
                
                CLMBordx.Close
                Set CLMBordx = Nothing
            
            Set cmd11.ActiveConnection = MediConnCISWS
            cmd11.CommandText = "ClaimDetails_ASCII_Utusan"
            cmd11.CommandType = adCmdStoredProc
            cmd11.CommandTimeout = 300
            Set Prm21 = cmd11.CreateParameter("@strAppend", adVarChar, adParamInput, 200, Trim(StrAppend))
            cmd11.Parameters.Append Prm21
            Set Prm22 = cmd11.CreateParameter("@strAppendCase", adVarChar, adParamInput, 10, 11)
            cmd11.Parameters.Append Prm22
            CLMBordx.CursorLocation = adUseClient
            CLMBordx.CursorType = adOpenForwardOnly
            CLMBordx.CacheSize = 100
            Set CLMBordx = cmd11.Execute
            
            Do Until CLMBordx.EOF
                DataString = ""
                DataString = "UMMB" & ","
                DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "") & ","
                DataString = DataString & "10A" & ","
                DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & "FINANCE" & ","
                DataString = DataString & IIf(Len(CLMBordx!ClinicCode), Trim(CLMBordx!ClinicCode), "") & ","
                DataString = DataString & "" & ","
                
                If Trim(CLMBordx!CASTreatReason1) = "304" Then
                    DataString = DataString & "03" & ","
                ElseIf Trim(CLMBordx!CASTreatReason1) <> "304" And Trim(CLMBordx!CLNType) = "S" Then
                    DataString = DataString & "02" & ","
                Else
                    DataString = DataString & "01" & ","
                End If
                
                DataString = DataString & IIf(Len(CLMBordx!InvoiceNo), Trim(CLMBordx!InvoiceNo), "") & ", "
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceDate), Trim(CLMBordx!InvoiceDate), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "Upload From MediExpress" & ","
                
                If Trim(CLMBordx!coverid) <> "00" Then
                    If Trim(CLMBordx!MBMCRELCode) = "W" Then
                        DataString = DataString & "I" & ","
                    ElseIf Trim(CLMBordx!MBMCRELCode) = "H" Then
                        DataString = DataString & "S" & ","
                    Else
                        DataString = DataString & "A" & ","
                    End If
                Else
                    DataString = DataString & "" & ","
                End If
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                If Trim(CLMBordx!coverid) <> "00" Then
                    DataString = DataString & IIf(Len(CLMBordx!MBMCName), Trim(CLMBordx!MBMCName), "") & ","
                End If
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & Chr(13)
                    
                    CLMBordx.MoveNext
                    
                    Print #1, DataString
                    Print #2, DataString
                Loop
                
                'End If
                
                CLMBordx.Close
                Set CLMBordx = Nothing
                
            Set cmd12.ActiveConnection = MediConnCISWS
            cmd12.CommandText = "ClaimDetails_ASCII_Utusan"
            cmd12.CommandType = adCmdStoredProc
            cmd12.CommandTimeout = 300
            Set Prm23 = cmd12.CreateParameter("@strAppend", adVarChar, adParamInput, 200, Trim(StrAppend))
            cmd12.Parameters.Append Prm23
            Set Prm24 = cmd12.CreateParameter("@strAppendCase", adVarChar, adParamInput, 10, 12)
            cmd12.Parameters.Append Prm24
            CLMBordx.CursorLocation = adUseClient
            CLMBordx.CursorType = adOpenForwardOnly
            CLMBordx.CacheSize = 100
            Set CLMBordx = cmd12.Execute
            
            Do Until CLMBordx.EOF
                DataString = ""
                DataString = "UMMB" & ","
                DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "") & ","
                DataString = DataString & "11A" & ","
                DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & "FINANCE" & ","
                DataString = DataString & IIf(Len(CLMBordx!ClinicCode), Trim(CLMBordx!ClinicCode), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceNo), Trim(CLMBordx!InvoiceNo), "") & ", "
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceDate), Trim(CLMBordx!InvoiceDate), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "Upload From MediExpress" & ","
                
                If Trim(CLMBordx!coverid) <> "00" Then
                    If Trim(CLMBordx!MBMCRELCode) = "W" Then
                        DataString = DataString & "I" & ","
                    ElseIf Trim(CLMBordx!MBMCRELCode) = "H" Then
                        DataString = DataString & "S" & ","
                    Else
                        DataString = DataString & "A" & ","
                    End If
                Else
                    DataString = DataString & "" & ","
                End If
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                If Trim(CLMBordx!coverid) <> "00" Then
                    DataString = DataString & IIf(Len(CLMBordx!MBMCName), Trim(CLMBordx!MBMCName), "") & ","
                End If
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & Chr(13)
                    
                    CLMBordx.MoveNext
                    
                    Print #1, DataString
                    Print #2, DataString
                Loop
                
                'End If
                
                CLMBordx.Close
                Set CLMBordx = Nothing
                
            Set cmd13.ActiveConnection = MediConnCISWS
            cmd13.CommandText = "ClaimDetails_ASCII_Utusan"
            cmd13.CommandType = adCmdStoredProc
            cmd13.CommandTimeout = 300
            Set Prm25 = cmd13.CreateParameter("@strAppend", adVarChar, adParamInput, 200, Trim(StrAppend))
            cmd13.Parameters.Append Prm25
            Set Prm26 = cmd13.CreateParameter("@strAppendCase", adVarChar, adParamInput, 10, 13)
            cmd13.Parameters.Append Prm26
            CLMBordx.CursorLocation = adUseClient
            CLMBordx.CursorType = adOpenForwardOnly
            CLMBordx.CacheSize = 100
            Set CLMBordx = cmd13.Execute
            
            Do Until CLMBordx.EOF
                DataString = ""
                DataString = "UMMB" & ","
                DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "") & ","
                DataString = DataString & "11B" & ","
                DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & "FINANCE" & ","
                DataString = DataString & IIf(Len(CLMBordx!ClinicCode), Trim(CLMBordx!ClinicCode), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceNo), Trim(CLMBordx!InvoiceNo), "") & ", "
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceDate), Trim(CLMBordx!InvoiceDate), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "Upload From MediExpress" & ","
                
                If Trim(CLMBordx!coverid) <> "00" Then
                    If Trim(CLMBordx!MBMCRELCode) = "W" Then
                        DataString = DataString & "I" & ","
                    ElseIf Trim(CLMBordx!MBMCRELCode) = "H" Then
                        DataString = DataString & "S" & ","
                    Else
                        DataString = DataString & "A" & ","
                    End If
                Else
                    DataString = DataString & "" & ","
                End If
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                If Trim(CLMBordx!coverid) <> "00" Then
                    DataString = DataString & IIf(Len(CLMBordx!MBMCName), Trim(CLMBordx!MBMCName), "") & ","
                End If
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & Chr(13)
                    
                    CLMBordx.MoveNext
                    
                    Print #1, DataString
                    Print #2, DataString
                Loop
                
                'End If
                
                CLMBordx.Close
                Set CLMBordx = Nothing
                
                
            cmdGenerate.Enabled = False
            Close #1
            Close #2
                
        MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
        cmdGenerate.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub



Private Sub ClaimASCII()
Const METHOD_NAME = "ClaimASCII"
On Error GoTo debugs
        tmpstring = ""
        z = False
        
        tmpBordxDay = ""
        tmpBordxMonth = ""
        tmpBordxYear = ""
        StrAppend = " AND BordxRefNo >= '" & Trim(txtBordxNoFr) & "'" & _
                    "AND BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
        
        
        'sql = "SELECT * From View_ClaimDetails_ASCII where 0 = 0 "
        
        'strsql = sql + StrAppend
        'CLMBordx.Open strsql, MediConnCISWS, adOpenStatic, adLockOptimistic


    tmpBordxDay = Day(Date)
    If Len(Trim(tmpBordxDay)) = 1 Then
        tmpBordxDay = "0" + tmpBordxDay
    End If
    tmpBordxMonth = Month(Date)
    If Len(Trim(tmpBordxMonth)) = 1 Then
        tmpBordxMonth = "0" + tmpBordxMonth
    End If
    'tmpBordxYear = Format(Year(CLMBordx!BordxDate), "yy")
    tmpBordxYear = Mid$(Date, 7, 4)
    
    txtFilName = tmpBordxYear & tmpBordxMonth & tmpBordxDay & "_BPOCLMTRH.KZULF"
    z = CreatePath(txtLocalFileLocation)
    
    Open txtPath & "\" & txtFilName For Output As #1
    Open txtLocalFileLocation & txtFilName For Output As #2
    
    'Header = txtFilName
    
    'Print #1, Trim(Header)
    'Print #2, Trim(Header)

            
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimDetails_ASCII_Utusan"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("@strAppend", adVarChar, adParamInput, 200, Trim(StrAppend))
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("@strAppendCase", adVarChar, adParamInput, 1, 1)
    cmd.Parameters.Append Prm2
    CLMBordx.CursorLocation = adUseClient
    CLMBordx.CursorType = adOpenForwardOnly
    CLMBordx.CacheSize = 100
    Set CLMBordx = cmd.Execute
    
       ' If CLMBordx.EOF Then
        '    MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'Else
           
            
        Do Until CLMBordx.EOF
            DataString = ""
            DataString = "UMMB" & ","
            DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "") & ","
            DataString = DataString & "01A" & ","
            DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & ","
            DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
            DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
            DataString = DataString & "FINANCE" & ","
            DataString = DataString & IIf(Len(CLMBordx!ClinicCode), Trim(CLMBordx!ClinicCode), "") & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & IIf(Len(CLMBordx!InvoiceNo), Trim(CLMBordx!InvoiceNo), "") & ", "
            DataString = DataString & "" & ","
            DataString = DataString & IIf(Len(CLMBordx!InvoiceDate), Trim(CLMBordx!InvoiceDate), "") & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "Upload From MediExpress" & ","
            
            If Trim(CLMBordx!coverid) <> "00" Then
                If Trim(CLMBordx!MBMCRELCode) = "W" Then
                    DataString = DataString & "I" & ","
                ElseIf Trim(CLMBordx!MBMCRELCode) = "H" Then
                    DataString = DataString & "S" & ","
                Else
                    DataString = DataString & "A" & ","
                End If
            Else
                DataString = DataString & "" & ","
            End If
            DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
            If Trim(CLMBordx!coverid) <> "00" Then
                DataString = DataString & IIf(Len(CLMBordx!MBMCName), Trim(CLMBordx!MBMCName), "") & ","
            End If
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & Chr(13)
            
            txtTotalVisits = txtTotalVisits + 1
            'lblTotalAmt = CDbl(lblTotalAmt) + IIf(IsNumeric(CLMBordx!CLMTotalActual), CDbl(CLMBordx!CLMTotalActual), 0)

            CLMBordx.MoveNext
            
            Print #1, DataString
            Print #2, DataString
        Loop
            
        'End If
            
        CLMBordx.Close
        Set CLMBordx = Nothing
        
        Set cmd2.ActiveConnection = MediConnCISWS
        cmd2.CommandText = "ClaimDetails_ASCII_Utusan"
        cmd2.CommandType = adCmdStoredProc
        cmd2.CommandTimeout = 300
        Set Prm3 = cmd2.CreateParameter("@strAppend", adVarChar, adParamInput, 200, Trim(StrAppend))
        cmd2.Parameters.Append Prm3
        Set Prm4 = cmd2.CreateParameter("@strAppendCase", adVarChar, adParamInput, 1, 2)
        cmd2.Parameters.Append Prm4
        CLMBordx.CursorLocation = adUseClient
        CLMBordx.CursorType = adOpenForwardOnly
        CLMBordx.CacheSize = 100
        Set CLMBordx = cmd2.Execute
        
        'If CLMBordx.EOF Then
            Do Until CLMBordx.EOF
                DataString = ""
                DataString = "UMMB" & ","
                DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "") & ","
                DataString = DataString & "01B" & ","
                DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & "FINANCE" & ","
                DataString = DataString & IIf(Len(CLMBordx!ClinicCode), Trim(CLMBordx!ClinicCode), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceNo), Trim(CLMBordx!InvoiceNo), "") & ", "
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceDate), Trim(CLMBordx!InvoiceDate), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "Upload From MediExpress" & ","
                
                If Trim(CLMBordx!coverid) <> "00" Then
                    If Trim(CLMBordx!MBMCRELCode) = "W" Then
                        DataString = DataString & "I" & ","
                    ElseIf Trim(CLMBordx!MBMCRELCode) = "H" Then
                        DataString = DataString & "S" & ","
                    Else
                        DataString = DataString & "A" & ","
                    End If
                Else
                    DataString = DataString & "" & ","
                End If
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                If Trim(CLMBordx!coverid) <> "00" Then
                    DataString = DataString & IIf(Len(CLMBordx!MBMCName), Trim(CLMBordx!MBMCName), "") & ","
                End If
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & Chr(13)
                
                    CLMBordx.MoveNext
                    
                    Print #1, DataString
                    Print #2, DataString
                Loop
                
           ' End If
            
            CLMBordx.Close
            Set CLMBordx = Nothing

        Set cmd3.ActiveConnection = MediConnCISWS
        cmd3.CommandText = "ClaimDetails_ASCII_Utusan"
        cmd3.CommandType = adCmdStoredProc
        cmd3.CommandTimeout = 300
        Set Prm5 = cmd3.CreateParameter("@strAppend", adVarChar, adParamInput, 200, Trim(StrAppend))
        cmd3.Parameters.Append Prm5
        Set Prm6 = cmd3.CreateParameter("@strAppendCase", adVarChar, adParamInput, 1, 3)
        cmd3.Parameters.Append Prm6
        CLMBordx.CursorLocation = adUseClient
        CLMBordx.CursorType = adOpenForwardOnly
        CLMBordx.CacheSize = 100
        Set CLMBordx = cmd3.Execute
        
        'If CLMBordx.EOF Then
            Do Until CLMBordx.EOF
                DataString = ""
            DataString = "UMMB" & ","
            DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "") & ","
            DataString = DataString & "01C" & ","
            DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & ","
            DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
            DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
            DataString = DataString & "FINANCE" & ","
            DataString = DataString & IIf(Len(CLMBordx!ClinicCode), Trim(CLMBordx!ClinicCode), "") & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & IIf(Len(CLMBordx!InvoiceNo), Trim(CLMBordx!InvoiceNo), "") & ", "
            DataString = DataString & "" & ","
            DataString = DataString & IIf(Len(CLMBordx!InvoiceDate), Trim(CLMBordx!InvoiceDate), "") & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "Upload From MediExpress" & ","
            
            If Trim(CLMBordx!coverid) <> "00" Then
                If Trim(CLMBordx!MBMCRELCode) = "W" Then
                    DataString = DataString & "I" & ","
                ElseIf Trim(CLMBordx!MBMCRELCode) = "H" Then
                    DataString = DataString & "S" & ","
                Else
                    DataString = DataString & "A" & ","
                End If
            Else
                DataString = DataString & "" & ","
            End If
            DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
            If Trim(CLMBordx!coverid) <> "00" Then
                DataString = DataString & IIf(Len(CLMBordx!MBMCName), Trim(CLMBordx!MBMCName), "") & ","
            End If
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & Chr(13)
                
                txtTotalVisits = txtTotalVisits + 1
                'lblTotalAmt = CDbl(lblTotalAmt) + IIf(IsNumeric(CLMBordx!CLMTotalActual), CDbl(CLMBordx!CLMTotalActual), 0)
    
                CLMBordx.MoveNext
                
                Print #1, DataString
                Print #2, DataString
            Loop
            
           ' End If
            CLMBordx.Close
            Set CLMBordx = Nothing
            
        Set cmd4.ActiveConnection = MediConnCISWS
        cmd4.CommandText = "ClaimDetails_ASCII_Utusan"
        cmd4.CommandType = adCmdStoredProc
        cmd4.CommandTimeout = 300
        Set Prm7 = cmd4.CreateParameter("@strAppend", adVarChar, adParamInput, 200, Trim(StrAppend))
        cmd4.Parameters.Append Prm7
        Set Prm8 = cmd4.CreateParameter("@strAppendCase", adVarChar, adParamInput, 1, 4)
        cmd4.Parameters.Append Prm8
        CLMBordx.CursorLocation = adUseClient
        CLMBordx.CursorType = adOpenForwardOnly
        CLMBordx.CacheSize = 100
        Set CLMBordx = cmd4.Execute
        
            'If CLMBordx.EOF Then
    
        Do Until CLMBordx.EOF
            DataString = ""
            DataString = "UMMB" & ","
            DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "") & ","
            DataString = DataString & "02A" & ","
            DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & ","
            DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
            DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
            DataString = DataString & "FINANCE" & ","
            DataString = DataString & IIf(Len(CLMBordx!ClinicCode), Trim(CLMBordx!ClinicCode), "") & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & IIf(Len(CLMBordx!InvoiceNo), Trim(CLMBordx!InvoiceNo), "") & ", "
            DataString = DataString & "" & ","
            DataString = DataString & IIf(Len(CLMBordx!InvoiceDate), Trim(CLMBordx!InvoiceDate), "") & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "Upload From MediExpress" & ","
            
            If Trim(CLMBordx!coverid) <> "00" Then
                If Trim(CLMBordx!MBMCRELCode) = "W" Then
                    DataString = DataString & "I" & ","
                ElseIf Trim(CLMBordx!MBMCRELCode) = "H" Then
                    DataString = DataString & "S" & ","
                Else
                    DataString = DataString & "A" & ","
                End If
            Else
                DataString = DataString & "" & ","
            End If
            DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
            If Trim(CLMBordx!coverid) <> "00" Then
                DataString = DataString & IIf(Len(CLMBordx!MBMCName), Trim(CLMBordx!MBMCName), "") & ","
            End If
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & Chr(13)
                
                CLMBordx.MoveNext
                
                Print #1, DataString
                Print #2, DataString
            Loop
            
            'End If
            
            CLMBordx.Close
            Set CLMBordx = Nothing
            
            
            Set cmd5.ActiveConnection = MediConnCISWS
            cmd5.CommandText = "ClaimDetails_ASCII_Utusan"
            cmd5.CommandType = adCmdStoredProc
            cmd5.CommandTimeout = 300
            Set Prm9 = cmd5.CreateParameter("@strAppend", adVarChar, adParamInput, 200, Trim(StrAppend))
            cmd5.Parameters.Append Prm9
            Set Prm10 = cmd5.CreateParameter("@strAppendCase", adVarChar, adParamInput, 1, 5)
            cmd5.Parameters.Append Prm10
            CLMBordx.CursorLocation = adUseClient
            CLMBordx.CursorType = adOpenForwardOnly
            CLMBordx.CacheSize = 100
            Set CLMBordx = cmd5.Execute
            
                'If CLMBordx.EOF Then
        
            Do Until CLMBordx.EOF
                DataString = ""
                DataString = "UMMB" & ","
                DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "") & ","
                DataString = DataString & "02B" & ","
                DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & "FINANCE" & ","
                DataString = DataString & IIf(Len(CLMBordx!ClinicCode), Trim(CLMBordx!ClinicCode), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceNo), Trim(CLMBordx!InvoiceNo), "") & ", "
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceDate), Trim(CLMBordx!InvoiceDate), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "Upload From MediExpress" & ","
                
                If Trim(CLMBordx!coverid) <> "00" Then
                    If Trim(CLMBordx!MBMCRELCode) = "W" Then
                        DataString = DataString & "I" & ","
                    ElseIf Trim(CLMBordx!MBMCRELCode) = "H" Then
                        DataString = DataString & "S" & ","
                    Else
                        DataString = DataString & "A" & ","
                    End If
                Else
                    DataString = DataString & "" & ","
                End If
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                If Trim(CLMBordx!coverid) <> "00" Then
                    DataString = DataString & IIf(Len(CLMBordx!MBMCName), Trim(CLMBordx!MBMCName), "") & ","
                End If
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & Chr(13)
                    
                    CLMBordx.MoveNext
                    
                    Print #1, DataString
                    Print #2, DataString
                Loop
                
                'End If
                
                CLMBordx.Close
                Set CLMBordx = Nothing
                                  
            Set cmd6.ActiveConnection = MediConnCISWS
            cmd6.CommandText = "ClaimDetails_ASCII_Utusan"
            cmd6.CommandType = adCmdStoredProc
            cmd6.CommandTimeout = 300
            Set Prm11 = cmd6.CreateParameter("@strAppend", adVarChar, adParamInput, 200, Trim(StrAppend))
            cmd6.Parameters.Append Prm11
            Set Prm12 = cmd6.CreateParameter("@strAppendCase", adVarChar, adParamInput, 1, 6)
            cmd6.Parameters.Append Prm12
            CLMBordx.CursorLocation = adUseClient
            CLMBordx.CursorType = adOpenForwardOnly
            CLMBordx.CacheSize = 100
            Set CLMBordx = cmd6.Execute
            
                'If CLMBordx.EOF Then
        
            Do Until CLMBordx.EOF
                DataString = ""
                DataString = "UMMB" & ","
                DataString = DataString & IIf(Len(CLMBordx!MBMEmployeeNo), Trim(CLMBordx!MBMEmployeeNo), "") & ","
                DataString = DataString & "03A" & ","
                DataString = DataString & IIf(Len(CLMBordx!DateVisit), Trim(CLMBordx!DateVisit), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                DataString = DataString & "FINANCE" & ","
                DataString = DataString & IIf(Len(CLMBordx!ClinicCode), Trim(CLMBordx!ClinicCode), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceNo), Trim(CLMBordx!InvoiceNo), "") & ", "
                DataString = DataString & "" & ","
                DataString = DataString & IIf(Len(CLMBordx!InvoiceDate), Trim(CLMBordx!InvoiceDate), "") & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "Upload From MediExpress" & ","
                
                If Trim(CLMBordx!coverid) <> "00" Then
                    If Trim(CLMBordx!MBMCRELCode) = "W" Then
                        DataString = DataString & "I" & ","
                    ElseIf Trim(CLMBordx!MBMCRELCode) = "H" Then
                        DataString = DataString & "S" & ","
                    Else
                        DataString = DataString & "A" & ","
                    End If
                Else
                    DataString = DataString & "" & ","
                End If
                DataString = DataString & IIf(Len(CLMBordx!CLMBordxAmt), Format(CLMBordx!CLMBordxAmt, "###0.00"), "") & ","
                If Trim(CLMBordx!coverid) <> "00" Then
                    DataString = DataString & IIf(Len(CLMBordx!MBMCName), Trim(CLMBordx!MBMCName), "") & ","
                End If
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & ","
                DataString = DataString & "" & Chr(13)
                    
                    CLMBordx.MoveNext
                    
                    Print #1, DataString
                    Print #2, DataString
                Loop
                
                'End If
                
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
    Call ActionLogUpdate(12, "frmClaimASCIIUtusan")
    txtPath = SRVPath & "ClinicalClaimASCII"
    txtLocalFileLocation = "C:\ClinicalClaimsASCII\"
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
