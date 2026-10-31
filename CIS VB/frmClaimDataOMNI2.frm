VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmClaimDataOMNI 
   Caption         =   "Generate Claim ASCII - OMNI"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2505
   ScaleWidth      =   7170
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Height          =   1575
      Left            =   80
      TabIndex        =   7
      Top             =   240
      Width           =   7095
      Begin VB.ComboBox CboExportStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1920
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox txtBordxNoFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtPaymentDateFr 
         Height          =   285
         Left            =   6840
         TabIndex        =   19
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPaymentDateTo 
         Height          =   285
         Left            =   8760
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   8400
         TabIndex        =   22
         Top             =   1110
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Reg. Date Fr"
         Height          =   255
         Left            =   5520
         TabIndex        =   21
         Top             =   1110
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label txtFilName 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   840
         Width           =   3225
      End
      Begin VB.Label Label1 
         Caption         =   "Export Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label txtLocalFileLocation 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   540
         Width           =   3225
      End
      Begin VB.Label Label24 
         Caption         =   "File on Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label txtPath 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   3225
      End
      Begin VB.Label Label28 
         Caption         =   "File on PC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   615
         Width           =   765
      End
      Begin VB.Label Label25 
         Caption         =   "Export File Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   900
         Width           =   1245
      End
      Begin VB.Label Label26 
         Caption         =   "Bordx No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   80
      TabIndex        =   0
      Top             =   1800
      Width           =   7095
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6240
         TabIndex        =   1
         Top             =   240
         Width           =   705
      End
      Begin VB.Label txtTotalVisits 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label17 
         Caption         =   "Total Visit(s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         TabIndex        =   5
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label lblTotalAmt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   1920
      End
      Begin VB.Label Label17 
         Caption         =   "Total Bordx Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         TabIndex        =   3
         Top             =   270
         Width           =   1200
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Generate Claim ASCII - OMNI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
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
Attribute VB_Name = "frmClaimDataOMNI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    Dim CLMBordx As New ADODB.Recordset
    Dim tmpcnt As Double
    Dim Trailer As String
    Dim tmpfilename As String
    Dim tmpLocalFileLocation As String
    Dim DataString As String
    Dim Header As String
    Dim tmpTotalAct As String
    Dim tmpTotalApp As String
    
    Dim cmd As New ADODB.Command
    Dim Prm1 As New ADODB.Parameter
    Dim Prm2 As New ADODB.Parameter
    Dim Total As Integer
    Dim tmpDiagnosis As String
    Dim X As Integer
    Dim StrSql As String
    Dim tmpMBMCovID As String
    Dim tmpproceed As Boolean
    
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    tmpcnt = 0
    tmpTotalAct = 0
    tmpTotalApp = 0
    tmpDiagnosis = ""
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimDetails_ASCII_OMNI"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("BordxRefNoFrom", adVarChar, adParamInput, 20, Trim(txtBordxNoFr))
    cmd.Parameters.Append Prm1
    CLMBordx.CursorLocation = adUseClient
    CLMBordx.CursorType = adOpenForwardOnly
    CLMBordx.CacheSize = 100
    Set CLMBordx = cmd.Execute
    
    
        If CLMBordx.EOF Then
            MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            txtLocalFileLocation = "C:\MedicaGenClaimASCII\"
            
            'txtFilName = Date & "_OUTPATIENTPROVIDER"
            'Open txtLocalFileLocation & txtFilName & ".txt" For Output As #1
            txtFilName = Format(Date, "yyyymmdd")
            txtFilName = txtFilName & "_OUTPATIENTCLAIM" & ".txt"
            Open txtLocalFileLocation & "\" & txtFilName For Output As #1
            
            Header = "01" & "|" & "TPA" & "|" & Format(Date, "yyyymmdd")
                'Debug.Print Trailer
            Print #1, Header
            
        Do Until CLMBordx.EOF
        
        tmpproceed = False
        If Trim(CLMBordx!coverid) = "00" Then
            If Trim(CLMBordx!MBMCLNStatus) = "A" Then
                tmpproceed = True
            Else
                tmpproceed = False
            End If
        ElseIf Trim(CLMBordx!coverid) <> "00" Then
            If Trim(CLMBordx!MBMCCLNStatus) = "A" Then
                tmpproceed = True
            Else
                tmpproceed = False
            End If
        End If
        'If CLMBordx!coverid = "00" Then
        If tmpproceed = True Then
            If CLMBordx!coverid = "00" Then
                tmpMBMCovID = IIf(Len(CLMBordx!MBMPrevMemNo), CLMBordx!MBMPrevMemNo, "")
            Else
                tmpMBMCovID = IIf(Len(CLMBordx!MBMAXANumber), CLMBordx!MBMAXANumber, "")
            End If
             If tmpMBMCovID <> "" Then
                tmpcnt = tmpcnt + 1
           ' DataString = ""
                DataString = "02" & "|"
                
                DataString = DataString & "N" & "|"
                
                'DataString = DataString & "C" & "|"
                
                If Trim(CLMBordx!CLNType) = "G" Then
                    DataString = DataString & "OPGP" & "|"
                Else
                    DataString = DataString & "OPSP" & "|"
                End If
                
                DataString = DataString & CLMBordx!CasNumber & "|"
                
                DataString = DataString & "" & "|"
                
                DataString = DataString & tmpMBMCovID & "|"
                
                DataString = DataString & Format(CLMBordx!DateVisit, "YYYYMMDD") & "|"
                
                'DataString = DataString & Format(CLMBordx!BordxDate, "YYYYMMDD") & "|"
                'DataString = DataString & "20130506" & "|"
                DataString = DataString & Format(CLMBordx!DateVisit, "YYYYMMDD") & "|"
                
                If CLMBordx!INVPayTo = "M" Then
                    DataString = DataString & "2" & "|"
                Else
                    DataString = DataString & "1" & "|"
                End If
                
                DataString = DataString & Trim(CLMBordx!ClinicCode) & "|"
                
                DataString = DataString & Trim(CLMBordx!CLNName) & "|"
                
                DataString = DataString & "{ZZZ999}" & "|"
                
                If CLMBordx!CASTreatReason5 <> "" Then
                    If CLMBordx!CASDiagnosis5 <> "" Then
                        
                        
                        'CLMBordx!CASDiagnosis5(KeyAscii As Integer)


                        tmpDiagnosis = ""
                      '  For x = 1 To Len(CLMBordx!CASDiagnosis5)
                      '      If Asc(Mid(CLMBordx!CASDiagnosis5, x, 1)) <> 13 Then tmpDiagnosis = tmpDiagnosis & Mid(CLMBordx!CASDiagnosis5, x, 1)
                      '
                      '      'Text2.Text = Noenterstr
                      '      'End If
                      '  Next


                        DataString = DataString & "{" & CLMBordx!CASDiagnosis5 & "}" & "|"
                    Else
                        DataString = DataString & "{" & CLMBordx!ReasonDesc & "}" & "|"
                        
                   '
                   '     For x = 1 To Len(CLMBordx!ReasonDesc)
                   '         If Asc(Mid(CLMBordx!ReasonDesc, x, 1)) <> 13 Then tmpDiagnosis = tmpDiagnosis & Mid(CLMBordx!ReasonDesc, x, 1)
                   '
                   '         'Text2.Text = Noenterstr
                   '         'End If
                   '     Next
                    
                       ' DataString = DataString & "{" & tmpDiagnosis & "}" & "|"
                    End If
                Else
                   '     For x = 1 To Len(CLMBordx!ReasonDesc)
                   '         If Asc(Mid(CLMBordx!ReasonDesc, x, 1)) <> 13 Then tmpDiagnosis = tmpDiagnosis & Mid(CLMBordx!ReasonDesc, x, 1)
                   '
                   '         'Text2.Text = Noenterstr
                   '         'End If
                   '     Next
                    
                    DataString = DataString & "{" & CLMBordx!ReasonDesc & "}" & "|"
                End If
                
                DataString = DataString & "{" & CLMBordx!MedInjType1 & "}" & "|"
                                                              
                DataString = DataString & "" & "|"
                                
                DataString = DataString & CLMBordx!CLMTotalActual & "|"
                tmpTotalAct = CDbl(tmpTotalAct) + CDbl(CLMBordx!CLMTotalActual)
    

                DataString = DataString & CLMBordx!CLMTotalActual & "|"
                tmpTotalApp = CDbl(tmpTotalApp) + CDbl(CLMBordx!CLMTotalActual)
                
                DataString = DataString & "0" & "|"
                
                If CLMBordx!GPMed <> 0 Then
                    DataString = DataString & CLMBordx!GPMed & "|"
                ElseIf CLMBordx!SPMed <> 0 Then
                    DataString = DataString & CLMBordx!SPMed & "|"
                Else
                    DataString = DataString & "0" & "|"
                End If
                
                If CLMBordx!GPMed <> 0 Then
                    DataString = DataString & CLMBordx!GPMed & "|"
                ElseIf CLMBordx!SPMed <> 0 Then
                    DataString = DataString & CLMBordx!SPMed & "|"
                Else
                    DataString = DataString & "0" & "|"
                End If
                
                DataString = DataString & "0" & "|"
                
                If CLMBordx!GPMed <> 0 Then
                    DataString = DataString & CDbl(CLMBordx!CLMTotalActual) - CDbl(CLMBordx!GPMed) & "|"
                Else
                    If IsNumeric(CLMBordx!SPMed) Then
                        DataString = DataString & CDbl(CLMBordx!CLMTotalActual) - CDbl(CLMBordx!SPMed) & "|"
                    Else
                        DataString = DataString & CDbl(CLMBordx!CLMTotalActual) & "|"
                    End If
                End If
                
                If CLMBordx!GPMed <> 0 Then
                    DataString = DataString & CDbl(CLMBordx!CLMTotalActual) - CDbl(CLMBordx!GPMed) & "|"
                Else
                    If IsNumeric(CLMBordx!SPMed) Then
                        DataString = DataString & CDbl(CLMBordx!CLMTotalActual) - CDbl(CLMBordx!SPMed) & "|"
                    Else
                        DataString = DataString & CDbl(CLMBordx!CLMTotalActual) & "|"
                    End If
                End If
                
                
                DataString = DataString & "0" & "|"
                
                DataString = DataString & "0" & "|"
                
                DataString = DataString & "0" & "|"
                
                DataString = DataString & "0" & "|"
                
                DataString = DataString & "0" & "|"
                
                DataString = DataString & "0" & "|"
                
                DataString = DataString & "0" & "|"
                
                DataString = DataString & "0" & "|"
                
                DataString = DataString & "0" & "|"
                
                DataString = DataString & "0" & "|"
                
                DataString = DataString & "0" & "|"
                
                DataString = DataString & Format(Date, "yyyymmdd") & " 01:00:00" & "|"
                
                DataString = DataString & "" & "|"
                
                DataString = DataString & "N" & "|"
                
                DataString = DataString & "" & "|"
                If CLMBordx!MCDays = "0" Then
                    DataString = DataString & "N" & "|"
                    
                    DataString = DataString & "" & "|"
                    
                    DataString = DataString & "" & "|"
                    
                    DataString = DataString & ""
                
                Else
                    DataString = DataString & "Y" & "|"
                
                    DataString = DataString & Format(CLMBordx!DateVisit, "YYYYMMDD") & "|"
                    
                    DataString = DataString & Format(CLMBordx!DateVisit, "YYYYMMDD") & "|"
                    
                    If CLMBordx!MCDays = "1" Then
                        DataString = DataString & Format(CLMBordx!DateVisit, "YYYYMMDD")
                    Else
                        DataString = DataString & Format(CLMBordx!DateVisit + CLMBordx!MCDays, "YYYYMMDD")
                    End If
                End If
                
                
            
                Print #1, DataString
                
                DataString = DataString & Chr(13)
            End If
            
            End If
            CLMBordx.MoveNext
            
            Loop
            
           
                CLMBordx.Close
                Set CLMBordx = Nothing

                Trailer = "03" & "|" & tmpcnt & "|" & tmpTotalAct & "|" & tmpTotalApp & "|" & "0"
                'Debug.Print Trailer
                Print #1, Trailer
                'Print #2, Trailer
                'cmdGenerate.Enabled = False
                Close #1
                
                lblTotalAmt = tmpTotalApp
                txtTotalVisits = tmpcnt
            
                Call CLNFile
            
                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                cmdGenerate.Enabled = True
        End If

End Sub

Private Sub CLNFile()
Dim HOS As New ADODB.Recordset
Dim PAY As New ADODB.Recordset
Dim tmpcnt As Double
Dim Trailer As String
Dim tmpfilename As String
Dim tmpLocalFileLocation As String
Dim DataString As String
Dim Header As String
Dim StrSql As String

    tmpcnt = 0
    
    'HOS.Open "select * from View_CLN_OMNI where BordxRefNo = '" & Trim(txtBordxNoFr) & "'", MediConnCISDB, adOpenKeyset, adLockOptimistic
        
    HOS.Open "SELECT * From CLN WHERE  (CLNPanel = 'Y' OR CLNPanel = 'C') AND (CLNOMNISent IS NULL)", medixmasterconn, adOpenKeyset, adLockOptimistic

    If HOS.EOF Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        txtLocalFileLocation = "C:\MedicaGenClaimASCII\"

        txtFilName = Format(Date, "yyyymmdd")
        txtFilName = txtFilName & "_OUTPATIENTPROVIDER" & ".txt"
        Open txtLocalFileLocation & "\" & txtFilName For Output As #1
        
        Header = "01" & "|" & "TPA" & "|" & Format(Date, "yyyymmdd")
            'Debug.Print Trailer
        Print #1, Header
            
        Do Until HOS.EOF
             
                tmpcnt = tmpcnt + 1
                DataString = "02" & "|"
                
                DataString = DataString & "N" & "|"
                           
                DataString = DataString & Trim(HOS!CLNCode) & "|"
                
                DataString = DataString & Trim(HOS!CLNName) & "|"
                
                DataString = DataString & Trim(HOS!CLNPhone) & "|"
                
                DataString = DataString & Trim(HOS!CLNAdd1) & "|"
                
                DataString = DataString & Trim(HOS!CLNAdd2) & "|"
                
                DataString = DataString & Trim(HOS!CLNAdd3) & "|"
                
                DataString = DataString & Trim(HOS!CLNCity) & "|"
                
                DataString = DataString & Trim(HOS!CLNPostCode) & "|"
                
                DataString = DataString & "" & "|"

                DataString = DataString & "MALAYSIA" & "|"
                
                DataString = DataString & "2" & "|"
                
                DataString = DataString & "N" & "|"
                
                DataString = DataString & "Y"
                

                
                Print #1, DataString
                
                DataString = DataString & Chr(13)
                
                
                
                StrSql = " update CLN set  CLNOMNISent = 'Y' Where CLNCode = '" & Trim(HOS!CLNCode) & "'"
                medixmasterconn.Execute StrSql
                
            HOS.MoveNext
            Loop

                Trailer = "03" & "|" & tmpcnt
                'Debug.Print Trailer
                Print #1, Trailer
                Close #1
                
                'MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                'cmdGenerate.Enabled = True
        End If
    HOS.Close
    Set HOS = Nothing


End Sub

Private Sub txtPaymentDateFr_LostFocus()
    If txtPaymentDateFr <> "__/__/____" Then
        If IsDate(txtPaymentDateFr) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtPaymentDateFr.SetFocus
        End If
    End If
End Sub

Private Sub txtPaymentDateTo_LostFocus()
    If txtPaymentDateTo <> "__/__/____" Then
        If IsDate(txtPaymentDateTo) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtPaymentDateTo.SetFocus
        End If
    End If

End Sub
