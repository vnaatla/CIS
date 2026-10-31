VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmClaimASCIICIMB 
   Caption         =   "Generate Claim ASCII - CIMB"
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
      TabIndex        =   15
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
      Begin VB.ComboBox CboCashType 
         Height          =   315
         Left            =   3840
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   1200
         Width           =   1320
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
         Top             =   1200
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
      Begin MSMask.MaskEdBox txtPaymentDateFr 
         Height          =   285
         Left            =   6000
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPaymentDateTo 
         Height          =   285
         Left            =   8760
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Cash Type"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   24
         Top             =   1245
         Width           =   855
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
         TabIndex        =   14
         Top             =   1200
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   240
         Width           =   3705
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   540
         Width           =   3705
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
         Index           =   0
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
         Width           =   3705
      End
      Begin VB.Label Label11 
         Caption         =   "Reg. Date Fr"
         Height          =   255
         Left            =   6120
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   8400
         TabIndex        =   5
         Top             =   1110
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Generate Claim ASCII - CIMB"
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
      TabIndex        =   22
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmClaimASCIICIMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
Const METHOD_NAME = "cmdGenerate_Click"
On Error GoTo debugs
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
    Dim total As Integer
    Dim tmpDiagnosis As String
    Dim X As Integer
    Dim StrSql As String
    Dim tmpMBMCovID As String
    
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    tmpcnt = 0
    tmpTotalAct = 0
    tmpTotalApp = 0
    tmpDiagnosis = ""
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimDetails_ASCII_CIMB"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("BordxRefNoFrom", adVarChar, adParamInput, 20, Trim(txtBordxNoFr))
    cmd.Parameters.Append Prm1
    CLMBordx.CursorLocation = adUseClient
    CLMBordx.CursorType = adOpenForwardOnly
    CLMBordx.CacheSize = 100
    Set CLMBordx = cmd.Execute
   
       ' strsql = ""
       ' strsql = " And CISDB.dbo.VIEW_CISCAS_Patient.GRPOMNICLM = 'Y' And CISDB.dbo.VIEW_CISCAS_Patient.RegisterDate >= '" & Format(txtPaymentDateFr, "yyyy/mm/dd") & "' and  CISDB.dbo.VIEW_CISCAS_Patient.RegisterDate <= '" & Format(txtPaymentDateTo, "yyyy/mm/dd") & "'"
       ' Set cmd.ActiveConnection = MediConnCISWS
       ' cmd.CommandText = "Claim_ASCII_OMNIRegDate"
       ' cmd.CommandType = adCmdStoredProc
       ' cmd.CommandTimeout = 300
       ' Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 3000, strsql)
       ' cmd.Parameters.Append Prm1
       ' CLMBordx.CursorLocation = adUseClient
       ' CLMBordx.CursorType = adOpenForwardOnly
      '  CLMBordx.CacheSize = 100
      '  Set CLMBordx = cmd.Execute
    DataString = ""
        If CLMBordx.EOF Then
            MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            If Mid(CboCashType, 1, 1) = "R" Then
        
                txtLocalFileLocation = "C:\MedicaGenClaimASCII\"
                
                'txtFilName = Date & "_OUTPATIENTPROVIDER"
                'Open txtLocalFileLocation & txtFilName & ".txt" For Output As #1
                txtFilName = Format(Date, "yyyymmdd")
                txtFilName = txtFilName & "CIMBMEDIX - Reimbursement" & ".txt"
                Open txtLocalFileLocation & "\" & txtFilName For Output As #1
            
            
                Do Until CLMBordx.EOF
                        tmpcnt = tmpcnt + 1
                   
                        DataString = l(Format(CLMBordx!InvoiceReceiptDate, "dd-mmm-yyyy"), 11)
                        
                        DataString = DataString & l(CLMBordx!MBMEmployeeNo, 7)
                        
                        DataString = DataString & l(CLMBordx!PatientName, 50)
                        
                        DataString = DataString & l(CLMBordx!CLMTotalActual, 10)
                        
                        DataString = DataString & l(CLMBordx!CLMTotalActual, 10)
                        
                        tmpTotalAct = CDbl(tmpTotalAct) + CDbl(CLMBordx!CLMTotalActual)
                        
                        
                        If Trim(CLMBordx!InvoiceNo) <> "" Then
                            DataString = DataString & l(CLMBordx!InvoiceNo, 30)
                        Else
                            DataString = DataString & l(Mid(CLMBordx!CASNumber, 3, 9), 30)
                        End If
                        
                        DataString = DataString & l(CLMBordx!NPClinic, 50)
                        
                        DataString = DataString & l("", 10)
                        
                        DataString = DataString & l(Format(CLMBordx!DateVisit, "dd-mmm-yyyy"), 11)
                        
                        DataString = DataString & l("CLINIC(NON-PANEL)", 50)
                        
                        If Len(CLMBordx!ICDCode) Then
                            DataString = DataString & l(CLMBordx!ICDCode, 50)
                        Else
                            DataString = DataString & l("R05", 50)
                        End If
                        
                        DataString = DataString & l("SELF", 10)
                        
                        DataString = DataString & l(CLMBordx!CASNumber, 25)
                        
                        DataString = DataString & l(CLMBordx!Bordxrefno, 25)
                        
                        DataString = DataString & l(Format(CLMBordx!BordxDate, "dd-mmm-yyyy"), 11)
                        
                        DataString = DataString & l("GP", 5)
                        
                        DataString = DataString & l("5210303", 7)
                        
                        DataString = DataString & l(0, 10)
                        
                        DataString = DataString & l(0, 10)
                        
                        DataString = DataString & l("001692205056", 50)
                        
                        Print #1, DataString
                        
                        DataString = DataString & Chr(13)
                    CLMBordx.MoveNext
                    
                    Loop
            
            ElseIf Mid(CboCashType, 1, 1) = "C" Then
        
                txtLocalFileLocation = "C:\MedicaGenClaimASCII\"
                
                'txtFilName = Date & "_OUTPATIENTPROVIDER"
                'Open txtLocalFileLocation & txtFilName & ".txt" For Output As #1
                txtFilName = Format(Date, "yyyymmdd")
                txtFilName = txtFilName & "CIMBMEDIX - Cashless" & ".txt"
                Open txtLocalFileLocation & "\" & txtFilName For Output As #1
            
            
                Do Until CLMBordx.EOF
                        tmpcnt = tmpcnt + 1
                   
                        DataString = l(CLMBordx!ClinicCode, 8)
                        
                        DataString = DataString & l(CLMBordx!CLNName, 50)
                        
                        'If Len(CLMBordx!InvoiceNo) Then
                        '    DataString = DataString & l(CLMBordx!InvoiceNo, 15)
                        'Else
                        '    DataString = DataString & l("A012547", 15)
                        'End If
                        
                        If Trim(CLMBordx!InvoiceNo) <> "" Then
                            DataString = DataString & l(CLMBordx!InvoiceNo, 15)
                        Else
                            DataString = DataString & l(Mid(CLMBordx!CASNumber, 3, 9), 15)
                        End If
                        
                        If Len(CLMBordx!InvoiceDate) And CLMBordx!InvoiceDate <> "01/01/1900" Then
                            DataString = DataString & l(Format(CLMBordx!InvoiceDate, "dd-mmm-yyyy"), 11)
                        Else
                            DataString = DataString & l(Format(CLMBordx!BordxDate, "dd-mmm-yyyy"), 11)
                        End If
                        DataString = DataString & l(CLMBordx!MBMEmployeeNo, 7)
                        
                        DataString = DataString & l(CLMBordx!PatientName, 50)
                        
                        DataString = DataString & l(CLMBordx!CLMTotalActual, 20)
                        
                        tmpTotalAct = CDbl(tmpTotalAct) + CDbl(CLMBordx!CLMTotalActual)
                        
                        If Len(CLMBordx!DateVisit) Then
                            DataString = DataString & l(Format(CLMBordx!DateVisit, "dd-mmm-yyyy"), 11)
                        End If
                        DataString = DataString & l("OUTPATIENT", 50)
                        
                        If Len(CLMBordx!ICDCode) Then
                            DataString = DataString & l(CLMBordx!ICDCode, 50)
                        Else
                            DataString = DataString & l("R05", 50)
                        End If
                        
                        DataString = DataString & l("SELF", 10)
                        
                        DataString = DataString & l(CLMBordx!CASNumber, 25)
                        
                        DataString = DataString & l(CLMBordx!Bordxrefno, 25)
                        
                        DataString = DataString & l(Format(CLMBordx!BordxDate, "dd-mmm-yyyy"), 11)
                        
                        DataString = DataString & l("GP", 15)
                        
                        DataString = DataString & l("5210301", 7)
                        
                        DataString = DataString & l(0, 10)
                        
                        DataString = DataString & l(0, 10)
                        
                        DataString = DataString & l("001692205056", 50)
                        
                        
                        Print #1, DataString
                        
                        DataString = DataString & Chr(13)
                    CLMBordx.MoveNext
                    
                    Loop
            
                
            
                lblTotalAmt = tmpTotalApp
                txtTotalVisits = tmpcnt
            
            End If
            
            CLMBordx.Close
            Set CLMBordx = Nothing

                Trailer = Format(tmpTotalAct, "000000000000000000")
                'Debug.Print Trailer
                Print #1, Trailer
                Close #1
                'Print #2, Trailer
                'cmdGenerate.Enabled = False
            
            
                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                cmdGenerate.Enabled = True
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CLNFile()
Const METHOD_NAME = "CLNFile"
On Error GoTo debugs
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
    
    HOS.Open "select * from View_CLN_OMNI where RegisterDate >= '" & Trim(txtPaymentDateFr) & "' and RegisterDate <= '" & Trim(txtPaymentDateTo) & "'", MediConnCISDB, adOpenKeyset, adLockOptimistic
        
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
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "frmClaimASCIICIMB")
    CboCashType.AddItem "C" & " - " & "Cashless"
    CboCashType.AddItem "R" & " - " & "Reimbursement"
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

