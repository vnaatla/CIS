VERSION 5.00
Begin VB.Form frmClaimsDataOMNI 
   Caption         =   "Claims Data - OMNI"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   80
      TabIndex        =   11
      Top             =   1800
      Width           =   7095
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
         TabIndex        =   13
         Top             =   240
         Width           =   705
      End
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
         TabIndex        =   12
         Top             =   240
         Width           =   720
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
         TabIndex        =   17
         Top             =   270
         Width           =   1200
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
         TabIndex        =   16
         Top             =   240
         Width           =   1920
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
         TabIndex        =   15
         Top             =   270
         Width           =   1080
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
         TabIndex        =   14
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
         TabIndex        =   2
         Top             =   1170
         Width           =   1455
      End
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
         TabIndex        =   1
         Top             =   1920
         Visible         =   0   'False
         Width           =   3255
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
         Top             =   1180
         Width           =   915
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
         TabIndex        =   9
         Top             =   900
         Width           =   1245
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
         TabIndex        =   8
         Top             =   615
         Width           =   765
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
         TabIndex        =   7
         Top             =   240
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
         TabIndex        =   6
         Top             =   315
         Width           =   1125
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
         TabIndex        =   5
         Top             =   540
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
         TabIndex        =   4
         Top             =   1920
         Visible         =   0   'False
         Width           =   1155
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
         TabIndex        =   3
         Top             =   840
         Width           =   3225
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
Attribute VB_Name = "frmClaimsDataOMNI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub CmdGenerate_Click()
    Dim HOS As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    Dim tmpcnt As Double
    Dim Trailer As String
    Dim tmpfilename As String
    Dim tmpLocalFileLocation As String
    Dim DataString As String
    Dim Header As String
    
    tmpcnt = 0
    
    Dim cmd As New ADODB.Command
    Dim Prm1 As New ADODB.Parameter
    Dim Prm2 As New ADODB.Parameter
    Dim total As Integer

    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
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
            txtFilName = txtFilName & "_OUTPATIENTPROVIDER" & ".txt"
            Open txtLocalFileLocation & "\" & txtFilName For Output As #1
            
            Header = "01" & "|" & "TPA" & "|" & Date
                'Debug.Print Trailer
            Print #1, Header
            
        Do Until HOS.EOF
             
                tmpcnt = tmpcnt + 1
           ' DataString = ""
                DataString = "02" & "|"
                
                DataString = DataString & "N" & "|"
                
                If CLMBordx!CLNType = "G" Then
                    DataString = DataString & "OPGP" & "|"
                Else
                    DataString = DataString & "OPSP" & "|"
                End If
                
                DataString = DataString & CLMBordx!CasNumber & "|"
                
                DataString = DataString & "" & "|"
                
                DataString = DataString & "" & "|"
                
                DataString = DataString & Format(CLMBordx!DateVisit, "DD/MM/YYYY") & "|"
                
                DataString = DataString & Format(CLMBordx!BordxDate, "DD/MM/YYYY") & "|"
                
                If CLMBordx!INVPayTo = "M" Then
                    DataString = DataString & "2" & "|"
                Else
                    DataString = DataString & "1" & "|"
                End If
                
                DataString = DataString & CLMBordx!ClinicCode & "|"
                
                DataString = DataString & CLMBordx!CLNName & "|"
                
                DataString = DataString & "ZZZ999" & "|"
                
                If ReasonDesc5 <> "" Then
                    DataString = DataString & CLMBordx!ReasonDesc5 & "|"
                Else
                    DataString = DataString & CLMBordx!ReasonDesc5 & "|"
                End If
                
                Print #1, DataString
                
                DataString = DataString & Chr(13)
            
            CLMBordx.MoveNext
            
            Loop

                Trailer = "3" & "|" & tmpcnt
                'Debug.Print Trailer
                Print #1, Trailer
                'Print #2, Trailer
                'cmdGenerate.Enabled = False
                Close #1
                
                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                cmdGenerate.Enabled = True
        End If
    HOS.Close
    Set HOS = Nothing

End Sub

