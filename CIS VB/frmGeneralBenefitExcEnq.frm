VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmGeneralBenefitExcEnq 
   Caption         =   "General Exclusions (Enquiry)"
   ClientHeight    =   9024
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   9528
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9024
   ScaleWidth      =   9528
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   45
      TabIndex        =   9
      Top             =   960
      Width           =   9075
      Begin VB.CheckBox chkPapSmear 
         Caption         =   "Pap-Smear"
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
         Left            =   4920
         TabIndex        =   33
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox chkDermatitis 
         Caption         =   "Dermatitis/Skin Disorder"
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
         Left            =   240
         TabIndex        =   32
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CheckBox chkDental 
         Caption         =   "Dental"
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
         Left            =   240
         TabIndex        =   31
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CheckBox chkGynaecological 
         Caption         =   "Gynaecological Disorder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   30
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkOptical 
         Caption         =   "Optical"
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
         Left            =   4920
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkPreEmpCheckup 
         Caption         =   "Pre-Employment Check-up"
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
         Left            =   4920
         TabIndex        =   28
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox chkMedCheckup 
         Caption         =   "Medical Check-up"
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
         Left            =   2520
         TabIndex        =   27
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox chkIHDCAD 
         Caption         =   "IHD/CAD"
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
         Left            =   2520
         TabIndex        =   26
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CheckBox chkDiabetes 
         Caption         =   "Diabetes"
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
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CheckBox chkThyroid 
         Caption         =   "Tyroid Disorder"
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
         Left            =   4920
         TabIndex        =   24
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CheckBox chkHRT 
         Caption         =   "HRT"
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
         Left            =   2520
         TabIndex        =   23
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox chkImmunisation 
         Caption         =   "Immunisation"
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
         Left            =   2520
         TabIndex        =   22
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CheckBox chkPregnacyTest 
         Caption         =   "Prenancy Related/Test"
         Height          =   255
         Left            =   4920
         TabIndex        =   21
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CheckBox chkHypertension 
         Caption         =   "Hypertension"
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
         Left            =   2520
         TabIndex        =   20
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox chkAbortion 
         Caption         =   "Abortion/Miscarriage"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkAntenatal 
         Caption         =   "Antenatal/Postnatal"
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
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox chkInfertility 
         Caption         =   "Infertility"
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
         Left            =   2520
         TabIndex        =   17
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox chkGenitourinary 
         Caption         =   "Genitourinary Disorder"
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
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   1995
      End
      Begin VB.CheckBox chkContraception 
         Caption         =   "Contraception"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox chkVitamins 
         Caption         =   "Vitamin purchase"
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
         Left            =   4920
         TabIndex        =   14
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CheckBox chkXRay 
         Caption         =   "X-Ray/MRI"
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
         Left            =   4920
         TabIndex        =   13
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox chkBloodTest 
         Caption         =   "Blood Test"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox chkTyphoid 
         Caption         =   "Typhoid Vaccine"
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
         Left            =   4920
         TabIndex        =   11
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CheckBox chkLTMedication 
         Caption         =   "Long-Term Medication"
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
         Left            =   2520
         TabIndex        =   10
         Top             =   2400
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   45
      TabIndex        =   6
      Top             =   8280
      Width           =   9075
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   8130
         TabIndex        =   3
         Top             =   170
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   7275
         TabIndex        =   2
         Top             =   170
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   520
      Left            =   45
      TabIndex        =   4
      Top             =   240
      Width           =   9075
      Begin VB.ComboBox cboGrpCompany 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   150
         Width           =   6615
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   255
         Left            =   8160
         TabIndex        =   1
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Group Company"
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
         TabIndex        =   5
         Top             =   200
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   34
      Top             =   4200
      Width           =   8895
      _ExtentX        =   15685
      _ExtentY        =   7218
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Remarks"
      TabPicture(0)   =   "frmGeneralBenefitExcEnq.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtRemarks"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Benifit Schedule"
      TabPicture(1)   =   "frmGeneralBenefitExcEnq.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lstPolicyWording"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.ListBox lstPolicyWording 
         Height          =   3120
         ItemData        =   "frmGeneralBenefitExcEnq.frx":0038
         Left            =   120
         List            =   "frmGeneralBenefitExcEnq.frx":003A
         TabIndex        =   36
         ToolTipText     =   "right click to randomize contents"
         Top             =   480
         Width           =   8655
      End
      Begin VB.TextBox txtRemarks 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   -74880
         MaxLength       =   15000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   480
         Width           =   8655
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "General Exclusions (Enquiry)"
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
      TabIndex        =   8
      Top             =   0
      Width           =   9255
   End
   Begin VB.Label Label1 
      Caption         =   "General Exclusions"
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
      Top             =   780
      Width           =   1335
   End
End
Attribute VB_Name = "frmGeneralBenefitExcEnq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpPolicyWording As String
Dim PictureFileName As String
Private Const SW_RESTORE        As Long = &H9&
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub checkbox(status As Integer)
    chkDermatitis.Value = status
    chkGynaecological.Value = status
    chkDental.Value = status
    chkOptical.Value = status
    chkMedCheckup.Value = status
    chkPreEmpCheckup.Value = status
    chkHypertension.Value = status
    chkDiabetes.Value = status
    chkIHDCAD.Value = status
    chkThyroid.Value = status
    chkImmunisation.Value = status
    chkHRT.Value = status
    chkPregnacyTest.Value = status
    chkAntenatal.Value = status
    chkAbortion.Value = status
    chkInfertility.Value = status
    chkContraception.Value = status
    chkGenitourinary.Value = status
    chkVitamins.Value = status
    chkBloodTest.Value = status
    chkXRay.Value = status
    chkPapSmear.Value = status
    chkTyphoid.Value = status
    chkLTMedication.Value = status
    txtRemarks = ""
End Sub

Private Sub lstPolicyWording_Click()
Const METHOD_NAME = "lstPolicyWording_Click"
On Error GoTo debugs

Dim lRet    As Long
Dim sFile   As String
    
    PictureFileName = SRVPath & "CLNPolicyWording\" & tmpPolicyWording & ""
    crt PictureFileName & "\" & lstPolicyWording.List(lstPolicyWording.ListIndex)
    sFile = PictureFileName & "\" & lstPolicyWording.List(lstPolicyWording.ListIndex)
    lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
debugs:
If Err.Description <> "" Then
   Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub





Sub crt(pth As String)
Dim f
f = FreeFile
Open App.path + "\temp.html" For Output As #f
    Print #f, "<html><body><center><img src=""" & pth & """></center></body></html>"
Close #f
DoEvents
'WebBrowser1.Navigate App.Path & "\temp.html"
End Sub



Sub addTheFiles()
Dim B As String
lstPolicyWording.Clear

    'PictureFileName = SRVPath & "CLNPolicyWording\" & tmpPolicyWording & ""
    PictureFileName = SRVPath & "CLNPolicyWording\" & cboGrpCompany & ""
    '\Medical Rpt"
    B = Dir(PictureFileName & "\*.*")
    If B <> "" Then
        If Right(B, 4) = ".jpg" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".tif" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".gif" Then
            lstPolicyWording.AddItem B
        End If
        
        If Right(B, 4) = ".JPG" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".TIF" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".GIF" Then
            lstPolicyWording.AddItem B
        End If
        
        If Right(B, 4) = ".pdf" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".PDF" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".txt" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".TXT" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = "XLSX" Then
                lstPolicyWording.AddItem B
        End If
         If Right(B, 4) = "xlsx" Then
                lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = "docx" Then
             lstPolicyWording.AddItem B
        End If
        
        If Right(B, 4) = "DOCX" Then
            lstPolicyWording.AddItem B
        End If
        
    Else
        Exit Sub
    End If
    Do: DoEvents
        B = Dir
        If B <> "" Then
            If Right(B, 4) = ".jpg" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = ".tif" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = ".gif" Then
                lstPolicyWording.AddItem B
            End If
            
            If Right(B, 4) = ".JPG" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = ".TIF" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = ".GIF" Then
                lstPolicyWording.AddItem B
            End If
            
            If Right(B, 4) = ".PDF" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = ".pdf" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = ".txt" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = ".TXT" Then
                lstPolicyWording.AddItem B
            End If
            
            If Right(B, 4) = "xlsx" Then
                lstPolicyWording.AddItem B
            End If
            If Right(B, 4) = "docx" Then
                lstPolicyWording.AddItem B
            End If
            
            If Right(B, 4) = "DOCX" Then
                lstPolicyWording.AddItem B
            End If
            
            If Right(B, 4) = "XLSX" Then
                lstPolicyWording.AddItem B
            End If
            
        Else
            Exit Do
        End If
    Loop
End Sub


Private Sub cboGrpCompany_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboGrpCompany.Text, cboGrpCompany.SelStart) + Chr(KeyAscii) + Mid(cboGrpCompany.Text, cboGrpCompany.SelStart + cboGrpCompany.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboGrpCompany.ListCount - 1
 
        If NewText = LCase(Left(cboGrpCompany.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboGrpCompany.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboGrpCompany.Text = ValidValue
                cboGrpCompany.SelStart = 0
                cboGrpCompany.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Public Sub cmdSearch_Click()
    Call checkbox("0")
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        SSTab1.Tab = 0
        Call CallExclusions
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
        
End Sub

Private Sub cmdUpdate_Click()
    Call UpdateExclusions
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
    Call ActionLogUpdate(12, "frmGeneralBenefitExcEnq")
    Call ComboList
    SSTab1.Tab = 0
End Sub

Private Sub ComboList()
Const METHOD_NAME = "ComboList"
On Error GoTo debugs
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        cboGrpCompany.Clear
        'strsql = "select GRPCoName from GRPCompany where GRPCoPAYCode= 'DC'"
        '08-07-2016 Request by Ms.Foo to call out all group companies.
        strsql = "select GRPCoName from GRPCompany "
        
            PAYGRP.Open strsql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            
            Do While Not PAYGRP.EOF
               cboGrpCompany.AddItem Trim(PAYGRP!GrpCoName)
               PAYGRP.MoveNext
            Loop
        PAYGRP.Close
        Set PAYGRP = Nothing
        
     '   MediConnCISDB.Close
     '   Set MediConnCISDB = Nothing
debugs:
If Err.Description <> "" Then
   Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CallExclusions()
Const METHOD_NAME = "CallExclusions"
On Error GoTo debugs
Dim strsql As String
Dim cboGrpCompany1 As String
Dim rst As New ADODB.Recordset
    Screen.MousePointer = vbHourglass

    If cboGrpCompany = "" Then
    
        MsgBox "Please Select Group Company!", vbOKOnly + vbCritical
    Else
        cboGrpCompany1 = ""
        cboGrpCompany1 = Trim(Replace(cboGrpCompany, "'", "''"))
        strsql = "Select * from GeneralExclusions where GRPCompany = '" & cboGrpCompany1 & "'"
        rst.Open strsql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        
        Do While Not rst.EOF
            With rst
                chkDermatitis.Value = IIf(Trim(!Dermatitis) = "Y", 1, 0)
                chkGynaecological.Value = IIf(Trim(!Gynaecological) = "Y", 1, 0)
                chkDental.Value = IIf(Trim(!Dental) = "Y", 1, 0)
                chkOptical.Value = IIf(Trim(!Optical) = "Y", 1, 0)
                chkMedCheckup.Value = IIf(Trim(!MedCheckup) = "Y", 1, 0)
                chkPreEmpCheckup.Value = IIf(Trim(!PreEmpCheckup) = "Y", 1, 0)
                chkHypertension.Value = IIf(Trim(!Hypertension) = "Y", 1, 0)
                chkDiabetes.Value = IIf(Trim(!Diabetes) = "Y", 1, 0)
                chkIHDCAD.Value = IIf(Trim(!IHDCAD) = "Y", 1, 0)
                chkThyroid.Value = IIf(Trim(!thyroid) = "Y", 1, 0)
                chkImmunisation.Value = IIf(Trim(!Immunisation) = "Y", 1, 0)
                chkHRT.Value = IIf(Trim(!HRT) = "Y", 1, 0)
                chkPregnacyTest.Value = IIf(Trim(!PregnacyTest) = "Y", 1, 0)
                chkAntenatal.Value = IIf(Trim(!Antenatal) = "Y", 1, 0)
                chkAbortion.Value = IIf(Trim(!Abortion) = "Y", 1, 0)
                chkInfertility.Value = IIf(Trim(!Infertility) = "Y", 1, 0)
                chkContraception.Value = IIf(Trim(!Contraception) = "Y", 1, 0)
                chkGenitourinary.Value = IIf(Trim(!Genitourinary) = "Y", 1, 0)
                chkVitamins.Value = IIf(Trim(!Vitamins) = "Y", 1, 0)
                chkBloodTest.Value = IIf(Trim(!BloodTest) = "Y", 1, 0)
                chkXRay.Value = IIf(Trim(!XRay) = "Y", 1, 0)
                chkPapSmear.Value = IIf(Trim(!PapSmear) = "Y", 1, 0)
                chkTyphoid.Value = IIf(Trim(!Typhoid) = "Y", 1, 0)
                chkLTMedication.Value = IIf(Trim(!LTMedication) = "Y", 1, 0)
                txtRemarks = IIf(Trim(Len(!Remarks)), !Remarks, "")
                tmpPolicyWording = IIf(Len(cboGrpCompany1), cboGrpCompany1, "")
            End With
        rst.MoveNext
        Loop
        rst.Close
        Set rst = Nothing
    End If
    Screen.MousePointer = vbDefault
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub UpdateExclusions()
Const METHOD_NAME = "UpdateExclusions"
On Error GoTo debugs
Dim strsql As String
Dim cboGrpCompany1  As String
Dim rst As New ADODB.Recordset

    Screen.MousePointer = vbHourglass

    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    
    If cboGrpCompany = "" Then
        MsgBox "Please Select Group Company!", vbOKOnly + vbCritical
    Else
        cboGrpCompany1 = Trim(Replace(cboGrpCompany, "'", "''"))
        strsql = "Select * from GeneralExclusions where GRPCompany = '" & cboGrpCompany1 & "'"
        rst.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
            With rst
                !Dermatitis = IIf(chkDermatitis.Value = 1, "Y", "N")
                !Gynaecological = IIf(chkGynaecological.Value = 1, "Y", "N")
                !Dental = IIf(chkDental.Value = 1, "Y", "N")
                !Optical = IIf(chkOptical.Value = 1, "Y", "N")
                !MedCheckup = IIf(chkMedCheckup.Value = 1, "Y", "N")
                !PreEmpCheckup = IIf(chkPreEmpCheckup.Value = 1, "Y", "N")
                !Hypertension = IIf(chkHypertension.Value = 1, "Y", "N")
                !Diabetes = IIf(chkDiabetes.Value = 1, "Y", "N")
                !IHDCAD = IIf(chkIHDCAD.Value = 1, "Y", "N")
                !thyroid = IIf(chkThyroid.Value = 1, "Y", "N")
                !Immunisation = IIf(chkImmunisation.Value = 1, "Y", "N")
                !HRT = IIf(chkHRT.Value = 1, "Y", "N")
                !PregnacyTest = IIf(chkPregnacyTest.Value = 1, "Y", "N")
                !Antenatal = IIf(chkAntenatal.Value = 1, "Y", "N")
                !Abortion = IIf(chkAbortion.Value = 1, "Y", "N")
                !Infertility = IIf(chkInfertility.Value = 1, "Y", "N")
                !Contraception = IIf(chkContraception.Value = 1, "Y", "N")
                !Genitourinary = IIf(chkGenitourinary.Value = 1, "Y", "N")
                !Vitamins = IIf(chkVitamins.Value = 1, "Y", "N")
                !BloodTest = IIf(chkBloodTest.Value = 1, "Y", "N")
                !XRay = IIf(chkXRay.Value = 1, "Y", "N")
                !PapSmear = IIf(chkPapSmear.Value = 1, "Y", "N")
                !Typhoid = IIf(chkTyphoid.Value = 1, "Y", "N")
                !LTMedication = IIf(chkLTMedication.Value = 1, "Y", "N")
                !GRPCompany = cboGrpCompany
                !Remarks = IIf(Len(txtRemarks), UCase(txtRemarks), "")
            End With
            rst.update
            MsgBox "Record Saved ! ", vbOKOnly
    End If
    Screen.MousePointer = vbDefault
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
 If SSTab1.Tab = 1 Then
    If Trim(tmpPolicyWording) <> "" Then
        Call LoadDoc
    End If
 End If
End Sub


Private Sub LoadDoc()
On Error Resume Next
Dim B As String, f As Integer, d As Integer, e As Integer
    Call addTheFiles
Exit Sub
End Sub

