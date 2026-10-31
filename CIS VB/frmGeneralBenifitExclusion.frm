VERSION 5.00
Begin VB.Form frmGeneralBenefitExclusion 
   Caption         =   "General Exclusions"
   ClientHeight    =   8856
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   12348
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8856
   ScaleWidth      =   12348
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   520
      Left            =   45
      TabIndex        =   32
      Top             =   240
      Width           =   12195
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   255
         Left            =   6550
         TabIndex        =   1
         Top             =   180
         Width           =   735
      End
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
         Width           =   5055
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
         TabIndex        =   33
         Top             =   200
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   45
      TabIndex        =   30
      Top             =   960
      Width           =   12195
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
         Left            =   7440
         TabIndex        =   26
         Top             =   960
         Width           =   2055
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
         Left            =   7440
         TabIndex        =   17
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox chkBloodTest 
         Caption         =   "Blood Test"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   2295
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
         Left            =   7440
         TabIndex        =   24
         Top             =   240
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
         Left            =   9840
         TabIndex        =   23
         Top             =   600
         Width           =   2055
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
         TabIndex        =   21
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
         Left            =   9840
         TabIndex        =   22
         Top             =   240
         Width           =   1995
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
         TabIndex        =   20
         Top             =   1320
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
         TabIndex        =   8
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox chkPregnacyTest 
         Caption         =   "Prenancy Related/Test"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   1320
         Width           =   2295
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
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   600
         Width           =   1935
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
         TabIndex        =   12
         Top             =   1680
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
         Left            =   7440
         TabIndex        =   10
         Top             =   1680
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
         Left            =   9840
         TabIndex        =   11
         Top             =   1320
         Width           =   2055
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
         Left            =   7440
         TabIndex        =   6
         Top             =   600
         Width           =   2055
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
         TabIndex        =   7
         Top             =   960
         Width           =   2295
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
         TabIndex        =   5
         Top             =   240
         Width           =   1455
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
         TabIndex        =   3
         Top             =   240
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
         TabIndex        =   4
         Top             =   1680
         Width           =   2055
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
         Left            =   9840
         TabIndex        =   2
         Top             =   960
         Width           =   2055
      End
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
         TabIndex        =   9
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Height          =   4215
      Left            =   45
      TabIndex        =   35
      Top             =   3840
      Width           =   12195
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
         Height          =   3660
         Left            =   120
         MaxLength       =   15000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   240
         Width           =   11895
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
      TabIndex        =   31
      Top             =   8160
      Width           =   12195
      Begin VB.CommandButton cmdreset 
         Caption         =   "Clear"
         Height          =   375
         Left            =   9480
         TabIndex        =   37
         Top             =   170
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10320
         TabIndex        =   27
         Top             =   170
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   11160
         TabIndex        =   28
         Top             =   170
         Width           =   855
      End
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
      TabIndex        =   36
      Top             =   780
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Remarks"
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
      TabIndex        =   34
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "General Exclusions"
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
      TabIndex        =   29
      Top             =   0
      Width           =   12255
   End
End
Attribute VB_Name = "frmGeneralBenefitExclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub cmdreset_Click()
        cmdSearch.Enabled = True
        cboGrpCompany.Enabled = True
        cboGrpCompany = ""
        Call checkbox("0")
        cmdUpdate.Enabled = False
        cboGrpCompany.SetFocus
End Sub

Public Sub cmdSearch_Click()
    Call checkbox("0")
    Call CallExclusions
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
    Call ActionLogUpdate(12, "frmGeneralBenefitExclusion")
    Call ComboList
End Sub

Private Sub ComboList()
Const METHOD_NAME = "ComboList"
On Error GoTo debugs
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset
    
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        
        cboGrpCompany.Clear
        'strsql = "select GRPCoName from GRPCompany where GRPCoPAYCode= 'DC'"
        '08-07-2016 Request by Ms.Foo to call out all group companies.
        strsql = "select GRPCoName from GRPCompany "
        PAYGRP.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If PAYGRP.RecordCount Then
            Do Until PAYGRP.EOF
               cboGrpCompany.AddItem Trim(PAYGRP!GrpCoName)
               PAYGRP.MoveNext
            Loop
        End If
        PAYGRP.Close
        Set PAYGRP = Nothing
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
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


  '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    
    If cboGrpCompany = "" Then
    
        MsgBox "Please Select Group Company!", vbOKOnly + vbCritical
    Else
        cboGrpCompany1 = ""
        cboGrpCompany1 = Trim(Replace(cboGrpCompany, "'", "''"))
        strsql = "Select * from GeneralExclusions where GRPCompany = '" & cboGrpCompany1 & "'"
        rst.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If rst.RecordCount > 0 Then
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
                cmdSearch.Enabled = False
                cboGrpCompany.Enabled = False
                cmdUpdate.Enabled = True
            End With
        Else
            MsgBox "No Record Found!", vbOKOnly + vbCritical
            cmdUpdate.Enabled = True
        End If
        rst.Close
        Set rst = Nothing
    End If
   ' MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing
    
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
Dim rst As New ADODB.Recordset
Dim cboGrpCompany1 As String
    Screen.MousePointer = vbHourglass
    
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    
    If cboGrpCompany = "" Then
        MsgBox "Please Select Group Company!", vbOKOnly + vbCritical
    Else
        cboGrpCompany1 = ""
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
                !LastUpdateUser = ChkStr(Logon)
                !LastUpdatedDate = Date & " " & Time
            End With
            rst.update
            MsgBox "Record Saved ! ", vbOKOnly
    End If
    
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
    Screen.MousePointer = vbDefault
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub txtCASRemark_Click()
    KeyPreview = False
End Sub

Private Sub txtCASRemark_LostFocus()
    KeyPreview = True
End Sub
