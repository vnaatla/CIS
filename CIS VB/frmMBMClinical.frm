VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMBMClinical 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Clinic"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTempMCOFees 
      Height          =   285
      Left            =   0
      MaxLength       =   6
      TabIndex        =   27
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtTempPrem 
      Height          =   315
      Left            =   480
      MaxLength       =   6
      TabIndex        =   26
      Top             =   3720
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtRemarks 
         Height          =   1845
         Left            =   2800
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   360
         Width           =   3450
      End
      Begin VB.ComboBox cboPLNCode 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   200
         Width           =   1500
      End
      Begin VB.ComboBox cboINSCode 
         Height          =   315
         ItemData        =   "frmMBMClinical.frx":0000
         Left            =   1200
         List            =   "frmMBMClinical.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   460
         Width           =   1500
      End
      Begin MSMask.MaskEdBox txtMBMBirthdate 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   720
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMBMPayorEffDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   960
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMBMPayorExpDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   1245
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboMBMRelCode 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Text            =   "C - Combined Limit"
         Top             =   1530
         Width           =   1545
      End
      Begin VB.Label Label4 
         Caption         =   "Supp Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1590
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "DOB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   770
         Width           =   1035
      End
      Begin VB.Label Label40 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   2790
         TabIndex        =   21
         Top             =   120
         Width           =   1590
      End
      Begin VB.Label Label21 
         Caption         =   "Plan Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label16 
         Caption         =   "Exp Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label Label20 
         Caption         =   "Ins Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   960
      End
      Begin VB.Label Label15 
         Caption         =   "Eff Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1050
         Width           =   1275
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   6375
      Begin VB.Label Label3 
         Caption         =   "Visit Limit"
         Height          =   255
         Left            =   3360
         TabIndex        =   29
         Top             =   495
         Width           =   840
      End
      Begin VB.Label LblVisitLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   28
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label lblMBMPremium 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label lblMCOFees 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label lblAnnualLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   15
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Label27 
         Caption         =   "Ann Limit"
         Height          =   255
         Left            =   3360
         TabIndex        =   12
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label26 
         Caption         =   "MCO Fees"
         Height          =   240
         Left            =   480
         TabIndex        =   11
         Top             =   500
         Width           =   870
      End
      Begin VB.Label Label25 
         Caption         =   "Premium"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   2880
      Width           =   6375
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   4320
         TabIndex        =   8
         Top             =   160
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtTempEffDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTempExpDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin VB.Label txtAGECode 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1080
      TabIndex        =   30
      Top             =   3120
      Width           =   1635
   End
End
Attribute VB_Name = "frmMBMClinical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Source As Integer
Public Access As String

Private Sub cboPLNCode_Click()
    If Source <> 3 Then
        If Len(cboPLNCode) >= 2 Then
            If txtMBMPayorEffDate <> "__/__/____" Then
                Call CallCLNLimit
                Call MemberCLNCallFees
                Call VisitLimit
            End If
        End If
    End If
End Sub

Private Sub cboPLNCode_Change()
    If Source <> 3 Then
        If Len(cboPLNCode) >= 2 Then
            If txtMBMPayorEffDate <> "__/__/____" Then
                Call CallCLNLimit
                Call MemberCLNCallFees
                Call VisitLimit
            End If
        End If
    End If
End Sub

Private Sub ComboListing()
    Dim PLN As New ADODB.Recordset
    Dim INS As New ADODB.Recordset
    
    PLN.Open "select PLNCode, PLNDesc from PLN", MediConnCISDB, adOpenKeyset, adLockOptimistic
    INS.Open "select INSCode, INSDesc from INS", MediConnCISDB, adOpenKeyset, adLockOptimistic

    Do Until PLN.EOF
        cboPLNCode.AddItem PLN!PLNCode & " - " & PLN!PLNDesc
        PLN.MoveNext
    Loop
    
    Do Until INS.EOF
        cboINSCode.AddItem INS!INSCode & " - " & INS!INSDesc
        INS.MoveNext
    Loop
    
    cboMBMRelCode.AddItem "C - Combined Limit"
    cboMBMRelCode.AddItem "S - Separated Limit"
    
    INS.Close
    Set INS = Nothing
    PLN.Close
    Set PLN = Nothing
End Sub

Private Sub cboINSCode_Change()
    If Source <> 3 Then
        If Len(cboPLNCode) >= 2 Then
            Call CallCLNLimit
            Call MemberCLNCallFees
            Call VisitLimit
        End If
    End If
End Sub

Private Sub cboINSCode_Click()
    
    If Source <> 3 Then
        If Len(cboPLNCode) >= 2 Then
            Call CallCLNLimit
            Call MemberCLNCallFees
            Call VisitLimit
        End If
    End If
End Sub

Private Sub cboINSCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    'Clipboard.clear
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboINSCode.Text, cboINSCode.SelStart) + Chr(KeyAscii) + Mid(cboINSCode.Text, cboINSCode.SelStart + cboINSCode.SelLength + 1))
  
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboINSCode.ListCount - 1
        
        If NewText = LCase(Left(cboINSCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINSCode.List(i)
        
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             
             If ValidCount = 1 Then
               cboINSCode.Text = ValidValue
               cboINSCode.SelStart = 0
               cboINSCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPLNCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    'Clipboard.clear
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPLNCode.Text, cboPLNCode.SelStart) + Chr(KeyAscii) + Mid(cboPLNCode.Text, cboPLNCode.SelStart + cboPLNCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPLNCode.ListCount - 1

        If NewText = LCase(Left(cboPLNCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPLNCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

           If ValidCount = 1 Then
             cboPLNCode.Text = ValidValue
             cboPLNCode.SelStart = 0
             cboPLNCode.SelLength = Len(ValidValue)
            End If
 
        End If
End Sub

Public Sub cmdClose_Click()
Dim RecordExit
    
    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        bExit = True
       Unload Me
    End If
End Sub

Private Sub cmdOK_Click()
    
    If Source = 3 Then
    
    Else
    
        If Trim(cboPLNCode) = "" Then
            MsgBox "Please Enter Plan Code!", vbCritical
            cboPLNCode.SetFocus
        ElseIf Trim(cboINSCode) = "" Then
            MsgBox "Please Enter Insured Type!", vbCritical
            cboINSCode.SetFocus
        ElseIf txtMBMPayorEffDate = "__/__/____" Then
            MsgBox "Please Enter Effective Date!", vbCritical
            txtMBMPayorEffDate.SetFocus
        ElseIf txtMBMPayorExpDate = "__/__/____" Then
            MsgBox "Please Enter ExpiryDate!", vbCritical
            txtMBMPayorExpDate.SetFocus
        End If
        'CLNPLNCode = Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") = 0, 4, InStr(1, cboPLNCode, "-") - 1)))
        'CLNINSType = Mid(cboINSCode, 1, 1)
        'CLNPAYEffDate = Format(txtMBMPayorEffDate, "dd/mm/yyyy")
        'CLNPAYExpDate = Format(txtMBMPayorExpDate, "dd/mm/yyyy")
        'CLMTempAnnLimit = CDbl(lblAnnualLimit)
        'CLMTempPrem = CDbl(lblMBMPremium)
        'CLNTempMCOFee = CDbl(lblMCOFees)
        'CLNRemark = ChkStr(Trim(txtRemarks))
        frmMBMClinical.Visible = False
    End If
    frmMBMClinical.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '--- Enable PrintScreen feature
    '--- IMPORTANT - Don't forget to set .KeyPreview = True for this form at design time!!
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If

End Sub

Private Sub Form_Load()
    
    If Source = 3 Then
    
        Call showMBMClinical
    
    Else
    
        Call ComboListing
    
    End If
End Sub

Private Sub txtMBMPayorEffDate_LostFocus()
Dim MBMValidDate As New ADODB.Recordset
    Dim strsql As String
    
    Dim tempdate As Date
    Dim answer

    If (txtMBMPayorEffDate) <> "__/__/____" Then
        If IsDate(txtMBMPayorEffDate) Then
            If Mid(frmNewMemberRegistration.cboHNC, 1, 1) = "C" Then
                If ChkStr(frmNewMemberRegistration.cboMBMRenewal) = "N" Then
                    frmNewMemberRegistration.txtMBMDateJoin = txtMBMPayorEffDate
                Else
                    If frmNewMemberRegistration.txtMBMDateJoin = "__/__/____" Then
                        frmNewMemberRegistration.txtMBMDateJoin = txtMBMPayorEffDate
                    End If
                End If
            End If
           strsql = "Select ValidID , MBMEffStart, MBMEffEnd From MBMValidDateRange "
           MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                           
           If MBMValidDate.recordCount Then
                If DateDiff("d", MBMValidDate!MBMEffStart, txtMBMPayorEffDate) >= 0 And DateDiff("d", MBMValidDate!MBMEffEnd, txtMBMPayorEffDate) <= 0 Then
                    tempdate = Format(DateAdd("yyyy", 1, txtMBMPayorEffDate), "dd/mm/yyyy")
                    txtMBMPayorExpDate = Format(DateAdd("d", -1, tempdate), "dd/mm/yyyy")
                
                If txtMBMBirthdate <> "__/__/____" Then
                    Call CLNMemberCallAge
                    Call MemberCLNCallFees
                    Call VisitLimit
                    Call CallCLNLimit
                End If
                    
                Else
                    MsgBox "Policy Effective Date is out of acceptable range! " & vbNewLine & _
                            "It must be in between " & MBMValidDate!MBMEffStart & " and " & MBMValidDate!MBMEffEnd & "! ", vbCritical + vbOKOnly
                    txtMBMPayorEffDate.SetFocus
                End If
           Else
                tempdate = Format(DateAdd("yyyy", 1, txtMBMPayorEffDate), "dd/mm/yyyy")
                txtMBMPayorExpDate = Format(DateAdd("d", -1, tempdate), "dd/mm/yyyy")
           
                If txtMBMBirthdate <> "__/__/____" Then
                    Call CLNMemberCallAge
                    Call MemberCLNCallFees
                    Call VisitLimit
                    Call CallCLNLimit
                End If
           
           End If
           MBMValidDate.Close
           Set MBMValidDate = Nothing
        Else
            answer = MsgBox("Invalid Date Format!", vbCritical + vbOKCancel)
            If answer = vbOK Then
                txtMBMPayorEffDate.SetFocus
            Else
                txtMBMPayorExpDate.SetFocus
            End If
        End If
    End If
    txtTempEffDate = txtMBMPayorEffDate
    txtTempExpDate = txtMBMPayorExpDate
End Sub

Private Sub txtMBMPayorExpDate_LostFocus()
Dim MBMValidDate As New ADODB.Recordset
Dim strsql As String
Dim answer
    
    If txtMBMPayorExpDate <> "__/__/____" Then
        If IsDate(txtMBMPayorExpDate) = True Then
            strsql = "Select ValidID , MBMExpStart, MBMExpEnd From MBMValidDateRange "
            MBMValidDate.MaxRecords = 1
            MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                 
            If MBMValidDate.recordCount Then
                If DateDiff("d", MBMValidDate!MBMExpStart, txtMBMPayorExpDate) >= 0 And DateDiff("d", MBMValidDate!MBMExpEnd, txtMBMPayorExpDate) <= 0 Then
                Else
                    MsgBox "Policy Expiry Date is out of acceptable range! " & vbNewLine & _
                            " It must be in between " & MBMValidDate!MBMExpStart & " and " & MBMValidDate!MBMExpEnd, vbCritical + vbOKOnly
                    txtMBMPayorExpDate.SetFocus
                End If
            End If
            MBMValidDate.Close
            Set MBMValidDate = Nothing
        
            If txtMBMPayorEffDate <> "__/__/____" And txtTempEffDate <> "__/__/____" And txtTempExpDate <> "__/__/____" And txtMBMPayorExpDate <> "__/__/____" Then
                If IsDate(txtMBMPayorEffDate) = True And IsDate(txtMBMPayorExpDate) Then
                 Call InsertCLNMCO
                 Call InsertCLNPrem
                 'Call InsertProviderFees
                
                 If CDate(txtMBMPayorEffDate) = CDate(txtTempEffDate) And CDate(txtTempExpDate) = CDate(txtMBMPayorExpDate) Then
                     lblMCOFees = txtTempMCOFees
                     lblMBMPremium = txtTempPrem
                     'txtMBMProviderCharges = txtTempProvider
                 End If
                End If
            End If
        Else
            MsgBox "Invalid Date Format!", vbCritical + vbOKOnly
'            If answer = vbOK Then
             txtMBMPayorExpDate.SetFocus
 '           End If
        End If
    End If
End Sub

Public Sub CallCLNLimit()
Dim ChkAnnLimit As New ADODB.Recordset
Dim strsql As String

    lblAnnualLimit = ""
    
    If Len(Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") = 0, 4, InStr(1, cboPLNCode, "-") - 1)))) And IsDate(txtMBMPayorEffDate) Then
        strsql = " Select PLNCode, ANNLmtInd From PLN Where PLNCode = '" & Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") = 0, 4, InStr(1, cboPLNCode, "-") - 1))) & "'"
        ChkAnnLimit.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If ChkAnnLimit.recordCount Then
        
            If ChkAnnLimit!ANNLmtInd = "2" Then
                lblAnnualLimit = "0.00"
            Else
                lblAnnualLimit = Format(GetCLNAnnualLimit(Trim(Mid(cboINSCode, 1, 1)), Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") = 0, 4, InStr(1, cboPLNCode, "-") - 1))), Trim(frmNewMemberRegistration.txtHLTCode), txtMBMPayorEffDate), "#,###,##0.00")
            End If
        
        End If
        ChkAnnLimit.Close
        Set ChkAnnLimit = Nothing
    End If

End Sub

Public Sub VisitLimit()
Dim ChkVisitLimit As New ADODB.Recordset
Dim strsql As String

    LblVisitLimit = ""
    
    If Len(Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") = 0, 4, InStr(1, cboPLNCode, "-") - 1)))) And IsDate(txtMBMPayorEffDate) Then
        strsql = " Select PLNCode, VISLmtInd From PLN Where PLNCode = '" & Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") = 0, 4, InStr(1, cboPLNCode, "-") - 1))) & "'"
        ChkVisitLimit.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If ChkVisitLimit.recordCount Then
        
            If ChkVisitLimit!VISLmtInd = "2" Then
                LblVisitLimit = "0.00"
            Else
                LblVisitLimit = Format(GetVisitLimit(Trim(Mid(cboINSCode, 1, 1)), Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") = 0, 4, InStr(1, cboPLNCode, "-") - 1))), Trim(frmNewMemberRegistration.txtHLTCode), txtMBMPayorEffDate), "#,###,##0.00")
            End If
        End If
        ChkVisitLimit.Close
        Set ChkVisitLimit = Nothing
    End If

End Sub

Private Sub MemberCLNCallFees()
Dim ChkPremium As New ADODB.Recordset
Dim strsql As String

    If Trim(frmNewMemberRegistration.txtPAYCode) <> "" And Trim(txtAGECode) <> "" And Trim(cboPLNCode) <> "" _
        And Trim(frmNewMemberRegistration.txtHLTCode) <> "" Then
        lblMBMPremium = "0"
        lblMCOFees = "0"
        
        If Len(Trim(frmNewMemberRegistration.txtPAYCode)) <> 0 Then
            
            If Len(Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") = 0, 4, InStr(1, cboPLNCode, "-") - 1)))) And IsDate(txtMBMPayorEffDate) Then
                strsql = " Select PLNCode, PRMInd From PLN Where PLNCode = '" & Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") = 0, 4, InStr(1, cboPLNCode, "-") - 1))) & "'"
                ChkPremium.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
                If ChkPremium.recordCount Then
        
                    If ChkPremium!PRMInd = "2" Then
                        lblMBMPremium = "0.00"
                        txtTempPrem = "0.00"
                        
                        lblMCOFees = "0.00"
                        txtTempMCOFees = "0.00"
                    Else
                        lblMBMPremium = Format(GetCLNPremium(Trim(Mid(cboINSCode, 1, 1)), Trim(frmNewMemberRegistration.txtPAYCode), Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)), _
                        Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") = 0, 4, InStr(1, cboPLNCode, "-") - 1))), Trim(frmNewMemberRegistration.txtHLTCode), txtMBMPayorEffDate), "#,##0.00")
                
                        txtTempPrem = Format(GetCLNPremium(Trim(Mid(cboINSCode, 1, 1)), Trim(frmNewMemberRegistration.txtPAYCode), Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)), _
                        Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") = 0, 4, InStr(1, cboPLNCode, "-") - 1))), Trim(frmNewMemberRegistration.txtHLTCode), txtMBMPayorEffDate), "#,##0.00")
                        
                        txtTempMCOFees = Format(GetCLNMCO(Trim(Mid(cboINSCode, 1, 1)), Trim(frmNewMemberRegistration.txtHLTCode), frmNewMemberRegistration.txtMBMAdultChild, txtMBMPayorEffDate, Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") = 0, 4, InStr(1, cboPLNCode, "-") - 1)))), "#,##0.00")
                        lblMCOFees = Format(GetCLNMCO(Trim(Mid(cboINSCode, 1, 1)), Trim(frmNewMemberRegistration.txtHLTCode), frmNewMemberRegistration.txtMBMAdultChild, txtMBMPayorEffDate, Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") = 0, 4, InStr(1, cboPLNCode, "-") - 1)))), "#,##0.00")
                    End If
                End If
            End If
            ChkPremium.Close
            Set ChkPremium = Nothing
        
            
        End If
    End If
        
End Sub

Private Sub CLNMemberCallAge()
Dim SelectAge As String
    
    If Trim(txtMBMPayorEffDate) <> "__/__/____" _
        And Len(Trim(txtMBMBirthdate)) <> 0 _
        And Trim(txtMBMBirthdate) <> "__/__/____" _
        And Len(Trim(frmNewMemberRegistration.txtHLTCode)) <> 0 Then
        If IsDate(txtMBMBirthdate) And IsDate(txtMBMPayorEffDate) Then
            SelectAge = GetAge(txtMBMPayorEffDate, txtMBMBirthdate)
            frmNewMemberRegistration.txtAGECode = GetAgeBand(Trim(frmNewMemberRegistration.txtHLTCode), SelectAge)
            If Len(Trim(txtAGECode)) <> 0 Then
                If Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)) = "01" Then
                    frmNewMemberRegistration.txtMBMAdultChild = "C"
                Else
                    frmNewMemberRegistration.txtMBMAdultChild = "A"
                End If
                'frmNewMemberRegistration.txtAGECode = txtAGECode
            End If
        End If
    End If
End Sub

Private Sub InsertCLNMCO()
Dim MBMInsertMCOFees As Double
Dim MBMMCOFess As Double
Dim PAYExpDate As String

    If Len(lblMCOFees) Then
        MBMInsertMCOFees = InsertMCOFees(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(lblMCOFees, "#,##0.00"))
        lblMCOFees = ""
        lblMCOFees = Round(MBMInsertMCOFees, 0)
    End If
End Sub

Private Sub InsertCLNPrem()
Dim MBMInsertPrem As Double
Dim MBMPrem As Double
Dim PAYExpDate As String

    If Trim(cboPLNCode) <> "" Then
            
        If Len(lblMBMPremium) And IsDate(txtMBMPayorEffDate) Then
            MBMInsertPrem = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(lblMBMPremium, "#,##0.00"))
            lblMBMPremium = ""
            lblMBMPremium = Round(MBMInsertPrem, 0)
        End If
            
    End If
    
End Sub

Public Function showMBMClinical() As Integer
Dim sqlstr As String
Dim MBM As New ADODB.Recordset
    
    
    sqlstr = " SELECT * From View_MBM_Clinical Where MBMNumber ='" & Trim(frmMBMEnquiry.txtMBMNumber) & "'"
    
    MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    
    If MBM.recordCount = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
        With MBM
            
            If Len(!PLNType) Then
                cboPLNCode = Trim(!PLNType)
            End If
            
            If Len(!InsType) Then
                cboINSCode = Trim(!InsType)
            End If
            
            If Len(!MBMDOB) Then
                txtMBMBirthdate = Format(!MBMDOB, "dd/mm/yyyy")
            End If
            
            If Len(!MBMPolEffDate) Then
                txtMBMPayorEffDate = Format(!MBMPolEffDate, "dd/mm/yyyy")
            End If
            
            If Len(!MBMPolExpDate) Then
                txtMBMPayorExpDate = Format(!MBMPolExpDate, "dd/mm/yyyy")
            End If
            
            If Len(!MBMRemarks) Then
                txtRemarks = Trim(!MBMRemarks)
            End If
            
            If Len(!MBMCLNPremium) Then
                lblMBMPremium = Format(!MBMCLNPremium, "#,##0.00")
            End If
            
            If Len(!MBMCLNMCOFee) Then
                lblMCOFees = Format(!MBMCLNMCOFee, "#,##0.00")
            End If
            
            If Len(!MBMCLNBalLmt) Then
                lblAnnualLimit = Format(!MBMCLNBalLmt, "#,##0.00")
            End If
            
            If Len(!MBMCLNVisitLmt) Then
                LblVisitLimit = Format(!MBMCLNVisitLmt, "#,##0.00")
            End If
            
            If !MBMCLNMemberType = "P" Then
                cboMBMRelCode = "C - Combined Limit"
            ElseIf !MBMCLNMemberType = "Q" Then
                cboMBMRelCode = "S - Separated Limit"
            Else
                cboMBMRelCode = ""
            End If
            
        
        End With
        
    End If
    MBM.Close
    Set MBM = Nothing
End Function

