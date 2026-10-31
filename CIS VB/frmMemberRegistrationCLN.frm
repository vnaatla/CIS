VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMemberRegistrationCLN 
   Caption         =   "Membership Registration"
   ClientHeight    =   4056
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   5748
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4056
   ScaleWidth      =   5748
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   580
      Left            =   50
      TabIndex        =   32
      Top             =   240
      Width           =   5655
      Begin VB.CommandButton cmbSearchMBM 
         Caption         =   "Search"
         Default         =   -1  'True
         Height          =   280
         Left            =   3720
         TabIndex        =   1
         Top             =   160
         Width           =   735
      End
      Begin VB.TextBox txtPatientIC 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   150
         Width           =   2475
      End
      Begin VB.Label Label75 
         Caption         =   "IC No"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   180
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   50
      TabIndex        =   18
      Top             =   800
      Width           =   5655
      Begin VB.TextBox txtPatientName 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   195
         Width           =   4375
      End
      Begin VB.TextBox txtIC 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   480
         Width           =   4395
      End
      Begin VB.ComboBox cboRace 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   770
         Width           =   1695
      End
      Begin VB.ComboBox cboGender 
         Height          =   315
         Left            =   3600
         TabIndex        =   5
         Top             =   770
         Width           =   1995
      End
      Begin VB.ComboBox cboStaffCat 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   1050
         Width           =   1695
      End
      Begin VB.TextBox txtEmpNo 
         Height          =   315
         Left            =   3600
         TabIndex        =   7
         Top             =   1050
         Width           =   1995
      End
      Begin MSMask.MaskEdBox txtEffDate 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   1330
         Width           =   975
         _ExtentX        =   1715
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtExpDate 
         Height          =   315
         Left            =   3000
         TabIndex        =   9
         Top             =   1340
         Width           =   1035
         _ExtentX        =   1820
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOB 
         Height          =   315
         Left            =   4560
         TabIndex        =   10
         Top             =   1340
         Width           =   1035
         _ExtentX        =   1820
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtMBMGroupCompany 
         Height          =   315
         Left            =   4800
         TabIndex        =   11
         Top             =   2400
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.ComboBox cboGrpCompany 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   44
         Top             =   1590
         Width           =   4395
      End
      Begin VB.ComboBox cboDept 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   1880
         Width           =   4395
      End
      Begin VB.Label lblAdultChild 
         Caption         =   "AdultChild"
         Height          =   255
         Left            =   3840
         TabIndex        =   42
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblAgeband 
         Caption         =   "AgeBand"
         Height          =   255
         Left            =   4080
         TabIndex        =   41
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "IC No"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblAnnualLimit 
         Caption         =   "BalLmt"
         Height          =   255
         Left            =   4440
         TabIndex        =   39
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblVisitLimit 
         Caption         =   "VisitLmt"
         Height          =   255
         Left            =   4200
         TabIndex        =   38
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblPLN 
         Caption         =   "PlnCode"
         Height          =   255
         Left            =   4320
         TabIndex        =   37
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblInsType 
         Caption         =   "InsType"
         Height          =   255
         Left            =   4200
         TabIndex        =   36
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Grp Company"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   35
         Top             =   1600
         Width           =   1215
      End
      Begin VB.Label lblPayor 
         Caption         =   "Payor"
         Height          =   255
         Left            =   3960
         TabIndex        =   34
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblMBMNo 
         Caption         =   "New MBM No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   31
         Top             =   2235
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblMBMNumber 
         Caption         =   "New MBM No :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   30
         Top             =   2235
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "DOB"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   28
         Top             =   1370
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Outlet/Dept"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Gender"
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   25
         Top             =   820
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Emp No"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   24
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   225
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Staff Cat"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Race"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   800
         Width           =   615
      End
      Begin VB.Label Label36 
         Caption         =   "Exp Date"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   20
         Top             =   1370
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1370
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   50
      TabIndex        =   26
      Top             =   3280
      Width           =   5655
      Begin VB.CommandButton cmdReg 
         Caption         =   "Save"
         Height          =   255
         Left            =   4320
         TabIndex        =   15
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   255
         Left            =   4920
         TabIndex        =   16
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         Top             =   160
         Width           =   615
      End
      Begin VB.TextBox txtDept 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   120
         Width           =   2115
      End
      Begin VB.Label Label2 
         Caption         =   "Outlet/Dept"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   29
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.Label lblMsg 
      Caption         =   "* Double-Click Member No to proceed to Invoice Processing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   3800
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Membership Registration"
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
      TabIndex        =   17
      Top             =   0
      Width           =   5805
   End
End
Attribute VB_Name = "frmMemberRegistrationCLN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StaffCat As Collection

Private Sub cboDept_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboDept.Text, cboDept.SelStart) + Chr(KeyAscii) + Mid(cboDept.Text, cboDept.SelStart + cboDept.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboDept.ListCount - 1
 
        If NewText = LCase(Left(cboDept.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboDept.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboDept.Text = ValidValue
                cboDept.SelStart = 0
                cboDept.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboGender_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboGender.Text, cboGender.SelStart) + Chr(KeyAscii) + Mid(cboGender.Text, cboGender.SelStart + cboGender.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboGender.ListCount - 1
 
        If NewText = LCase(Left(cboGender.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboGender.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboGender.Text = ValidValue
                cboGender.SelStart = 0
                cboGender.SelLength = Len(ValidValue)
            End If
        End If
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
                ValidValue = cboDept.List(i)
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

Private Sub CboRace_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboRace.Text, cboRace.SelStart) + Chr(KeyAscii) + Mid(cboRace.Text, cboRace.SelStart + cboRace.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboRace.ListCount - 1
 
        If NewText = LCase(Left(cboRace.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboRace.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboRace.Text = ValidValue
                cboRace.SelStart = 0
                cboRace.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboStaffCat_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboStaffCat.Text, cboStaffCat.SelStart) + Chr(KeyAscii) + Mid(cboStaffCat.Text, cboStaffCat.SelStart + cboStaffCat.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboStaffCat.ListCount - 1
 
        If NewText = LCase(Left(cboStaffCat.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboRace.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboStaffCat.Text = ValidValue
                cboStaffCat.SelStart = 0
                cboStaffCat.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboStaffCat_LostFocus()
Dim strsql As String
Dim CAT As New ADODB.Recordset
    lblInsType = ""
    lblPLN = ""
    
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    
        strsql = " Select INSType, PlanCode,Category from WebCategory where Category = '" & cboStaffCat & "'"
        CAT.Open strsql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            
        Do While Not CAT.EOF
            lblPLN = CAT!plancode
            lblInsType = CAT!INSType
            CAT.MoveNext
        Loop
        CAT.Close
        Set CAT = Nothing
        
        If txtEffDate <> "__/__/____" And cboStaffCat <> "" Then
            Call CallCLNLimit
        End If
   ' MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
End Sub

Private Sub cmbSearchMBM_Click()
Dim tmpICNumber As String
    Screen.MousePointer = vbHourglass
    
       ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
       ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
            tmpICNumber = txtPatientIC
            Call ClearForm(Me)
            lblMBMNumber.Visible = False
            lblMBMNo.Visible = False
            lblMBMNo = ""
            txtPatientIC = tmpICNumber
            cmdReg.Enabled = True
            lblMsg.Visible = False
            Call MBMDetails
            
      '  MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
       ' medixmasterconn.Close
      '  Set medixmasterconn = Nothing
       ' MediConnCISWS.Close
       ' Set MediConnCISWS = Nothing
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub MBMDetails()
Dim sql As String
Dim ADORSTMBM As New ADODB.Recordset
    sql = ""
    sql = "Exec  SP_SearchCISMBM_IC '" & Trim(txtPatientIC) & "'"
    
    ADORSTMBM.Open sql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    
    If Not ADORSTMBM.EOF Then
        With ADORSTMBM
            lblPayor = !PAYCode
            txtPatientName = !MBMName
            txtIC = txtPatientIC
            cboRace = IIf(Len(!RACCode), !RACCode, "")
            cboGender = IIf(Len(!MBMSex), !MBMSex, "")
            txtPatientIC = IIf(Len(!MBMIcBcPp), !MBMIcBcPp, "")
            txtDOB = IIf(Len(!MBMDOB), !MBMDOB, "__/__/____")
            
            If CDate(!MBMPolEXPDate) > CDate(Date) Then
                txtEffDate = !MBMPolEffDate
                txtExpDate = !MBMPolEXPDate
                cboStaffCat = "CREW"
            End If
            cboDept = IIf(Len(!MBMDepartment), !MBMDepartment, "")
            
            cboGrpCompany = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
            
            txtEmpNo = IIf(Len(!MBMEmployeeNo), !MBMEmployeeNo, "")
            
            If CDate(!MBMPolEXPDate) > CDate(Date) Then
                lblMBMNo.Visible = True
                lblMBMNumber.Visible = True
                lblMBMNo = !MBMNumber
                cmdReg.Enabled = False
                lblMsg.Visible = True
            Else
                cmdReg.Enabled = True
            End If
        'Call CallCLNLimit
        'lblAnnLmt = tmpAnnLmt
        'lblMBMVisitLimit = tmpVisitLmt
        End With
    Else
        MsgBox "No record found!", vbCritical
    End If
    ADORSTMBM.Close
    Set ADORSTMBM = Nothing
End Sub

Private Sub cmdReg_Click()
On Error GoTo ErrorSave
Dim MembershipNo As String
Dim TmpChkDigits As String
Dim TempMBMName As String
Dim TempDOB As String
Dim TempPayEffDate As String
Dim TempPayExpDate As String
Dim TempValueDate As String
Dim strsqlMBMOth As String
Dim tempGrpCompany As String
Dim tempDept As String
Dim strsqlMBM As String
Dim strsqlMBMTwo As String
Dim strsqlMBMCRef As String
Dim strsqlMBMOthTwo As String
Dim ClnOthersSql As String
Dim startTrans As Boolean

    If txtPatientIC = "" And txtIC = "" Then
        MsgBox "Please enter IC No! ", vbCritical
    ElseIf txtPatientName = "" Then
        MsgBox "Please enter Member Name! ", vbCritical
    ElseIf cboRace = "" Then
        MsgBox "Please enter Race Code! ", vbCritical
    ElseIf cboGender = "" Then
        MsgBox "Please enter Gender! ", vbCritical
    ElseIf cboStaffCat = "" Then
        MsgBox "Please enter Staff Category! ", vbCritical
    ElseIf txtEmpNo = "" Then
        MsgBox "Please enter Employee No! ", vbCritical
    ElseIf txtEffDate = "__/__/____" Then
        MsgBox "Please enter Effective Date! ", vbCritical
    ElseIf txtExpDate = "__/__/____" Then
        MsgBox "Please enter Expiry Date! ", vbCritical
    ElseIf txtDOB = "__/__/____" Then
        MsgBox "Please enter Birth Date! ", vbCritical
    ElseIf cboGrpCompany = "" Then
        MsgBox "Please enter Group Company! ", vbCritical
    ElseIf cboDept = "" Then
        MsgBox "Please enter Department! ", vbCritical
    Else
        
        Screen.MousePointer = vbHourglass
        
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
       ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
       ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
            TempMBMName = Replace(Trim(txtPatientName), "'", "''")
            TempDOB = Format(txtDOB, "yyyy/mm/dd")
            TempValueDate = Format(Date, "yyyy/mm/dd")
            TempPayEffDate = Format(txtEffDate, "yyyy/mm/dd")
            TempPayExpDate = Format(txtExpDate, "yyyy/mm/dd")
            tempGrpCompany = Replace(Trim(cboGrpCompany), "'", "''")
            tempDept = Replace(Trim(cboDept), "'", "''")
            
            MembershipNo = "DC" & Trim(lblInsType) & GenerateMembershipNo("DC", Trim(lblInsType))
            TmpChkDigits = GenerateChkDigits
            MembershipNo = MembershipNo & TmpChkDigits
            strsqlMBM = ""
            
            startTrans = True
            medixmasterconn.beginTrans
            medixmasterconnMaint.beginTrans
            
            strsqlMBM = "insert into MBM (MBMNumber, MBMIcBcPp, MBMName, " & _
                        "MBMSex, MBMAdultChild, AGECode, MBMValueDate, " & _
                        "MBMBordxDate, MBMMemberType, HLTCode, MBMBatchNo, MBMCOVID, HSCLNInd,MBMExpiredStatus,MBMDOB,RACCode) " & _
                        "Values('" & MembershipNo & "', '" & Trim(txtPatientIC) & "' , '" & TempMBMName & "', " & _
                        "'" & cboGender & "', '" & Mid(lblAdultChild, 1, 1) & "', '" & Mid(lblAgeband, 1, 2) & "', '" & TempValueDate & "', " & _
                        "'" & TempValueDate & "', 'P', 'N', '1','00','C','N','" & TempDOB & "','" & cboRace & "')"
            medixmasterconn.Execute strsqlMBM
            
        
            strsqlMBMTwo = "insert into MBMTwo (MBMNumber, MBMSalutation, MBMTelnoH, " & _
                        "MBMTelNoM, MBMTelNoO, MBMEmail, " & _
                        "NATCode, SRVCodeList, LGNCode, MBMRegUser, MBMLastUpdateUser, MBMUploadStatus) " & _
                        "Values('" & MembershipNo & "', '' , '', " & _
                        "'', '', '', " & _
                        "'', 'ME', '', '" & Logon & "', '" & Logon & "','M')"
            medixmasterconn.Execute strsqlMBMTwo
                
    
    
            strsqlMBMOth = "insert into MBMOthers (MBMNumber, MBMGroupCompany, MBMEmployeeNo, MBMDepartment, RELCODE )" & _
                            "Values('" & MembershipNo & "', '" & tempGrpCompany & "', '" & txtEmpNo & "', '" & tempDept & "', 'P')"
            medixmasterconn.Execute strsqlMBMOth
    
            strsqlMBMOthTwo = "insert into MBMOthersTwo (MBMNumber)" & _
                        "Values('" & MembershipNo & "') "
            medixmasterconn.Execute strsqlMBMOthTwo
    
            strsqlMBMCRef = ""
            strsqlMBMCRef = "insert into MBMCrossReference (MBMNumber, MBMIcBcPp, MBMName, " & _
                        " MBMAdultChild, AGECode, " & _
                        "MBMBordxDate, MBMMemberType, HLTCode, MBMBatchNo, MBMCOVID, MBMStatus) " & _
                        "Values('" & MembershipNo & "', '" & Trim(txtPatientIC) & "' , '" & TempMBMName & "', " & _
                        "'" & Mid(lblAdultChild, 1, 1) & "', '" & Mid(lblAgeband, 1, 2) & "', " & _
                        "'" & TempValueDate & "', 'P', 'N', '1','00', 'A')"
            medixmasterconnMaint.Execute strsqlMBMCRef
    
            ClnOthersSql = "Insert Into MBMCLNOthers(MBMNumber,  PAYCode, " & _
                         "PLNType, InsType, MBMPolEffDate, MBMPolExpDate, " & _
                         "MBMCLNMemberType, MBMCLNExpiryStatus, MBMCLNDataStatus, MBMCLNStatus, " & _
                         "MBMCLNVisitLmt, MBMCLNBalLmt)" & _
                         "Values('" & MembershipNo & "','DC'," & _
                         "'" & Trim(lblPLN) & "','" & lblInsType & "','" & TempPayEffDate & "', '" & TempPayExpDate & "'," & _
                         "'P','N','A','A'," & _
                         "" & lblVisitLimit & "," & lblAnnualLimit & ")"
            medixmasterconn.Execute ClnOthersSql
            medixmasterconn.commitTrans
            medixmasterconnMaint.commitTrans
            startTrans = False
            lblMBMNo.Visible = True
            lblMBMNumber.Visible = True
            lblMBMNo = MembershipNo
            lblMsg.Visible = True
            cmdReg.Enabled = False
        
     '   MediConnCISDB.Close
     '   Set MediConnCISDB = Nothing
     '   medixmasterconn.Close
     '   Set medixmasterconn = Nothing
     '   MediConnCISWS.Close
     '   Set MediConnCISWS = Nothing
     '   medixmasterconnMaint.Close
     '   Set medixmasterconnMaint = Nothing
        
        MsgBox "Membership Generated, Your Membership Number is : " & MembershipNo, vbOKOnly + vbInformation, DialogBoxTitle
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If startTrans = True Then
        medixmasterconn.RollbackTrans
        medixmasterconnMaint.RollbackTrans
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        'MediConnCISWS.Close
        'Set MediConnCISWS = Nothing
        'medixmasterconnMaint.Close
        'Set medixmasterconnMaint = Nothing
    End If
    MsgBox Err.Description

End Sub

Private Sub CmdSearch_Click()
Dim strsql As String
Dim DEPT As New ADODB.Recordset
    strsql = ""
    If txtDept = "" Then
        MsgBox "Please enter Department! ", vbCritical
    Else
      '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
      '  MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
            strsql = " Select GrpCompCode, GrpCompDept From GroupCompanyDept Where GrpCompCode = 'G001' And GrpCompDept  LIKE '%" & Trim(txtDept) & "%' "
            DEPT.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            cboDept.Clear
            Do While Not DEPT.EOF
                cboDept.AddItem DEPT!GrpCompDept
                DEPT.MoveNext
            Loop
            DEPT.Close
            Set DEPT = Nothing
        
       ' MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
       ' MediConnCISWS.Close
       ' Set MediConnCISWS = Nothing
    End If
End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "frmMemberRegistrationCLN")
    Call ComboListing
    'Call MBMDetails
End Sub

Private Sub ComboListing()
Dim RAC As New ADODB.Recordset
Dim CAT As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strsql As String
Dim GRPCompany As New ADODB.Recordset


    cboRace.Clear
    cboGender.Clear
    cboStaffCat.Clear
    
 ''   MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "SearchRace"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 15
        RAC.CursorLocation = adUseClient
        RAC.CursorType = adOpenForwardOnly
        RAC.CacheSize = 100
        Set RAC = cmd.Execute
    
        Do Until RAC.EOF
            With RAC
                cboRace.AddItem !RACCode
                RAC.MoveNext
            End With
        Loop
        RAC.Close
        Set RAC = Nothing
        
        cboGender.AddItem "M"
        cboGender.AddItem "F"
        
        'Set CAT = New Collection
        strsql = " Select INSType, PlanCode,Category from WebCategory"
        CAT.Open strsql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
            
        Do Until CAT.EOF
            cboStaffCat.AddItem CAT!Category
            'StaffCat.Add CAT!INSType, CAT!plancode, CAT!Category
            CAT.MoveNext
        Loop
        
        CAT.Close
        Set CAT = Nothing
        


   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        cboGrpCompany.Clear
        GRPCompany.Open "Select GRPCoCode, GRPCoName from GRPCompany ", MediConnCISDB, adOpenKeyset, adLockOptimistic
            
        Do Until GRPCompany.EOF
            cboGrpCompany.AddItem Trim(GRPCompany!GrpCoName)
            GRPCompany.MoveNext
        Loop
        
        GRPCompany.Close
        Set GRPCompany = Nothing
        
 '   MediConnCISDB.Close
 '   Set MediConnCISDB = Nothing
 '   medixmasterconn.Close
 '   Set medixmasterconn = Nothing
 '   MediConnCISWS.Close
 '   Set MediConnCISWS = Nothing
End Sub

Private Sub lblMBMNo_DblClick()
    frmInvoice.txtMBMNo = lblMBMNo
    frmInvoice.CmdGo_Click
    Unload Me
End Sub

Private Sub txtDOB_LostFocus()
    If txtDOB <> "__/__/____" Then
        If IsDate(txtDOB) Then
            If txtEffDate <> "__/__/____" Then
                Call CLNMemberCallAge
            End If
        Else
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOB.SetFocus
        End If
    End If
End Sub

Private Sub txtEffDate_LostFocus()
Dim tmpExpDate As String

  '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        If txtEffDate <> "__/__/____" Then
            If IsDate(txtEffDate) Then
                If txtEffDate <> "__/__/____" And cboStaffCat <> "" Then
                    Call CallCLNLimit
                End If
                If txtEffDate <> "__/__/____" Then
                    tmpExpDate = "31/12/" & Year(txtEffDate)
                    txtExpDate = tmpExpDate
                End If
                If txtDOB <> "__/__/____" Then
                    Call CLNMemberCallAge
                End If
    
            Else
                MsgBox "Invalid Date Format! ", vbCritical
                 txtEffDate.SetFocus
            End If
        End If
        
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
End Sub
Private Sub CLNMemberCallAge()
Dim SelectAge As String
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        If Trim(txtEffDate) <> "__/__/____" And txtDOB <> "__/__/____" Then
            If IsDate(txtDOB) And IsDate(txtEffDate) Then
    
                SelectAge = GetAge(txtEffDate, CDate(txtDOB))
                lblAgeband = GetAgeBand(SelectAge)
    
                If Len(Trim(lblAgeband)) <> "" Then
                    If Trim(Mid(lblAgeband, 1, InStr(1, lblAgeband, "-") - 1)) = "01" Then
                        lblAdultChild = "C"
                    Else
                        lblAdultChild = "A"
                    End If
                End If
            End If
        End If
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub

Private Sub txtExpDate_LostFocus()
    If IsDate(txtExpDate) Then
    Else
        If txtExpDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtExpDate.SetFocus
        End If
    End If
End Sub

Public Sub CallCLNLimit()
Dim ChkAnnLimit As New ADODB.Recordset
Dim strsql As String
Dim ChkVisitLimit As New ADODB.Recordset
Dim tmpPln As String
Dim tmpIns As String

    lblAnnualLimit = ""
    lblVisitLimit = ""
    If Trim(lblPLN) <> "" And IsDate(txtEffDate) Then
        strsql = " Select PLNCode, ANNLmtInd From PLN Where PLNCode = '" & lblPLN & "'"
        ChkAnnLimit.Open strsql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not ChkAnnLimit.EOF
            If ChkAnnLimit!AnnLmtIND = "N" Then
                lblAnnualLimit = "0.00"
            Else
                lblAnnualLimit = Format(GetCLNAnnualLimit(lblInsType, lblPLN, "N", txtEffDate), "#,###,##0.00")
            End If
            ChkAnnLimit.MoveNext
        Loop
        ChkAnnLimit.Close
        Set ChkAnnLimit = Nothing
        
        strsql = " Select PLNCode, VISLmtInd From PLN Where PLNCode = '" & Trim(lblPLN) & "'"
        ChkVisitLimit.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If ChkVisitLimit.RecordCount Then
            If ChkVisitLimit!VisLmtIND = "N" Then
                lblVisitLimit = "0.00"
            Else
                lblVisitLimit = GetVisitLimit(lblInsType, lblPLN, "N", txtEffDate, "P", "DC")
            End If
        End If
        
        ChkVisitLimit.Close
        Set ChkVisitLimit = Nothing
    End If
End Sub

