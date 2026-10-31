VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUpdateInvoice 
   Caption         =   "Update Invoice"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2235
   ScaleWidth      =   7785
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   7575
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton cmdShow 
         Appearance      =   0  'Flat
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   300
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   4920
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtVersion 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         MaxLength       =   1
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtCaseNo 
         Height          =   285
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtHospName 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtInvNo 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtInvDate 
         Height          =   300
         Left            =   4080
         TabIndex        =   5
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "Ver"
         Height          =   255
         Left            =   3000
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Invoice Number"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Invoice Date"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Hospital Name"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Case Number"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   7575
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   315
         Left            =   3960
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   315
         Left            =   6360
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   5160
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Update Invoice"
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
      Height          =   220
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   7725
   End
End
Attribute VB_Name = "frmUpdateInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub EntriesEnable(status As Boolean)
    
    txtCaseNo.Enabled = status
    txtHospName.Enabled = status
    txtVersion.Enabled = status
    cmdUpdate.Enabled = status
    
End Sub

Private Sub cboHospital_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHospital.Text, cboHospital.SelStart) + Chr(KeyAscii) + Mid(cboHospital.Text, cboHospital.SelStart + cboHospital.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboHospital.ListCount - 1
 
        If NewText = LCase(Left(cboHospital.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHospital.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboHospital.Text = ValidValue
                cboHospital.SelStart = 0
                cboHospital.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cmdClear_Click()

    Call ClearForm(Me)
    Call EntriesEnable(True)
    txtCaseNo.SetFocus
    cboHospital.Enabled = False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdShow_Click()
    If txtCaseNo = "" Then
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
        txtCaseNo.SetFocus
    ElseIf txtVersion = "" Then
        MsgBox "Please Enter Case Version", vbOKOnly + vbCritical, DialogBoxTitle
        txtVersion.SetFocus
    Else
        Call ShowInvoice
    End If

End Sub

Private Sub ShowInvoice()
Dim sql As String
Dim INV As New ADODB.Recordset
Dim strsqlHOS As String
Dim HOs As New ADODB.Recordset


    sql = "Select InvNo,InvDate,CASNumber,INVWSversion, INVPayee from ClaimInvoice where CASNumber = '" & Trim(txtCaseNo) & "' and INVWSversion = '" & Trim(txtVersion) & "'"
    INV.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If INV.recordCount = 0 Then
       MsgBox "No record found!", vbCritical + vbOKOnly
       cmdUpdate.Enabled = False
    Else
        With INV
            txtCaseNo = !CasNumber
            txtVersion = !INVWSversion
            
            If Len(!INVNo) Then
                txtInvNo = Trim(!INVNo)
            End If
            
            If Len(!InvDate) Then
                txtInvDate.mask = ""
                txtInvDate = Format(!InvDate, "dd/mm/yyyy")
                txtInvDate.mask = "##/##/####"
            End If
            cboHospital.Enabled = False
            strsqlHOS = "select HOSCode, HOSAccCode, HOSName  from HOS where HOSCode = '" & INV!INVPayee & "'"
            HOs.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            If HOs.recordCount Then
                If Len(HOs!HOSName) Then
                    cboHospital.Enabled = True
                    cboHospital.Text = Trim(HOs!HOSAccCode) & "-" & Trim(HOs!HOSName)
                End If
            Else
                
            End If
            HOs.Close
            Set HOs = Nothing
            
            Call EntriesEnable(False)
            
            cmdUpdate.Enabled = True
            If Len(Trim(cboHospital.Text)) Then
                cboHospital.SetFocus
            Else
                txtInvNo.SetFocus
            End If
        End With
    End If

    INV.Close
    Set INV = Nothing
End Sub

Private Sub cmdUpdate_Click()
Dim UpdateInv As String
Dim CAS As New ADODB.Recordset
Dim HOspitals As String
Dim HOSSql As String
Dim HOs As New ADODB.Recordset
Dim AA As Integer
Dim HOspitalName As String

    If txtCaseNo = "" Then
        MsgBox "Please Enter Case No!", vbCritical + vbOKOnly, DialogBoxTitle
        txtCaseNo.SetFocus
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            UpdateInv = "Select InvNo,InvDate,INVWSversion,INVPayee from ClaimInvoice Where CASNumber ='" & Trim(txtCaseNo) & "' and INVWSversion = '" & Trim(txtVersion) & "'"
            CAS.Open UpdateInv, medixmasterconnWS, adOpenKeyset, adLockOptimistic

            With CAS
                !INVNo = txtInvNo
                !InvDate = CDate(txtInvDate)
                HOspitals = Replace(cboHospital, "'", "''")
                AA = InStr(1, HOspitals, "-")
                HOspitalName = Trim(Right(HOspitals, IIf(AA - 1 <= 0, 0, Len(Trim(HOspitals)) - AA))) 'PPPP

                HOSSql = "select HosCode, HOSAccCode, HosName from HOS where HosName = '" & Trim(HOspitalName) & "'"
                HOs.Open HOSSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                
                If HOs.recordCount Then
                    !INVPayee = HOs!HOSCode
                End If
                HOs.Close
                Set HOs = Nothing
            End With
            
            CAS.Update
            
            CAS.Close
            Set CAS = Nothing
            
            MsgBox "Record Updated Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
            'cboHospital.Enabled = False
        End If
    End If
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
Dim HOs As New ADODB.Recordset

    HOs.Open "Select HOSCode,HOSAccCode, HOSName from HOS ORDER BY HOSAccCode ", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until HOs.EOF
        cboHospital.AddItem HOs!HOSAccCode & " - " & HOs!HOSName
        HOs.MoveNext
    Loop
    HOs.Close
    Set HOs = Nothing

End Sub

Private Sub txtInvDate_LostFocus()
    If IsDate(txtInvDate) Then
    Else
        If txtInvDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtInvDate.SetFocus
        End If
    End If
End Sub
