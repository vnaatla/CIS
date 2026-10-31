VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3504
   ClientLeft      =   2808
   ClientTop       =   12
   ClientWidth     =   3600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleMode       =   0  'User
   ScaleWidth      =   3456
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCheckLogon 
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
      Left            =   45
      TabIndex        =   8
      Top             =   3195
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txtnewpass 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2115
      Width           =   2295
   End
   Begin VB.TextBox txtnewconfirm 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2565
      Width           =   2295
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   2565
      TabIndex        =   5
      Top             =   3105
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1245
      TabIndex        =   2
      Top             =   1035
      Width           =   1022
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2460
      TabIndex        =   3
      Top             =   1035
      Width           =   1022
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1260
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   495
      Width           =   2295
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2430
      TabIndex        =   4
      Top             =   1035
      Visible         =   0   'False
      Width           =   1022
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3412.8
      Y1              =   1664.713
      Y2              =   1664.713
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter a new Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   13
      Top             =   1755
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "New Password"
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
      Left            =   45
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Re-Type New Password"
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
      Left            =   45
      TabIndex        =   11
      Top             =   2565
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
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
      Index           =   1
      Left            =   45
      TabIndex        =   10
      Top             =   585
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
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
      Index           =   0
      Left            =   45
      TabIndex        =   9
      Top             =   90
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim counter As Integer
Dim CurrentUser As String
Dim tmpEncryptDeCryptPass As String
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChange_Click()
    If txtUserName = "" And txtPassword = "" Then
        MsgBox "Please Enter User ID and Password", vbOKOnly + vbCritical, DialogBoxTitle
        txtUserName.SetFocus
    Else
        Me.Height = 4100
        txtnewpass.SetFocus
    End If
End Sub

Private Sub cmdChange_LostFocus()
    txtUserName.SetFocus
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
Dim RecordSaved
Dim StrSql As String
Dim USR As New ADODB.Recordset
Dim Logincnt As Integer
Dim fMainForm As New MDIMain
    
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    If txtUserName = "" Then
        MsgBox "Please Enter User Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtUserName.SetFocus
    Else
        If counter = 3 Then
            StrSql = "UPDATE USR SET USRStatus = 'S' WHERE (USRCode = '" & txtUserName & "')"
            USR.Open StrSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            MsgBox "Your Account Been Suspended", vbOKOnly + vbCritical, DialogBoxTitle
            counter = 0
        Else
            StrSql = "Select * from USR Where USRCode ='" & Trim(txtUserName) & "' And USRStatus ='A'"
            USR.Open StrSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            
                If USR.RecordCount = 0 Then
                    MsgBox "Invalid User Name!! Please enter again..", vbOKOnly + vbCritical, DialogBoxTitle
                    txtUserName = ""
                    txtPassword = ""
                    txtUserName.SetFocus
                    USR.Close
                Else
                    Call Crypt(txtPassword)
                    If UCase(tmpEncryptDeCryptPass) <> UCase(USR!USRPassword) Then
                        MsgBox "Invalid Password!! Please enter again..", vbOKOnly + vbCritical, DialogBoxTitle
                        counter = counter + 1
                        If counter = 1 Then
                            txtCheckLogon = txtUserName
                        Else
                            If txtCheckLogon = txtUserName Then
                                counter = counter + 1 - 1
                            Else
                                counter = 1
                                txtCheckLogon = txtUserName
                            End If
                        End If
                        Debug.Print counter
                        txtPassword = ""
                        txtPassword.SetFocus
                        USR.Close
                    Else
                        Access = USR!USRAccess
                        Logincnt = IIf(Len(USR!CIScnt), USR!CIScnt, 0)
                        Logon = UCase(txtUserName)
                        Call OutputToTxtfile
                        USR.Close
                            
                           'If Logincnt < 2 Then
                                'Update Login cnt
                                 StrSql = "Update USR set CIScnt= ISNULL(CIScnt,0) + 1, actpass='" & txtPassword & "' Where USRCode ='" & Trim(txtUserName) & "' And USRStatus ='A'"
                                 USR.Open StrSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                                 
                                 Dim systime As String
                                 Dim SysDt As String
                                 
                                 systime = CStr(DateTime.Time)
                                 SysDt = CStr(DateTime.Date)
                                 
                                 StrSql = "Insert into Tbl_USRTrack (USRCode,SYSName,LoginTime,LoginDate) Values ('" + txtUserName + "','CIS','" + systime + "','" + SysDt + "')"
                                 USR.Open StrSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                                 
                                     
                                Set fMainForm = New MDIMain
                                Load fMainForm
                                Unload Me
                                fMainForm.Show
                           'Else
                           '     MsgBox "You already Login 2 CIS Systems.", vbOKOnly + vbCritical, DialogBoxTitle
                           'End If
                        End If
                    End If
                End If
        End If
End Sub
'------------------Code For BTBG amount calculation -------------
Public Sub UpdateFeldaBTBGAmt()
Const METHOD_NAME = "UpdateFeldaBTBGAmt"
On Error GoTo debugs

Dim GetRs As ADODB.Recordset
Dim SqlQry As String
Dim InsrtQry As String
Dim MBMNumber As String
Dim UpdateRS As ADODB.Recordset
Dim FrstRecord As Boolean
Dim IsFirst As Boolean
Dim SumVal As Double
Dim SumVal1 As Double
FrstRecord = True
IsFirst = True
SqlQry = ""
Dim CASNum As String
CASNum = ""
InsrtQry = ""
SqlQry = "select  DISTINCT   tmp_ClaimAbove500.Amount, tmp_ClaimAbove500.MBMNumber,CISWorksheet..Claim.CLMTotalActual,RegisterDate,CISCas.CasNumber  from tmp_ClaimAbove500 " & _
        " inner join CISCas on  CISCas.MBMNumber = tmp_ClaimAbove500 .MBMNumber inner join CISWorksheet..Claim on CISWorksheet..Claim.CasNumber = CISCas.CasNumber " & _
        "order by tmp_ClaimAbove500 .MBMNumber,RegisterDate "
        '" Where CISWorksheet..Claim.BordxRefNo is not null  order by tmp_ClaimAbove500 .MBMNumber,RegisterDate "

Set GetRs = Nothing
Set GetRs = New ADODB.Recordset
GetRs.Open SqlQry, MediConnCISDB, adOpenDynamic, adLockReadOnly
Do While Not GetRs.EOF
    With GetRs
        CASNum = !CasNumber
        If FrstRecord Then
            MBMNumber = !MBMNumber
            FrstRecord = False
            SumVal = 0
        End If
        
        If MBMNumber = !MBMNumber Then
            SumVal = SumVal + CDbl(!CLMTotalActual)
                    If SumVal > !Amount Then
                        If IsFirst Then
                            SumVal = SumVal - CDbl(!CLMTotalActual)
                            SumVal1 = !Amount - SumVal
                            Set UpdateRS = Nothing
                            Set UpdateRS = New ADODB.Recordset
                            UpdateRS.Open "Tbl_tmpMixedCase", MediConnCISDB, adOpenDynamic, adLockOptimistic
                            UpdateRS.AddNew
                            UpdateRS!CasNumber = CASNum
                            UpdateRS!PendingAmount = SumVal
                            UpdateRS!CalValue = SumVal1
                            UpdateRS.update
                            UpdateRS.Close
                            SumVal = SumVal + CDbl(!CLMTotalActual)
                            IsFirst = False
                        Else
                            InsrtQry = "Update CISWorksheet..CLaim  set CLMBTBGAmt = (CLMTotalActual * 0.2) ,CLMFeldaFullAmt =0,CLMFeldaASO80=(CLMTotalActual * 0.8),CLMFeldaASO20=(CLMTotalActual * 0.2)  where CASNumber ='" + CASNum + "' "
                            Set UpdateRS = Nothing
                            Set UpdateRS = New ADODB.Recordset
                            UpdateRS.Open InsrtQry, MediConnCISWS, adOpenDynamic, adLockOptimistic
                            SumVal = SumVal + CDbl(!CLMTotalActual)
                        End If
                    Else
                            InsrtQry = ""
                            InsrtQry = "Update CISWorksheet..CLaim  set CLMFeldaFullAmt = " + CStr(!CLMTotalActual) + "  where CASNumber ='" + CASNum + "' "
                            Set UpdateRS = Nothing
                            Set UpdateRS = New ADODB.Recordset
                            UpdateRS.Open InsrtQry, MediConnCISWS, adOpenDynamic, adLockOptimistic
                            Set UpdateRS = Nothing
                            InsrtQry = ""
                    End If
        Else
            MBMNumber = !MBMNumber
            IsFirst = True
            SumVal = 0
            SumVal = SumVal + CDbl(!CLMTotalActual)
                If SumVal > !Amount Then
                    If IsFirst Then
                        SumVal = SumVal - CDbl(!CLMTotalActual)
                        SumVal1 = !Amount - SumVal
                        Set UpdateRS = Nothing
                        Set UpdateRS = New ADODB.Recordset
                        UpdateRS.Open "Tbl_tmpMixedCase", MediConnCISDB, adOpenDynamic, adLockOptimistic
                        UpdateRS.AddNew
                        UpdateRS!CasNumber = CASNum
                        UpdateRS!PendingAmount = SumVal
                        UpdateRS!CalValue = SumVal1
                        UpdateRS.update
                        UpdateRS.Close
                        SumVal = SumVal + CDbl(!CLMTotalActual)
                        IsFirst = False
                    Else
                        InsrtQry = "Update CISWorksheet..CLaim  set CLMBTBGAmt = (CLMTotalActual * 0.2) ,CLMFeldaFullAmt =0,CLMFeldaASO80=(CLMTotalActual * 0.8),CLMFeldaASO20=(CLMTotalActual * 0.2)  where CASNumber ='" + CASNum + "' "
                        Set UpdateRS = Nothing
                        Set UpdateRS = New ADODB.Recordset
                        UpdateRS.Open InsrtQry, MediConnCISWS, adOpenDynamic, adLockOptimistic
                        SumVal = SumVal + CDbl(!CLMTotalActual)
                    End If
                Else
                    InsrtQry = ""
                    InsrtQry = "Update CISWorksheet..CLaim  set CLMFeldaFullAmt = " + CStr(!CLMTotalActual) + "  where CASNumber ='" + CASNum + "' "
                    Set UpdateRS = Nothing
                    Set UpdateRS = New ADODB.Recordset
                    UpdateRS.Open InsrtQry, MediConnCISWS, adOpenDynamic, adLockOptimistic
                    Set UpdateRS = Nothing
                    InsrtQry = ""
                End If
        End If
    End With
    GetRs.MoveNext
Loop
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdSave_Click()
Const METHOD_NAME = "cmdSave_Click"
On Error GoTo debugs

On Error Resume Next
    Dim updatepass As String
    Dim USRUDP As New ADODB.Recordset
    
    If txtUserName = "" Then
        MsgBox "Please Enter User Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtUserName.SetFocus
    ElseIf txtPassword = "" Then
        MsgBox "Please Enter Password", vbOKOnly + vbCritical, DialogBoxTitle
        txtPassword.SetFocus
    ElseIf txtnewpass = "" Then
        MsgBox "Please Enter Valid New Password", vbOKOnly + vbCritical, DialogBoxTitle
        txtnewpass.SetFocus
    ElseIf txtnewconfirm = "" Then
        MsgBox "Please Enter Valid Re-Type Password", vbOKOnly + vbCritical, DialogBoxTitle
        txtnewconfirm.SetFocus
    ElseIf txtnewpass <> txtnewconfirm Then
        MsgBox "Password Verification Is In Valid", vbOKOnly + vbCritical, DialogBoxTitle
        txtnewpass = ""
        txtnewconfirm = ""
        txtnewpass.SetFocus
    Else
        updatepass = "Select * From USR Where USRPassword = '" & txtPassword & "' And USRCode = '" & txtUserName & "'"
        USRUDP.Open updatepass, medixmasterconn, adOpenKeyset, adLockOptimistic
            If USRUDP.RecordCount = 0 Then
                MsgBox "No Record Match For User ID or Password", vbOKOnly + vbCritical, DialogBoxTitle
                txtUserName.SetFocus
            Else
                USRUDP!USRPassword = UCase(txtnewpass)
                USRUDP.update
                MsgBox "Password Update Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
            End If
        'MsgBox "Password Update Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
        txtPassword = ""
        Me.Height = 2000
        txtPassword.SetFocus
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub Form_Activate()
    txtUserName.SetFocus
End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "frmLogin")
    Dim sBuffer As String
    Dim lSize As Long
    
    Me.Caption = DialogBoxTitle
    Me.Height = 2000
   
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOK_Click
    End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPassword.SetFocus
    End If
End Sub

Sub Create_File()
    Dim fso, txtfile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtfile = fso.CreateTextFile("LoginPath", True)
    txtfile.Write ("This is a test. ") ' Write a line.
    ' Write a line with a newline character.
    txtfile.WriteLine ("Testing 1, 2, 3.")
    ' Write three newline characters to the file.
    txtfile.WriteBlankLines (3)
    txtfile.Close
End Sub


Private Sub OutputToTxtfile()
'On Error Resume Next
Dim MyUser, MyDate, MyTime
Dim i As Integer
   
   'If Dir("\\svr-super\Medix-System\HIS-Development\FileLogin\Login.txt") = "Login.txt" Then
   If Dir(DocPath & "FileLogin\LoginCIS.txt") = "LoginCIS.txt" Then
        Open DocPath & "FileLogin\LoginCIS.txt" For Append As #1
        'Open "\\svr-super\Medix-HIS\HIS-Development\FileLogin\Login.txt" For Append As #1
        MyUser = txtUserName: MyDate = Date: MyTime = Format(Now, "hh:mm:ss am/pm")
        Write #1, "Login";
        Write #1, MyUser;
        Write #1, MyDate;
        Write #1, MyTime;
        Write #1,
        Close #1
    Else
    Call Create_File
    Open DocPath & "FileLogin\LoginCIS.txt" For Output As #1
    'Open "\\svr-super\Medix-System\HIS-Development\FileLogin\Login.txt" For Output As #1
        Write #1, "Login";
        Write #1, txtUserName;
        Write #1, Date;
        Write #1, Format(Now, "hh:mm:ss am/pm");
        Write #1,
        Close #1
    End If
End Sub

Public Function Crypt(strOriginal) As String
Dim valString$, X%

    For X = 1 To Len(strOriginal)
        'If encrypting = True Then
            valString = Asc(Mid(strOriginal, X, 1)) + (X ^ 2)
                Do While valString > 255
                    valString = valString - 255
                Loop
        'Else
        '    valString = Asc(Mid(strOriginal, x, 1)) - (x ^ 2)
        '        Do While valString <= 0
        '            valString = valString + 255
        '        Loop
        'End If
        Crypt = Crypt & Chr(valString)
    Next
    tmpEncryptDeCryptPass = Crypt
End Function



