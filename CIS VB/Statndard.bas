Attribute VB_Name = "Statndard"
Option Explicit
Public Const VK_CAPITAL = &H14
Public Type KeyboardBytes
   kbByte(0 To 255) As Byte
End Type
Public kbArray As KeyboardBytes
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long
Public Declare Function GetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
Public Declare Function SetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
Public Declare Function MakeSureDirectoryPathExists Lib "IMAGEHLP.DLL" (ByVal DirPath As String) As Long
'***************************************************************
'Windows API/Global Declarations for :strGetMachineName
'***************************************************************
Private Const MAX_LEN = 260
'-- data types for registry entries
Const REG_NONE As Long = 0
Const REG_SZ As Long = 1
Const REG_EXPAND_SZ As Long = 2
Const REG_BINARY As Long = 3
Const REG_DWORD As Long = 4
Const REG_DWORD_LITTLE_ENDIAN As Long = 4
Const REG_DWORD_BIG_ENDIAN As Long = 5
Const REG_LINK As Long = 6
Const REG_MULTI_SZ As Long = 7
Const REG_RESOURCE_LIST As Long = 8
Private Const ERROR_SUCCESS As Long = 0&
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
    '-- top level registry hives
Const HKEY_CLASSES_ROOT As Long = &H80000000
Const HKEY_CURRENT_CONFIG As Long = &H80000005
Const HKEY_CURRENT_USER As Long = &H80000001
Const HKEY_LOCAL_MACHINE As Long = &H80000002
Const HKEY_PERFORMANCE_DATA As Long = &H80000004
Const HKEY_USERS As Long = &H80000003


Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKEY As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long


Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Const SC_CLOSE = &HF060&
    Public Const MF_BYCOMMAND = &H0&
                        

Const lpnLength As Integer = 255

Public Function GetUserName()

On Error Resume Next
      ' Buffer size for the return string.
'      Const lpnLength As Integer = 255
      ' Get return buffer space.
'      Dim status As Integer
      ' For getting user information.
'      Dim lpName As String
'      Dim lpUserName As String

      ' Assign the buffer size constant to lpUserName.
'      lpUserName = Space$(lpnLength + 1)

      ' Get the log-on name of the person using product.
'      status = WNetGetUser(lpName, lpUserName, lpnLength)

'      If status = NoError Then
'         lpUserName = Left$(lpUserName, InStr(lpUserName, Chr(0)) - 1)
'         GetUserName = lpUserName
'      Else
'         MsgBox "Unable to get the name."
'         End
'      End If

      
   End Function


Public Function ClearForm(frmCall)
On Error Resume Next
Dim ctl As Control

For Each ctl In frmCall.Controls
    If TypeOf ctl Is TextBox Then
        ctl.Text = ""
    ElseIf TypeOf ctl Is ListBox Then
            ctl.Clear
    ElseIf TypeOf ctl Is ComboBox Then
            ctl.Text = ""
    'ElseIf TypeOf ctl Is OptionButton Then
            'ctl.value = False
    ElseIf TypeOf ctl Is checkbox Then
            ctl.Value = False
    'ElseIf TypeOf ctl Is Label Then
            'ctl.Text = ""
    ElseIf TypeOf ctl Is MaskEdBox Then
        ctl.Text = "__/__/____"
    'ElseIf TypeOf ctl Is ListView Then
    '        ctl.ListItems.Clear
    End If
Next ctl

'On Error GoTo 0

End Function

Public Sub CenterForm(frm As Form)
    frm.Top = (Screen.Height * 0.85) / 2 - frm.Height / 2
    frm.Left = Screen.Width / 2 - frm.Width / 2
End Sub

Public Function strGetMachineName()

    Dim strCurrentMachine As String
    Dim REGISTRYPATH As String
    
    REGISTRYPATH = "System\CONTROLSET001\Control\COMPUTERNAME\COMPUTERNAME"
    strGetMachineName = GetRegistryValue(HKEY_LOCAL_MACHINE, REGISTRYPATH, "ComputerName")
End Function


'****************************************************************
'
'GetRegistryValue
'****************************************************************
'
Public Function GetRegistryValue(xlngTopKey As Long, xstrSubKey As String, xstrValueName As String) As String
    On Error Resume Next
    Const constProcName = "GetRegistryValue"
    Dim lngRtn As Long
    Dim lngKeyHwnd As Long
    Dim strRtnValue As String
    Dim abytValue() As Byte
    Dim lngRtnValue As Long
    '-- Create under the passed, top-level subkey
    lngRtn = RegOpenKey(xlngTopKey, xstrSubKey, lngKeyHwnd)


    If lngRtn <> ERROR_SUCCESS Then
        Exit Function
    Else
        '-- lngKeyHwnd is handle to the open key
    End If

    '-- Set the variables to pass
    strRtnValue = String(MAX_LEN, Chr$(0))
    lngRtnValue = Len(strRtnValue)
    '-- Set the byte array
    abytValue = strRtnValue
    '-- Get the value
    lngRtn = RegQueryValueEx(lngKeyHwnd, _
    xstrValueName, _
    0&, REG_SZ, _
    abytValue(0), _
    lngRtnValue)
    '-- Close the handle to the subkey
    lngRtn = RegCloseKey(lngKeyHwnd)
    '-- Convert byte array into string
    strRtnValue = abytValue
    '-- Convert to UniCode, and trim based on return length
    strRtnValue = Mid$(StrConv(strRtnValue, vbUnicode), 1, lngRtnValue)
    '-- Trim any trailing null
    lngRtn = InStr(strRtnValue, Chr$(0))


    If lngRtn > 0 Then
        strRtnValue = Mid$(strRtnValue, 1, lngRtn - 1)
    End If

    '-- Set the return value
    GetRegistryValue = strRtnValue
    On Error GoTo 0
End Function


Public Function ReadINI(Section, KeyName, Filename As String) As String

    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), Filename))
End Function

'*******************************************************
'* Procedure Name: WriteINI*
'*=====================================================*
'*Writes a string to an INI file. To use, call the *
'*function and pass it the sSection, sKeyName, the New *
'*String and the INI File Name,*
'*[Ret=WriteINI(Section,Key,String,INIFile)]. *
'*Returns a 1 if there were no errors and *
'*a 0 if there were errors.*
'*******************************************************

Public Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function

Public Sub EnableCmdButton(frm As Form)
    frm.cmdAdd.Enabled = True
    frm.CmdEdit.Enabled = True
    frm.cmdDelete.Enabled = True
End Sub

Public Sub DisableCmdButton(frm As Form)
        frm.cmdAdd.Enabled = False
        frm.CmdEdit.Enabled = False
        'Frm.cmdDelete.Enabled = False
End Sub
' to set the command to become active or not
Public Sub setCmdButton(frm As Form, status As Boolean)
        frm.cmdAdd.Enabled = status
        frm.CmdEdit.Enabled = status
        'Frm.CmdDelete.Enabled = status
        frm.cmdsave.Enabled = status
End Sub
' for the text comparison
Public Function ChkStr(str As String) As String
    ChkStr = UCase(Trim(str))
End Function

Public Function RemStr(str As String) As String
    RemStr = Replace(Trim(str), "'", "''")
End Function


'Public Function IsDate(str As String) As Boolean
'Dim tmpYear As String
'    On Error GoTo errDate
'    IsDate = True
'    'If Mid(str, 7, 4) > 0 Then
'    '    str = CDate(str)
'    'End If
'   ' If IsDate$(str) = False Then
'    '    IsDate = False
'    'Else
'    tmpYear = Mid(str, 7, 4)
'        If IsNumeric(tmpYear) = False Then
'            IsDate = False
'        Else
'            If tmpYear <= 1900 Then
'                IsDate = False
'            End If
'        End If
'        If Mid(str, 4, 2) > 12 Or Mid(str, 4, 2) <= 0 Then
'            IsDate = False
'        ElseIf Mid(str, 1, 2) > 31 Or Mid(str, 3, 1) <= 0 Then
'                IsDate = False
'        End If
' Exit Function
'errDate:
'    IsDate = False
'End Function


Public Function KeyNumber(KeyAscii As Integer) As Integer
    If KeyAscii > 26 Then
       If InStr("-.0123456789", Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
    KeyNumber = KeyAscii
End Function

Public Function KeyCharacter(KeyAscii As Integer) As Integer
    Dim char
  
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
    If KeyAscii > 10 Then
       If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ ", Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
    KeyCharacter = KeyAscii
End Function


Public Sub Tool_Tip()
    'With frmChartReport
    '    .DTPicker1.ToolTipText = "Select Start Date"
    '    .DTPicker2.ToolTipText = "Select End Date"
     '   .ListView1.ToolTipText = "Display Product Result"
     '   .MSChart1.ToolTipText = "Displays Product Statistics"
     '   .DTPicker3.ToolTipText = "Select Start Date"
     '   .DTPicker4.ToolTipText = "Select End Date"
     ''   .ListView2.ToolTipText = "Display Product Result"
     '   .MSChart3.ToolTipText = "Displays Product Statistics"
    'End With
End Sub

