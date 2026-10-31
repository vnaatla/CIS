Attribute VB_Name = "modsubclass"
Option Explicit

' Holds a reference to the previous window
' procedure before we started subclassing
Public OldProc As Long

Declare Function SetWindowLong Lib "user32" _
  Alias "SetWindowLongA" _
 (ByVal hwnd As Long, ByVal nIndex As Long, _
  ByVal dwNewLong As Long) As Long
  
Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" _
 (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Const GWL_WNDPROC = (-4)

Declare Function CallWindowProc Lib "user32" _
  Alias "CallWindowProcA" _
 (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
  ByVal Msg As Long, _
  ByVal wParam As Long, ByVal lParam As Long) As Long

' The message that we are going to monitor
Public Const WM_CONTEXTMENU = &H7B&

Public Function WndProc(ByVal hwnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As Long) As Long

If wMsg = WM_CONTEXTMENU Then
  '// Handle this message and call our
  '// popup routine
  WndProc = 0
  'Frmtest.ShowPopup
  Exit Function
End If

' Pass on all the other unhandled messages
WndProc = CallWindowProc(OldProc, hwnd, wMsg, wParam, lParam)

End Function






