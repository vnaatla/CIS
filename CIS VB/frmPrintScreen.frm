VERSION 5.00
Begin VB.Form frmPrintScreen 
   Caption         =   "frmPrintScreen"
   ClientHeight    =   3192
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3192
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPrintScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mcsClassName As String = "frmPrintScreen"
Private Sub Form_Load()
  Call ActionLogUpdate(12, "frmPrintScreen")
End Sub

Public Sub PrintBitMap(pbmpBitMap As Variant) ', ByVal pintHeight As Integer, ByVal pintWidth As Integer)
    Const csMethodName As String = "PrintBitMap"

    On Error GoTo ErrorHandler

    With Me
        .Visible = False
        .Height = Screen.Height + 300 'pintHeight + 120
        .Width = Screen.Width + 150 'pintWidth + 120
        .Picture = pbmpBitMap
        .PrintForm
    End With
    Unload Me

    Exit Sub

ErrorHandler:
    '--- Add your error handler here:
    MsgBox Err.Number & " " & Err.Source & vbCrLf & Err.Description, vbExclamation, mcsClassName & "." & csMethodName
End Sub




