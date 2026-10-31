Attribute VB_Name = "ModPrint"
Option Explicit

'***************************************************************
' Simple Print Screen by Gary Choma
'
' Add this module and frmPrintScreen to your VB project to give
' your users a simple mechanism to quickly and easily print screen
' shots of your forms.  See comments in Sub "PrintScreenSnapShot"
' for more.
'***************************************************************
Private Const mcsClassName As String = "modPrintScreen"

'--- Windows API/Global Declarations for PrintScreenSnapShot
Public Declare Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long)

Public Const TheScreen = 1
Public Const TheForm = 0

Public Sub PrintScreenSnapShot()
'--- Call this from any form in your project.
'--- Example: To make "F6" key a print screen button of a form in your project,
'--- set the form's .KeyPreview property to "True" at design time, and add an
'--- event handler for the form's KeyDown event that looks something like this:
'
'    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'        'Enable PrintScreen feature
'        '--- IMPORTANT - Don't forget to set .KeyPreview = True for this form
'        '--- at design time!!
'        If Shift = 0 Then
'            Select Case KeyCode
'                Case vbKeyF6
'                    KeyCode = 0
'                    PrintScreenSnapShot
'            End Select
'        End If
'    End Sub
'
'--- NOTE: I haven't fully finished tweaking this routine for handling different
'--- types of clipboard contents. Right now, if you just have plain text or RTF
'--- text in the Windows clipboard, it will cache, and restore the contents. This
'--- is necessary because this PrintScreen features uses the clipboard to handle
'--- the PrintScreen image. Because I put in an "On Error Resume Next" the print
'--- screen should work, but the Windows clipboard contents may get lost if it
'--- contains formats other than plain or RTF text. The code to manage the
'--- clipboard data was copied from VB Help.

    Const csMethodName As String = "PrintScreenSnapShot"
    Dim frmPrint As frmPrintScreen
    Dim ClpFmt As Variant
    Dim vContents As Variant

    On Error GoTo ErrorHandler

    If PrinterIsInstalled Then
        On Error Resume Next    ' Set up error handling.
        If Clipboard.GetFormat(vbCFText) Then ClpFmt = ClpFmt + 1
        If Clipboard.GetFormat(vbCFBitmap) Then ClpFmt = ClpFmt + 2
        If Clipboard.GetFormat(vbCFDIB) Then ClpFmt = ClpFmt + 4
        If Clipboard.GetFormat(vbCFRTF) Then ClpFmt = ClpFmt + 8
        
        'On Error GoTo ErrorHandler
        
        '--- Cache current contents of clipboard:
        Select Case ClpFmt
            Case 1
                'Msg = "The Clipboard contains only text."
                vContents = Clipboard.GetText(vbCFText)
            Case 2, 4, 6
                'Msg = "The Clipboard contains only a bitmap."
                vContents = Clipboard.GetData(ClpFmt)
            Case 3, 5, 7
                'Msg = "The Clipboard contains text and a bitmap."
                vContents = Clipboard.GetData(ClpFmt)
            Case 8, 9
                'Msg = "The Clipboard contains only rich text."
                vContents = Clipboard.GetText(vbCFRTF)
            Case Else
                'Msg = "There is nothing on the Clipboard."
        End Select
        'MsgBox Msg  ' Display message.
        
        '--- Do a <Print Scrn> with an API call:
        keybd_event vbKeySnapshot, TheForm, 0&, 0&

        '--- Give Windows a chance to update the clipboard
        DoEvents
        
        Set frmPrint = New frmPrintScreen
        frmPrint.PrintBitMap Clipboard.GetData(vbCFBitmap)  ', Me.Height, Me.Width
        Set frmPrint = Nothing
        
        '--- Restore contents of clipboard
        Clipboard.Clear
        Select Case ClpFmt
            Case 1
                'Msg = "The Clipboard contains only text."
                Clipboard.SetText vContents, vbCFText
            Case 2, 4, 6
                'Msg = "The Clipboard contains only a bitmap."
                 Clipboard.SetData vContents, ClpFmt
            Case 3, 5, 7
                'Msg = "The Clipboard contains text and a bitmap."
                Clipboard.SetData vContents, ClpFmt
            Case 8, 9
                'Msg = "The Clipboard contains only rich text."
                Clipboard.SetText vContents, vbCFRTF
            Case Else
                'Msg = "There is nothing on the Clipboard."
        End Select
    End If

    Exit Sub

ErrorHandler:
    '--- Add your error handler here:
    MsgBox Err.Number & " " & Err.Source & vbCrLf & Err.Description, vbExclamation, _
        mcsClassName & "." & csMethodName
End Sub

Private Function PrinterIsInstalled() As Boolean
    Dim dummy As String

    On Error Resume Next
    dummy = Printer.DeviceName
    
    If Err.Number Then
        MsgBox "No default printer installed." & vbCrLf _
            & "To install and select a default printer, select the " _
            & "Setting / Printers command in the Start menu, and then " _
            & "double-click on the Add Printer icon.", _
            vbExclamation, "Printer Error"
        PrinterIsInstalled = False
    Else
        PrinterIsInstalled = True
    End If
End Function






Public Function PrintTxtBoxes(frmForm As Form)
        Dim sTxt As Object
    Printer.Print Tab(4), "TextBox print demo"
    Printer.Print
    For Each sTxt In frmForm.Controls
        If sTxt.Tag = "T" Then
            Printer.Print sTxt.Name, ": ", sTxt.Text
        End If
        DoEvents
    Next sTxt
    Printer.EndDoc
    'You can remove this
    MsgBox "Done!"
End Function

'Public Function PrintLV(lv As ListView)
    
    'Printer.FontSize = 12
    'Dim subit As Variant
    'Dim i As Integer
    'Dim X As Integer


    'For i = 1 To lv.listitems.count


'subit = lv.listitems(i).Text & vbTab


    'For X = 1 To Subs


'subit = subit & lv.listitems(i).SubItems(X) & vbTab
'Next
'Printer.Print subit


'subit = ""
'Next
'Printer.EndDoc
'End Function


Public Function PrintLV(LV As ListView)
    Dim i As Integer
    Dim sData As String
    Dim X As Integer
    Dim sIdx As Integer
    sIdx = LV.ColumnHeaders.Count - 1
    On Error Resume Next
    Printer.Print Tab(4), "ListView print demo"
    Printer.Print
    For i = 1 To LV.ListItems.Count
        sData = LV.ListItems.Item(i).Text
        For X = 1 To sIdx
            sData = sData & " " & LV.ListItems.Item(i).SubItems(X)
        Next
        Printer.Print sData
        sData = ""
    Next
    Printer.EndDoc
    'LV.listitems.Clear
    MsgBox "Done!!"

End Function


