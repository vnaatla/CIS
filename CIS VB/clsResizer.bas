Attribute VB_Name = "clsResizer"
Private mFormWidth As Long
Private mFormHeight As Variant
Private marrControlsSizes As Variant
Private miControlsCount As Integer
'Public Event ResizeAllComplete()

Public Sub GetMySizes(pForm As Form)
Dim X As Integer 'counter
On Error GoTo EH
'Set the object properties to the properties of the calling form...
mFormWidth = pForm.ScaleWidth
mFormHeight = pForm.ScaleHeight
miControlsCount = pForm.Controls.Count
'Size the array to hold the required number of controls by 6 (for the 6 properties to be stored for each ctl)
' ...these being a reference to the control itself and the top, left, width, height & font size of the control
ReDim marrControlsSizes(miControlsCount, 5)
'Populate the array...
For X = 0 To miControlsCount - 1
With pForm
marrControlsSizes(X, 0) = .Controls(X)
marrControlsSizes(X, 1) = .Controls(X).Top
marrControlsSizes(X, 2) = .Controls(X).Left
marrControlsSizes(X, 3) = .Controls(X).Width
marrControlsSizes(X, 4) = .Controls(X).Height
marrControlsSizes(X, 5) = .Controls(X).FontSize
End With
Next X

Exit Sub
EH:
If Err.Number = 438 Or Err.Number = 383 Then Resume Next
Err.Raise vbObjectError + 600, "clsResizer.GetMySizes", Err.Description
End Sub


Public Sub ResizeAll(pForm As Form)
Dim Frm As Form
Dim ctl As Control
Dim ScaleX As Double
Dim ScaleY As Double
Dim LeftArg As Double
On Error GoTo EH
'Set the object properties to the properties of the calling form...
Set Frm = pForm
'Calculate the scale
ScaleX = Frm.ScaleWidth / mFormWidth
ScaleY = Frm.ScaleHeight / mFormHeight
'Iterate through the array of stored controls and original sizes to resize the controls on the form
For X = 0 To miControlsCount - 1
Set ctl = Frm.Controls(X)
With ctl
'Set the font size...
If ScaleX < ScaleY Then
ctl.FontSize = marrControlsSizes(X, 5) * ScaleX
Else
ctl.FontSize = marrControlsSizes(X, 5) * ScaleY
End If
'The following is to enable compatibility with MS tabbed dialog control...
If marrControlsSizes(X, 2) < 0 Then
LeftArg = (marrControlsSizes(X, 2) + 75000) * ScaleX - 75000
Else
LeftArg = marrControlsSizes(X, 2) * ScaleX
End If
'=====================================================
ctl.Move LeftArg, marrControlsSizes(X, 1) * ScaleY, marrControlsSizes(X, 3) * marrControlsSizes(X, 4) * ScaleY
ctl.Width = marrControlsSizes(X, 3) * ScaleX
ctl.Height = marrControlsSizes(X, 4) * ScaleY
End With
Next X
Set Frm = Nothing
'RaiseEvent ResizeAllComplete
Exit Sub

EH:
If Err.Number = 438 Or Err.Number = 383 Then Resume Next
Err.Raise vbObjectError + 601, "clsResizer.ResizeAll", Err.Description
End Sub

