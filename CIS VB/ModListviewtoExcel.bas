Attribute VB_Name = "ModListViewtoExcel"


' =============== Invoice Module Report ====================
Public Function ExportToExcelInvoice(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Worksheet\INVOIC~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                '    Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                '       ' If tag is empty, format as text
                '        Case "string", ""
                '        .NumberFormat = "@"
                '        .HorizontalAlignment = xlLeft
                '        Case "number"
                '        .NumberFormat = "#,##0.00_);(#,##0.00)"
                '        .HorizontalAlignment = xlRight
                '        Case "date"
                '        .NumberFormat = "dd/mm/yyyy"
                '        .HorizontalAlignment = xlCenter 'xlRight
                '    End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelInvoice = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function
' =============== Invoice Module Report ====================

' =============== Worksheet Module Report ====================

Public Function ExportToExcelWksToAC(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Worksheet\SendtoAC.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelWksToAC = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelWorksheet(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Worksheet\WO1C64~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                    '   ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter 'xlRight
                    'End Select
                End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelWorksheet = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function
' =============== Worksheet Module Report ====================

Public Function ExportToExcelClaimAnalysis(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Worksheet\CLAIMA~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                '    Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                '       ' If tag is empty, format as text
                '        Case "string", ""
                '        .NumberFormat = "@"
                '        .HorizontalAlignment = xlLeft
                '        Case "number"
                '        .NumberFormat = "#,##0.00_);(#,##0.00)"
                '        .HorizontalAlignment = xlRight
                '        Case "date"
                '        .NumberFormat = "dd/mm/yyyy"
                '        .HorizontalAlignment = xlCenter 'xlRight
                '    End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelClaimAnalysis = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis1(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Illness.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis1 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis111(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Durati~2.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis111 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis112(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Durati~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis112 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis15(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\General Listing.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis15 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis31(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Ageban~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis31 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis32(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Ageban~2.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis32 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis33(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Ageban~3.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis33 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis34(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Ageban~4.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis34 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis35(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\AG33D4~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis35 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis36(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\AG938F~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis36 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis41(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Agentn~4.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis41 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis42(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Agentn~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis42 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis43(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Agentn~3.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis43 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis44(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Agentn~2.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis44 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis2(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Genera~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis2 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis211(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Planno~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis211 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis212(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Planno~2.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis212 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis213(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Planno~3.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis213 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis214(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Planno~4.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis214 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis215(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\PL133C~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis215 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis26(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\CLMCHK~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "B") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis26 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis61(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Gender~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis61 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis62(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\GE841A~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis62 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis63(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\PL133C~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis63 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis64(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Gender~3.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis64 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis65(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Gender~4.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis65 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis66(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\GE053E~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis66 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis151(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Diagno~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis151 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis152(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Diagno~2.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis152 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis153(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Diagno~3.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis153 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis8(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\State.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis8 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelAnalysis12(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\ClaimTracking\Doctor.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelAnalysis12 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelPayment(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Accoun~1\Paymen~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelPayment = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelReceipt(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Accoun~1\Receip~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelReceipt = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelReceiptMBM(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Accoun~1\Receip~2.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelReceiptMBM = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelReceiptPayorBor(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Accoun~1\Receip~3.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelReceiptPayorBor = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelReceiptPayorWks(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Accoun~1\Receip~4.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelReceiptPayorWks = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelReceiptMNRBBor(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Accoun~1\Re8016~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelReceiptMNRBBor = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelReceiptMNRBWks(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Accoun~1\ReB004~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelReceiptMNRBWks = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelError(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Error.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                '    Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                '       ' If tag is empty, format as text
                '        Case "string", ""
                '        .NumberFormat = "@"
                '        .HorizontalAlignment = xlLeft
                '        Case "number"
                '        .NumberFormat = "#,##0.00_);(#,##0.00)"
                '        .HorizontalAlignment = xlRight
                '        Case "date"
                '        .NumberFormat = "dd/mm/yyyy"
                '        .HorizontalAlignment = xlCenter 'xlRight
                '    End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems
    
    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "C") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelError = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

'============= Case Module Report =======================================

Public Function ExportToExcelCaseTrack1(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\CaseTracking\Genera~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                '    Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                '       ' If tag is empty, format as text
                '        Case "string", ""
                '        .NumberFormat = "@"
                '        .HorizontalAlignment = xlLeft
                '        Case "number"
                '        .NumberFormat = "#,##0.00_);(#,##0.00)"
                '        .HorizontalAlignment = xlRight
                '        Case "date"
                '        .NumberFormat = "dd/mm/yyyy"
                '        .HorizontalAlignment = xlCenter 'xlRight
                '    End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems
    
    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelCaseTrack1 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelCaseTrack2(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Worksheet\CASETR~1.XLT")
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\CaseTracking\ADMISS~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                '    Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                '       ' If tag is empty, format as text
                '        Case "string", ""
                '        .NumberFormat = "@"
                '        .HorizontalAlignment = xlLeft
                '        Case "number"
                '        .NumberFormat = "#,##0.00_);(#,##0.00)"
                '        .HorizontalAlignment = xlRight
                '        Case "date"
                '        .NumberFormat = "dd/mm/yyyy"
                '        .HorizontalAlignment = xlCenter 'xlRight
                '    End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems
    
    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelCaseTrack2 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelCaseTrack3(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Worksheet\CASETR~1.XLT")
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\CaseTracking\DISCHA~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                '    Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                '       ' If tag is empty, format as text
                '        Case "string", ""
                '        .NumberFormat = "@"
                '        .HorizontalAlignment = xlLeft
                '        Case "number"
                '        .NumberFormat = "#,##0.00_);(#,##0.00)"
                '        .HorizontalAlignment = xlRight
                '        Case "date"
                '        .NumberFormat = "dd/mm/yyyy"
                '        .HorizontalAlignment = xlCenter 'xlRight
                '    End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems
    
    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelCaseTrack3 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelCaseTrack4(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Worksheet\CASETR~1.XLT")
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\CaseTracking\NoBrea~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                '    Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                '       ' If tag is empty, format as text
                '        Case "string", ""
                '        .NumberFormat = "@"
                '        .HorizontalAlignment = xlLeft
                '        Case "number"
                '        .NumberFormat = "#,##0.00_);(#,##0.00)"
                '        .HorizontalAlignment = xlRight
                '        Case "date"
                '        .NumberFormat = "dd/mm/yyyy"
                '        .HorizontalAlignment = xlCenter 'xlRight
                '    End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems
    
    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelCaseTrack4 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function


Public Function ExportToExcelCaseTrack5(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Worksheet\CASETR~1.XLT")
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\CaseTracking\BreakD~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                '    Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                '       ' If tag is empty, format as text
                '        Case "string", ""
                '        .NumberFormat = "@"
                '        .HorizontalAlignment = xlLeft
                '        Case "number"
                '        .NumberFormat = "#,##0.00_);(#,##0.00)"
                '        .HorizontalAlignment = xlRight
                '        Case "date"
                '        .NumberFormat = "dd/mm/yyyy"
                '        .HorizontalAlignment = xlCenter 'xlRight
                '    End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems
    
    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelCaseTrack5 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelLG(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\LG\LGTRAC~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                '    Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                '       ' If tag is empty, format as text
                '        Case "string", ""
                '        .NumberFormat = "@"
                '        .HorizontalAlignment = xlLeft
                '        Case "number"
                '        .NumberFormat = "#,##0.00_);(#,##0.00)"
                '        .HorizontalAlignment = xlRight
                '        Case "date"
                 '       .NumberFormat = "dd/mm/yyyy"
                '        .HorizontalAlignment = xlCenter 'xlRight
                '    End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelLG = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

'============= Case Module Report =======================================

'========================= MBM TRACKING REPORT ==================================
Public Function ExportToExcelMBMTracking1T1(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\GENERA~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "S") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "U") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "V") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking1T1 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking1T2(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\GENERA~2.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "R") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "T") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "U") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking1T2 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function


Public Function ExportToExcelMBMTracking1T3(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Label.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "C") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking1T3 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking2(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\INSTYP~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "D") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "H") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking2 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking3(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\PLANNO~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "H") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "I") = FrmMBMTracking.LblAccess
    
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking3 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking41(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\AgentDetail.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "M") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "N") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "O") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking41 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking42(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\AgentG~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "C") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "D") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking42 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking43(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Agent_~3.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "D") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "H") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking43 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking44(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Agent_~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "H") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "I") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking44 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking45(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Agent_~2.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "H") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "I") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking45 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking46(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Agentb~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "C") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "D") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking46 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking51(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Ageban~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "H") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "I") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking51 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking52(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Ageban~2.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "D") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "H") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking52 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking53(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Ageban~3.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "D") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "H") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking53 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking6(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Gender.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "H") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "I") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking6 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking7(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Race.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "H") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "I") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking7 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking8(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\States.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "D") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "H") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking8 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking9(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Payor.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "C") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "D") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking9 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking10(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\MCOLis~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "I") = FrmMBMTracking.LblMCO
            
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking10 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking11(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\MCOSum~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "D") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.LblMCO
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking11 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking12(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Takeov~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "C") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "D") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking12 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking131(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Relati~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.lbltotalSuppCount
            
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking131 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking132(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Relati~2.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.lbltotalSuppCount
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking132 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking141(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Effper~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
            
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking141 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking142(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Effper~2.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking142 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking151(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Branch.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "D") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "H") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking151 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking152(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Branch~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
    
    objWorksheet.Cells(10 + intCounter, "D") = FrmMBMTracking.lblTotalPrincipalCount
    objWorksheet.Cells(10 + intCounter, "E") = FrmMBMTracking.lbltotalSuppCount
    objWorksheet.Cells(10 + intCounter, "F") = FrmMBMTracking.LblMCO
    objWorksheet.Cells(10 + intCounter, "G") = FrmMBMTracking.LblPrem
    objWorksheet.Cells(10 + intCounter, "H") = FrmMBMTracking.LblAccess
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking152 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking18(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Date.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking18 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelMBMTracking191(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMTracking\Genera~3.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                        'Case "string", ""
                        '.NumberFormat = "@"
                        
                        '.HorizontalAlignment = xlLeft
                        'Case "number"
                        '.NumberFormat = "#,##0.00_);(#,##0.00)"
                        '.HorizontalAlignment = xlRight
                        'Case "Numeric"
                        '.NumberFormat = "dd/mm/yyyy"
                        '.HorizontalAlignment = xlCenter 'xlRight
                        '.NumberFormat = "Integer"
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelMBMTracking191 = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function
'========================= MBM TRACKING REPORT ==================================

'========================= BORDX ==================================

Public Function ExportToExcelBordxTracking(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Bordx\Bordxt~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelBordxTracking = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function



Public Function ExportToExcelNonSihat(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    

    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Bordx\outsta~2.xlt")
  
    Set objWorksheet = objWorkbook.Sheets(1)
    
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(9)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                '    Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                '       ' If tag is empty, format as text
                '        Case "string", ""
                '        .NumberFormat = "@"
                '        .HorizontalAlignment = xlLeft
                '        Case "number"
                '        .NumberFormat = "#,##0.00_);(#,##0.00)"
                '        .HorizontalAlignment = xlRight
                '        Case "date"
                '        .NumberFormat = "dd/mm/yyyy"
                '        .HorizontalAlignment = xlRight
                '    End Select
                'End With
            End If
        Next i
        ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
    End If
   
intCounter = 0
intStartRow = 10

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems
    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range

With objWorksheet
    If lvw Is Nothing = False Then
        .Range(.Cells(10, 1), _
        .Cells(9 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(3, "D") = FrmOSBordx.txtOSDate
        objWorksheet.Cells(4, "A") = FrmOSBordx.Label3
        objWorksheet.Cells(8, "A") = FrmOSBordx.Label5
    End If
    
    objWorksheet.Cells(17 + intCounter, "A") = "Total Due"
    objWorksheet.Cells(17 + intCounter, "D") = FrmOSBordx.LblTotalDue
    objWorksheet.Cells(20 + intCounter, "A") = "Prepared by :"
    objWorksheet.Cells(20 + intCounter, "B") = Logon
    objWorksheet.Cells(21 + intCounter, "A") = "Date Prepared :"
    objWorksheet.Cells(21 + intCounter, "B") = Format(Date, "dd/mm/yyyy")
    
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelNonSihat = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function


Public Function ExportToExcelBordxToPayor(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Double
    Dim intCounter As Double
    Dim intStartRow As Double
    Dim strArray() As String
    Dim intVisibleColumns() As Double
    Dim intColumns As Double
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Bordx\Sendto~1.XLT")
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Bordx\Sendto~2.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelBordxToPayor = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelBordxReturnPayor(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Bordx\Return~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelBordxReturnPayor = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelBordxToMNRB(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Bordx\Sendto~2.XLT")
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Bordx\Sendto~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelBordxToMNRB = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelBordxToAC(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Bordx\SendtoAC.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelBordxToAC = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelBordxACPeriod(lvw As MSComctlLib.ListView) As Boolean
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Bordx\ACPeriod.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelBordxACPeriod = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelBordxControlSihat(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Bordx\BORDXC~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                '    Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                '       ' If tag is empty, format as text
                '        Case "string", ""
                '        .NumberFormat = "@"
                '        .HorizontalAlignment = xlLeft
                '        Case "number"
                '        .NumberFormat = "#,##0.00_);(#,##0.00)"
                '        .HorizontalAlignment = xlRight
                '        Case "date"
                '        .NumberFormat = "dd/mm/yyyy"
                '        .HorizontalAlignment = xlCenter 'xlRight
                '    End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelBordxControlSihat = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function


Public Function ExportToExcelBordxControlNonSihat(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Bordx\BORDXC~2.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                '    Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                '       ' If tag is empty, format as text
                '        Case "string", ""
                '        .NumberFormat = "@"
                '        .HorizontalAlignment = xlLeft
                '        Case "number"
                '        .NumberFormat = "#,##0.00_);(#,##0.00)"
                '        .HorizontalAlignment = xlRight
                '        Case "date"
                '        .NumberFormat = "dd/mm/yyyy"
                '        .HorizontalAlignment = xlCenter 'xlRight
                '    End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelBordxControlNonSihat = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Function ExportToExcelBordx(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Bordx\BORDXN~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                '    Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                '        Case "string", ""
                '        .NumberFormat = "@"
                '        .HorizontalAlignment = xlLeft
                '        Case "number"
                '        .NumberFormat = "#,##0.00_);(#,##0.00)"
                '        .HorizontalAlignment = xlRight
                '        Case "date"
                '        .NumberFormat = "dd/mm/yyyy"
                '        .HorizontalAlignment = xlCenter 'xlRight
                '    End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelBordx = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function
'========================= BORDX ==================================

'========================= DrCrListing ==================================
Public Function ExportToExcelDrCrListing(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
  
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
        
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\DRCRListing\DRCRListing.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    
        
    intCounter = 0
        
    ' Title 1
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    objRange.Font.Color = vbBlack
    'Set objRange = objWorksheet.Cells.NumberFormat
        'objRange.Worksheet = General
        'objRange.Cells.NumberFormat = General
    'objRange.Cells.NumberFormat
    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
                'With objWorksheet.Columns(intColumns)
                    'Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                       ' If tag is empty, format as text
                    '    Case "string", ""
                    '    .NumberFormat = "@"
                    '    .HorizontalAlignment = xlLeft
                    '    Case "number"
                    '    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    '    .HorizontalAlignment = xlRight
                    '    .NumberFormat = "Integer"
                    '    Case "date"
                    '    .NumberFormat = "dd/mm/yyyy"
                    '    .HorizontalAlignment = xlCenter
                        
                    'End Select
                'End With
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
intCounter = 0
intStartRow = 5

If lvw Is Nothing = False Then
For Each itm In lvw.ListItems

    ' A response of vbYes meant to export all the items

    If lngResults = vbYes Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If i = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm
End If

' Send entire array to Excel range


With objWorksheet
    If lvw Is Nothing = False And intCounter <> 0 Then
        .Range(.Cells(5, 1), _
        .Cells(4 + intCounter, intColumns)) = strArray
    End If
End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
objWorksheet.Columns.AutoFit
objExcel.Interactive = True

ExportToExcelDrCrListing = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Public Sub ExportListViewtoExcel(lvwList As Control)
Dim vntHeader As Variant
Dim vntData As Variant
Dim X As Long
Dim Y As Long
Dim intCol As Double
Dim lngRow As Long
    
    'Get Counts
    'intCol = CInt(lvwList.ColumnHeaders.Count - 1)
    lngRow = CLng(lvwList.ListItems.Count - 1)
    
    For X = 0 To CInt(lvwList.ColumnHeaders.Count - 1)
        If lvwList.ColumnHeaders(X + 1).Text <> "" Then
            intCol = intCol + 1
        Else
            Exit For
        End If
    Next
    
    ReDim vntHeader(0)
    ReDim vntData(intCol - 1, lngRow)
    
    'Create Header Array


    For X = 0 To intCol - 1
        ReDim Preserve vntHeader(X)
        vntHeader(X) = lvwList.ColumnHeaders(X + 1).Text
    Next
    
    'Create Data Array


    For X = 0 To lngRow
        vntData(0, X) = lvwList.ListItems.Item(X + 1).Text
        For Y = 1 To intCol - 1
            vntData(Y, X) = lvwList.ListItems.Item(X + 1).SubItems(Y)
        Next
    Next
    
    'Create Excel Object
    OpenExcel vntData, vntHeader
    
End Sub


Private Sub ExportRecords(vntData As Variant, vntHeader As Variant, ws As Worksheet)
Dim lngRow As Long
Dim intCol As Double
Dim varData As Variant
Dim intStart As Double
Dim X As Double
Dim Y As Double
    'Select all Cells and and set the number


        'format to string
        ws.Cells.Select
        ws.Cells.NumberFormat = "@"
        ws.Cells(1, 1).Select
        lngRow = UBound(vntData, 2) + 2
        intCol = UBound(vntData, 1) + 1
        intStart = 2 'Start from line 2
        'Freeze Row 2
        ws.Rows(2).Select
        ws.Activate
        'ActiveWindow..FreezePanes = True
        'Add Headers


        For X = 1 To intCol
            varData = vntHeader(X - 1)
            ws.Cells(1, X) = CStr(varData)
            ws.Cells(1, X).Font.Bold = True
        Next
        
        'Add Data


        For Y = 1 To intCol


            For X = intStart To lngRow
                varData = vntData(Y - 1, X - 2)


                If IsNull(varData) Then 'Make sure no null values, Excel will choke
                    'Add 1 to Move down a column
                    ws.Cells(X + 1, Y) = ""
                Else
                    ws.Cells(X + 1, Y) = CStr(varData) 'Convert to String to preserve formatting
                End If
            Next
        Next
        'Resize Columns to Fit
        ws.Columns.AutoFit
    End Sub


Private Sub OpenExcel(vntData As Variant, vntHeader As Variant)
On Error GoTo Err_OpenExcel
Dim objExcel As Excel.Application
Dim objWrkSht As Worksheet
'Dim X As Integer
    
    'Create Excel Object
    Set objExcel = CreateObject("Excel.Application")
    'Add the Workbook
    objExcel.Workbooks.Add
        
    Set objWrkSht = objExcel.ActiveWorkbook.Sheets(1)
    objExcel.Visible = True
    'Fill the Workbook with data
    ExportRecords vntData, vntHeader, objWrkSht
    objExcel.Interactive = True
    ' Clean up:
    'Set objExlSht = Nothing
    Set objWrkSht = Nothing
    Set objExcel = Nothing
Err_OpenExcel:


    Select Case Err
        Case 0
        Case 439
        MsgBox "You must have Microsoft Excel installed on your PC.", vbCritical, "Application Not Found"
        Case Else
        MsgBox Err & ": " & error, vbCritical, "OpenExcel Error"
    End Select
End Sub

Public Sub ExporttoExcelPosting(lvwList As Control, ServerPath As String, Filename As String)
Dim vntHeader As Variant
Dim vntData As Variant
Dim X As Long
Dim Y As Long
Dim intCol As Integer
Dim lngRow As Long
    
    'Get Counts
    'intCol = CInt(lvwList.ColumnHeaders.Count - 1)
    lngRow = CLng(lvwList.ListItems.Count - 1)
    
    For X = 0 To CInt(lvwList.ColumnHeaders.Count - 1)
        If lvwList.ColumnHeaders(X + 1).Text <> "" Then
            intCol = intCol + 1
        Else
            Exit For
        End If
    Next
    
    ReDim vntHeader(0)
    ReDim vntData(intCol - 1, lngRow)
    
    'Create Header Array


    For X = 0 To intCol - 1
        ReDim Preserve vntHeader(X)
        vntHeader(X) = lvwList.ColumnHeaders(X + 1).Text
    Next
    
    'Create Data Array


    For X = 0 To lngRow
        vntData(0, X) = lvwList.ListItems.Item(X + 1).Text
        For Y = 1 To intCol - 1
            vntData(Y, X) = lvwList.ListItems.Item(X + 1).SubItems(Y)
        Next
    Next
    
    'Create Excel Object
    OpenExcelPosting vntData, vntHeader, ServerPath, Filename
    
End Sub

Private Sub OpenExcelPosting(vntData As Variant, vntHeader As Variant, ServerPath As String, Filename As String)
On Error GoTo Err_OpenExcel
Dim objExcel As Excel.Application
Dim objWrkSht As Worksheet
'Dim X As Integer
    'Set fso = CreateObject("Scripting.FileSystemObject")
    'Set txtSVRfile = fso.CreateTextFile((Trim(txtSVRPath) & Trim(txtFilName & ".xls")), True)
    'Set txtfile = fso.CreateTextFile((Trim(txtPath) & Trim(txtFilName & ".xls")), True)
    'txtfile = (Trim(txtPath) & Trim(txtFilName & ".xls"))
    'txtSVRfile = (Trim(txtSVRPath) & Trim(txtFilName & ".xls"))
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtSVRfile = fso.CreateTextFile((Trim(ServerPath) & Trim(Filename & ".xls")), True)
    'Set txtfile = fso.CreateTextFile((Trim(ServerPath) & Trim(txtFilName & ".xls")), True)
    'txtfile = (Trim(txtPath) & Trim(txtFilName & ".xls"))
    txtSVRfile = (Trim(ServerPath) & Trim(Filename & ".xls"))
        
   ' Set objExcel = CreateObject("Excel.Application")
    ''Add the Workbook
    'objExcel.Workbooks.Add
   '
   ' Set objWrkSht = objExcel.ActiveWorkbook.Sheets(1)
   ' objExcel.Visible = True
        
    'Create Excel Object
    Set objExcel = CreateObject("Excel.Application")
    Set xlBook = objExcel.Workbooks.Open(txtSVRfile)
        'Set xlBook = objExcel.Workbooks.Open(Trim(ServerPath) & Trim(Filename & ".xls"))
        Set objWrkSht = xlBook.Sheets(1)
    
    'Add the Workbook
    'objExcel.Workbooks.Add
    'Set xlSheet = xlBook.Sheets(1)
    'Set objWrkSht = objExcel.ActiveWorkbook.Sheets(1)
    'xlBook.Visible = True
    'Fill the Workbook with data
    ExportRecords vntData, vntHeader, objWrkSht
    xlBook.Save
    'xlBook.SaveAs (Trim(ServerPath) & Trim(Filename & ".xls"))
    
    'objExcel.Interactive = True
    
    
    'xlBook.SaveAs Trim(txtSVRfile)
    'objExcel.Save
    ' Clean up:
    'Set objExlSht = Nothing
    
    Set objWrkSht = Nothing
    'objExcel.ActiveWorkbook.Close (xlBook.Save, Trim(ServerPath) & Trim(Filename & ".xls"))
    objExcel.Quit
    Set objExcel = Nothing
    
    Set xlBook = Nothing
Err_OpenExcel:


    Select Case Err
        Case 0
        Case 439
        MsgBox "You must have Microsoft Excel installed on your PC.", vbCritical, "Application Not Found"
        Case Else
        MsgBox Err & ": " & error, vbCritical, "OpenExcel Error"
    End Select
End Sub

