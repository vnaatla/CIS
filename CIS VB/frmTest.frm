VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Test Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CheckFolderExists()
Dim strFolderName As String
Dim strFolderExists, tmpCASnumber As String
 
    Dim sql, StrSql1 As String
    Dim ado As New ADODB.Recordset
    Dim USR As New ADODB.Recordset
        
    sql = "Select CasNumber,* from cisdb..tmp_CasFloderstatus WHERE  0=0 "
    'and Remark IS NOT NULL AND BinRemark IS NOT NULL
    ado.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do Until ado.EOF
        tmpCASnumber = IIf(Len(ado!CasNumber), ado!CasNumber, "")
        ''strFolderName = DocPath & "Outpatient\" & "Reim_ScannedDoc\" & tmpCASnumber + "\"""
        'strFolderName = DocPath & "Outpatient\" & "SP_ScannedDoc\" & tmpCASnumber + ".pdf"
        'strFolderName = DocPath & "Outpatient\" & "Reim_ScannedDoc\" & tmpCASnumber
        strFolderName = DocPath & "Outpatient\" & "Reim_ScannedDoc\" & tmpCASnumber + ".pdf"
        strFolderExists = Dir(strFolderName, vbDirectory)
        If strFolderExists = "" Then
            StrSql1 = "UPDATE cisdb..tmp_CasFloderstatus SET BinRemark = 'The selected case file Missing' WHERE casnumber = '" & tmpCASnumber & "'"
            USR.Open StrSql1, MediConnCISDB, adOpenKeyset, adLockOptimistic
        End If
    ado.MoveNext
    Loop
    ado.Close
    Set ado = Nothing
    

    
'    strFolderName = DocPath & "Outpatient\" & "Reim_ScannedDoc\" & "EA11074356\"
'    ''strFolderName = DocPath & "Outpatient\" & "Reim_ScannedDoc\EA11074356\" & EA11074356 & ""
'    strFolderExists = Dir(strFolderName, vbDirectory)
'
'    If strFolderExists = "" Then
'        MsgBox "The selected folder doesn't exist"
'    Else
'        MsgBox "The selected folder exists"
'    End If
    
End Sub

Private Sub Form_Load()
'Call CheckFolderExists
'MsgBox "Records Checked.Please check..!"

strUrl = "https://chatuat.hla.com.my/HLAApiWS/HlaApiWebService.asmx"
strSoapAction = "http://tempuri.org/CheckPremium"

Dim strXmlBody As String
strXmlBody = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
                "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
                "<soap:Body>" & _
                "<CheckPremium xmlns=""http://tempuri.org/"">" & _
                "<Data>" & _
                "<ApiID>string</ApiID>" & _
                "<ApiPassword>string</ApiPassword>" & _
                "<LifeAssuredID>string</LifeAssuredID>" & _
                "<NewNRICNumber>string</NewNRICNumber>" & _
                "<OtherIdenNumber>string</OtherIdenNumber>" & _
                "<TPAClaimID>string</TPAClaimID>" & _
                "<LossDate>string</LossDate>" & _
                "<RiderCode>string</RiderCode>" & _
                "<PolicyNo>string</PolicyNo>" & _
                "</Data>" & _
                "</CheckPremium>" & _
                "</soap:Body>" & _
                "</soap:Envelope>"

Dim objDom As Object
Dim objXmlHttp As Object
Dim strRet As String

Set objDom = CreateObject("MSXML2.DOMDocument")
Set objXmlHttp = CreateObject("MSXML2.XMLHTTP")
objDom.async = False
objDom.LoadXML XmlBody
objXmlHttp.Open "POST", AsmxUrl, False
objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
objXmlHttp.Send objDom.XML
strRet = objXmlHttp.responseText
End Sub




'Dim MyWord As Word.Application
'Dim WordDoc As Word.Document
'Dim MyRange As Word.Range
'Dim MyRange2 As Word.Range
'Dim MyTable As Word.Table
'Dim MyCell As Word.Cell
'Dim MyCells As Word.Cells
'Dim MyCols As Word.Columns
'Dim RetSt As String

'Private Sub Form_Load()
'    Call ActionLogUpdate(12, "frmTest")
'    Dim WordApp As Word.Application
'    Dim BaseDoc As Word.Document
'    Set WordApp = New Word.Application
'    Set BaseDoc = WordApp.Documents.Open(LocCardPath & "\HCT_GL_Deferment_Letter.dot")
'    WordApp.Visible = True
'
'    WordApp.Selection.TypeText Chr(13)
'
'
'
'
'     'With WordApp.Selection
'     '   .TypeText "Aktuelle Kontakte mit " & Format(Date, "Long Date")
'     '   .TypeParagraph
'     '   .MoveLeft Unit:=wdWord, Count:=11, Extend:=wdExtend
'     '   .Font.Size = 14
'     '   .Font.Bold = wdToggle
'     '   .MoveDown Unit:=wdLine, Count:=1
'     '   'insert two column table to hold contact data(one column for contact names, the other for user comments)
'     '   .Tables.Add Range:=.Range, NumRows:=1, NumColumns:=2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
'     '   With .Tables(1)
'     '     .Style = "Table Grid"
'     '     .ApplyStyleHeadingRows = True
'     '     .ApplyStyleLastRow = False
'     '     .ApplyStyleFirstColumn = True
'     '     .ApplyStyleLastColumn = False
'     '     .ApplyStyleRowBands = True
'     '     .ApplyStyleColumnBands = False
'     '   End With
'    'End With
'
'
'
'
'    'Set MyRange = BaseDoc.Range(Start:=0, End:=0)
'    Set MyRange = BaseDoc.Paragraphs(1).Range
'
'    Dim tblNew As Table
'
'    Set tblNew = BaseDoc.Tables.Add(MyRange, 5, 2)
'
'    With tblNew
'        .Style = "Table Grid"
'        '.Borders(wdBorderTop).LineStyle = wdLineStyleDashDot
'        .Cell(Row:=1, Column:=1).Merge MergeTo:=.Cell(Row:=1, Column:=2)
'        '.Range.Font.Name = "Times New Roman"
'        '.Range.Font.Size = 12
'        .Rows.Height = 10
'    End With
'
'    'tblNew.Columns(1)
'
'    tblNew.Cell(1, 1).Range.InsertAfter "Reference"
'    tblNew.Cell(2, 1).Range.InsertAfter "Name of Patient"
'    tblNew.Cell(3, 1).Range.InsertAfter "Membership Number"
'    tblNew.Cell(4, 1).Range.InsertAfter "Case Number"
'    tblNew.Cell(5, 1).Range.InsertAfter "Treatment Date"
'
'    WordApp.Selection.EndKey (6)
'
'    WordApp.Selection.TypeParagraph
'
'    WordApp.Selection.TypeText Chr(13)
'
'    WordApp.Selection.TypeText ("Dear Sir/Madam," + vbCrLf + "Your request for GL(s) refers.")
'    WordApp.Selection.TypeText Chr(13)
'    WordApp.Selection.TypeText ("Please provide us the following document(s) in order for us to process the GL")
'
'    WordApp.Selection.TypeParagraph
'
'    With BaseDoc.Paragraphs(20)
'     .LineSpacingRule = wdLineSpaceSingle
'     .SpaceAfter = 1
'     .Range.Font.Bold = True
'     .Range.ListFormat.ApplyBulletDefault
'    End With
'
'    'WordApp.Selection.Range.ListFormat.ApplyBulletDefault
'    WordApp.Selection.TypeText ("Document1" + vbNewLine)
'    WordApp.Selection.TypeText ("Document2" + vbNewLine)
'    WordApp.Selection.TypeText ("Document3" + vbNewLine)
'
'
'
'   'Dim MyRangeQues As Word.Range
'   'Set MyRangeQues = BaseDoc.Paragraphs(23).Range
'
'   ' Dim tblNewQues As Table
'
'    Dim intX As Integer
'    Dim intY As Integer
'
'   ' Set tblNewQues = BaseDoc.Tables.Add(MyRangeQues, 5, 2)
'
'
'    'WordApp.Selection.TypeParagraph
'
'    'With BaseDoc.Paragraphs(23)
'    ' .LineSpacingRule = wdLineSpaceSingle
'    ' .SpaceAfter = 1
'    ' .Range.Font.Bold = True
'     '.Range.ListFormat.RemoveNumbers
'    'End With
'
'   WordApp.Selection.Style = "Normal"
'
'
'    WordApp.Selection.TypeText ("Additional remark" + vbNewLine)
'
'
'    For intX = 23 To 48
'        Dim tblNewQues As Table
'        Set tblNewQues = BaseDoc.Tables.Add(BaseDoc.Paragraphs(intX).Range, 2, 1)
'        tblNewQues.Style = "Table Grid"
'        tblNewQues.Cell(1, 1).Range.InsertAfter ("Question" + vbCrLf)
'        tblNewQues.Cell(1, 1).Range.Paragraphs(2).Range.Font.Bold = True
'        tblNewQues.Cell(1, 1).Range.InsertAfter ("test Q")
'        tblNewQues.Cell(2, 1).Range.InsertAfter ("Response")
'        tblNewQues.Rows(2).Height = 80
'        WordApp.Selection.EndKey (6)
'
'        WordApp.Selection.TypeParagraph
'        intX = intX + 5
'    Next intX
'    'WordApp.Selection.TypeText ("asfkj alwjd lajkd")
'
'    'Set MyRange2 = BaseDoc.Paragraphs(25).Range
'
'    'Dim tblNew2 As Table
'
'    'Set tblNew2 = BaseDoc.Tables.Add(MyRange2, 5, 2)
'    'tblNew2.Cell(1, 1).Range.InsertAfter "Reference"
'    'tblNew2.Cell(2, 1).Range.InsertAfter "Name of Patient"
'    'tblNew2.Cell(3, 1).Range.InsertAfter "Membership Number"
'
'
'    'WordApp.Selection.EndKey (6)
'    '    With BaseDoc.Paragraphs(2)
'    '        .Range.ListFormat.ApplyBulletDefault
'    '    End With
'
'
'
'
'
'
' 'With styTable.Table
'
' '.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
' '.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
' '.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
' '.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
'
' 'Apply a double border to the heading row
' '.Condition(wdFirstRow).Borders(wdBorderBottom) _
' '.LineStyle = wdLineStyleDouble
'
' 'Apply a double border to the last column
' '.Condition(wdLastColumn).Borders(wdBorderLeft) _
' '.LineStyle = wdLineStyleDouble
'
' 'Apply shading to last row
' '.Condition(wdLastRow).Shading _
' '.BackgroundPatternColor = wdColorGray125
'
' 'End With
'
'    'Set tblNew.Style = styTable
'
'    'With tblNew
'    '    For intX = 1 To 4
'    '        For intY = 1 To 2
'    '            .Cell(intX, intY).Range.InsertAfter "Cell: R" & intX & ", C" & intY
'    '        Next intY
'    '    Next intX
'    '    .Columns.AutoFit
'    'End With
'
'
'
'    'WordApp.Documents.Close
'    'WordApp.Quit
'
'    'Set WordApp = Nothing
'
'End Sub
'
