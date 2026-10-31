Attribute VB_Name = "modDownloadBinaryFile"
'************************************************************************************************************
'*  Author: SRI                                                                                             *
'*  Created Date: 09/09/2019                                                                                *
'*  Description: This program downloads the varBinary stored files                                          *
'************************************************************************************************************
Dim rs  As ADODB.Recordset
Dim sql As String
Dim oStream As ADODB.Stream
Dim hwnd As Long

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function ReimbFileList(caseNo As String) As String()
    Dim arr() As String
    Dim i As Integer
    
    Dim dbCn As New ADODB.Connection
    Set dbCn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    'dbCn.Open (MediConnCISDB)
    dbCn.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    dbCn.CursorLocation = adUseClient
    Dim cmd As New ADODB.Command
    Set cmd = New ADODB.Command

    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "spGetReimbClaimFileCount"
        .Parameters.Append .CreateParameter("@CaseNo", adVarChar, adParamInput, 20, caseNo)
        
        Set .ActiveConnection = dbCn
        Set rs = .Execute()
     
        ReDim arr(1 To rs.RecordCount)
    End With
    
    Do While Not rs.EOF
        arr(rs.AbsolutePosition) = Trim(CStr(rs("FileName")))
        rs.MoveNext
    Loop
    
    ReimbFileList = arr
End Function

Public Function downloadBinaryFile(caseNo As String, fullFileName As String) As Boolean
    Set rs = New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = MediConnCISDB
    Dim Prm1 As ADODB.Parameter
    Set Prm1 = cmd.CreateParameter("fullMbmNo", adVarChar, adParamInput, 2000, Null)
    cmd.Parameters.Append Prm1
    Dim prm2 As ADODB.Parameter
    Set prm2 = cmd.CreateParameter("caseNo", adVarChar, adParamInput, 2000, caseNo)
    cmd.Parameters.Append prm2
    Dim prm3 As ADODB.Parameter
    Set prm3 = cmd.CreateParameter("fileName", adVarChar, adParamInput, 2000, IIf(fullFileName = "", Null, fullFileName))
    cmd.Parameters.Append prm3
    cmd.CommandText = "spGetflexCAREReimbClaim"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.CacheSize = 100
    
    Set rs = cmd.Execute
    
    Dim fpath As String
    
    Do Until rs.EOF
        Set oStream = New ADODB.Stream
        With oStream
             .Type = adTypeBinary
             .Open
             .Write rs.Fields(9).Value
             .SaveToFile Environ("USERPROFILE") & "\Downloads\" & rs.Fields(7).Value, adSaveCreateOverWrite
             .Close
        End With
        
         fpath = Environ("USERPROFILE") & "\Downloads\" & rs.Fields(7).Value
        
        Set oStream = Nothing
        rs.MoveNext
    Loop
     
    rs.Close
    Dim hwnd As Long
    ShellExecute hwnd, "open", fpath, vbNullString, vbNullString, 1
End Function


