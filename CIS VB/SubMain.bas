Attribute VB_Name = "SubMain"
Option Explicit
Global medixmasterconn As New ADODB.Connection
Global medixmasterconnA As New ADODB.Connection
Global medixmasterconnWS As New ADODB.Connection
Global medixmasterconnPayment As New ADODB.Connection
Global MediConnCISPayment As New ADODB.Connection
Global MediConnCISWS As New ADODB.Connection
Global medixmasterconnMaint As New ADODB.Connection
Global MediConnCISDB As New ADODB.Connection
Global MediConnCISDentWS As New ADODB.Connection
Global MediConnCISWSPluto As New ADODB.Connection
Global MediConnCISDBPluto As New ADODB.Connection
Global medixmasterSUNSys As New ADODB.Connection


'======== Jupiter Server ================================
Global MediNeptuneConnCISDB As New ADODB.Connection
Global MediNeptuneConnCISWS As New ADODB.Connection
Global MediNeptuneConnCISPayment As New ADODB.Connection
Global medixneptuneconn As New ADODB.Connection
Global medixneptuneconnA As New ADODB.Connection
Global medixneptuneconnWS As New ADODB.Connection
Global medixneptuneconnPayment As New ADODB.Connection
'======== Jupiter Server ================================

Const NoError = 0
Global strPendingReasonId As String
Global strAddPendRemark As String
Global strQAtoHos As String

Global Const DialogBoxTitle = "MedixCIS - V1.1.5"
Global Const Version = "MediExpress Clinical Information System"
'----------- Change This Path If You Use For Testing  -----------------------

Global TrialVers As Boolean
Global Expired As Boolean
Global FirstRun As String
Global FinalRun As String
Global ExpiredKey As String

'Global Const BOSPath = "\\HIS\MedixHIS\"
Public BOSPath As String
Public crdPath As String
Public SRVPath  As String
Public LocSignPath As String
Public LocCardPath As String
Public DocPath As String
Public BordxSavePath As String
Public MCOSavePath As String
'Global Const crdPath = "\\HIS\Medix-HIS\HIS-Development\"
'Global Const SRVPath = "\\PDC\"
'Global Const ServerPath = "\\NEPTUNE\MedixHIS"
'----------------------------------------------------------------------------
Public fMainForm As MDIForm

Sub Main()
On Error GoTo Err_InitialiseError
Dim f As New frmLogin

BOSPath = ReadINI("path", "SPATH", App.path & "\SVRSuper.ini")
crdPath = ReadINI("path", "reportPath", App.path & "\SVRSuper.ini") '
SRVPath = ReadINI("path", "server", App.path & "\SVRSuper.ini")
LocSignPath = ReadINI("path", "LocSignPath", App.path & "\SVRSuper.ini")
LocCardPath = ReadINI("path", "LocreportPath", App.path & "\SVRSuper.ini")
DocPath = ReadINI("path", "DocPath", App.path & "\SVRSuper.ini")
BordxSavePath = ReadINI("path", "BordxPath", App.path & "\SVRSuper.ini")
MCOSavePath = ReadINI("path", "MCOPath", App.path & "\SVRSuper.ini")

medixmasterconnA.Open "File Name=" & BOSPath & "Conn\mediconnectA.UDL"
medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
'medixmasterSUNSys.Open "File Name=" & BOSPath & "Conn\MediConnSUNSystem.UDL"

'MediConnCISWSPluto.Open "File Name=" & BOSPath & "Conn\MediConnCISWSPluto.UDL"
'MediConnCISDentWS.Open "File Name=" & BOSPath & "Conn\MediConnCISDentWS.UDL"

'========== Jupiter Server ==================================================
'MediNeptuneConnCISDB.Open "File Name=" & BOSPath & "Conn\MediNeptuneConnCISDB.UDL"
'MediNeptuneConnCISWS.Open "File Name=" & BOSPath & "Conn\MediNeptuneConnCISWS.UDL"
'MediNeptuneConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediNeptuneConnCISPayment.UDL"
'medixneptuneconnA.Open "File Name=" & BOSPath & "Conn\medineptuneconnectA.UDL"
'medixneptuneconn.Open "File Name=" & BOSPath & "Conn\medineptuneconnect.UDL"
'medixneptuneconnWS.Open "File Name=" & BOSPath & "Conn\medineptuneconnectWS.UDL"
'medixneptuneconnPayment.Open "File Name=" & BOSPath & "Conn\medineptuneconnectPayment.UDL"
'========== Jupiter Server ==================================================

'
    f.Show

Exit_Initialise:
    Exit Sub

Err_InitialiseError:
    MsgBox Err.Description
    Resume Exit_Initialise
End Sub



