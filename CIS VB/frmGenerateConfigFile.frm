VERSION 5.00
Begin VB.Form frmGenerateConfigFile 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmGenerateConfigFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGenerate_Click()
Dim strsql As String
Dim CLMBordx As New ADODB.Recordset
Dim PAYCode As String
Dim sql As String
Dim ClaimNo As String
Dim VersionNo As String
Dim strAppend As String
Dim DataString As String
Dim tmpMCDays As Integer
Dim tmpClaimActual As Double
Dim tmpTotalAmount As Double
Dim tmpBordxDay As String
Dim tmpBordxMonth As String
Dim tmpBordxYear As String
Dim tmpRelationship As String
Dim tmpbordxdate As String
Dim z As Boolean
Dim Header As String
Dim Trailer As String
Dim tmpstring As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim total As Integer

    MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
        sql = "SELECT * From CLNConfigfile "
        strsql = sql + strAppend
        CLMBordx.Open sql, MediConnCISWS, adOpenStatic, adLockOptimistic

    
        If CLMBordx.EOF Then
            MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            
            txtFilName = "CC" & tmpBordxYear & tmpBordxMonth & tmpBordxDay & "01"
            z = CreatePath(txtLocalFileLocation)
            
            Open txtPath & "\" & txtFilName For Output As #1
            
            Header = txtFilName
            
            Print #1, Trim(Header)

        Do Until CLMBordx.EOF
            DataString = ""
            tmpClaimActual = 0
            tmpTotalAmount = 0
            tmpbordxdate = CLMBordx!BordxDate
            DataString = "1" & ","
            DataString = DataString & IIf(Len(CLMBordx!BordxDate), Trim(CLMBordx!BordxDate), "") & ","
            

            CLMBordx.MoveNext
            
            Print #1, DataString
        Loop
                Print #1, Trailer

                cmdGenerate.Enabled = False
                Close #1
                
        MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
        cmdGenerate.Enabled = True
    End If
    
    CLMBordx.Close
    Set CLMBordx = Nothing
    
    MediConnCISWS.Close
    Set MediConnCISWS = Nothing
End Sub
