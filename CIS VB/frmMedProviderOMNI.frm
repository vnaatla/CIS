VERSION 5.00
Begin VB.Form frmMedProviderOMNI 
   Caption         =   "Form1"
   ClientHeight    =   3024
   ClientLeft      =   60
   ClientTop       =   516
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3024
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtLocalFileLocation 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtFilName 
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmMedProviderOMNI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()

Call ActionLogUpdate(12, "frmMedProviderOMNI")

End Sub

Private Sub CmdGenerate_Click()
    Dim HOS As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    Dim tmpcnt As Double
    Dim Trailer As String
    Dim tmpfilename As String
    Dim tmpLocalFileLocation As String
    Dim DataString As String
    Dim Header As String
    
    tmpcnt = 0
    
    HOS.Open "select * from View_CLN_OMNI where CLNPanel = 'C' or CLNPanel = 'Y'", MediConnCISDB, adOpenKeyset, adLockOptimistic
        
    If HOS.EOF Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        txtLocalFileLocation = "C:\MedicaGenClaimASCII\"

        txtFilName = Format(Date, "yyyymmdd")
        txtFilName = txtFilName & "_OUTPATIENTPROVIDER" & ".txt"
        Open txtLocalFileLocation & "\" & txtFilName For Output As #1
        
        Header = "01" & "|" & "TPA" & "|" & Format(Date, "yyyymmdd")
            'Debug.Print Trailer
        Print #1, Header
            
        Do Until HOS.EOF
             
                tmpcnt = tmpcnt + 1
                DataString = "02" & "|"
                
                DataString = DataString & "N" & "|"
                           
                DataString = DataString & Trim(HOS!CLNCode) & "|"
                
                DataString = DataString & Trim(HOS!CLNName) & "|"
                
                DataString = DataString & Trim(HOS!CLNPhone) & "|"
                
                DataString = DataString & Trim(HOS!CLNAdd1) & "|"
                
                DataString = DataString & Trim(HOS!CLNAdd2) & "|"
                
                DataString = DataString & Trim(HOS!CLNAdd3) & "|"
                
                DataString = DataString & Trim(HOS!CLNCity) & "|"
                
                DataString = DataString & Trim(HOS!CLNPostCode) & "|"
                
                DataString = DataString & "" & "|"

                DataString = DataString & "MALAYSIA" & "|"
                
                DataString = DataString & "2" & "|"
                
                DataString = DataString & "N" & "|"
                
                DataString = DataString & "Y"
                

                
                Print #1, DataString
                
                DataString = DataString & Chr(13)
            HOS.MoveNext
            Loop

                Trailer = "03" & "|" & tmpcnt
                'Debug.Print Trailer
                Print #1, Trailer
                Close #1
                
                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                cmdGenerate.Enabled = True
        End If
    HOS.Close
    Set HOS = Nothing

End Sub

