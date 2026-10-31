VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmailNotificationProvider 
   Caption         =   "Email Notification (Medical Provider)"
   ClientHeight    =   7272
   ClientLeft      =   60
   ClientTop       =   516
   ClientWidth     =   11808
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7272
   ScaleWidth      =   11808
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   6480
      Width           =   11655
      Begin VB.TextBox txtemailAdd 
         Height          =   285
         Left            =   2520
         TabIndex        =   30
         Top             =   240
         Width           =   4200
      End
      Begin VB.CommandButton CmdSend 
         Caption         =   "Email"
         Height          =   375
         Left            =   8280
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10440
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   7440
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print to Excel"
         Height          =   375
         Left            =   9240
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9240
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   8520
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtSendDate 
         Height          =   285
         Left            =   960
         TabIndex        =   31
         Top             =   240
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Send Date"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "BCC"
         Height          =   255
         Left            =   2160
         TabIndex        =   32
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Email Notification (Medical Provider)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   0
         Width           =   11445
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11655
      Begin VB.ComboBox cboEntityType 
         Height          =   315
         ItemData        =   "frmEmailNotificationProvider.frx":0000
         Left            =   6480
         List            =   "frmEmailNotificationProvider.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   480
         Width           =   3495
      End
      Begin VB.CheckBox checkall 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox TxtRecMNRB2 
         Height          =   285
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox TxtRecMNRB1 
         Height          =   285
         Left            =   7800
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox ChkDFMNRB 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5280
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   6480
         TabIndex        =   1
         Top             =   180
         Width           =   3495
      End
      Begin MSMask.MaskEdBox TxtDFMNRB2 
         Height          =   285
         Left            =   3480
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2985
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDFMNRB1 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPaymentDateFr 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   165
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPaymentDateTo 
         Height          =   285
         Left            =   3360
         TabIndex        =   8
         Top             =   165
         Width           =   1695
         _ExtentX        =   2985
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtPVNoTo 
         Height          =   285
         Left            =   3360
         MaxLength       =   15
         TabIndex        =   10
         Top             =   420
         Width           =   1680
      End
      Begin VB.TextBox TxtPVNoFr 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   9
         Top             =   420
         Width           =   1440
      End
      Begin VB.Label lblentitytype 
         Caption         =   "Entity Type"
         Height          =   255
         Left            =   5280
         TabIndex        =   37
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblheckall 
         Caption         =   "Select All"
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   9480
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Rec. No (MNRB)fr."
         Height          =   255
         Left            =   6360
         TabIndex        =   18
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   195
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Payment Date Fr"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   200
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Date Fr MNRB"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Status"
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   12
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label66 
         Caption         =   "PV No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstEmailListing 
      Height          =   5175
      Left            =   120
      TabIndex        =   22
      Top             =   1365
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   9123
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Clinic Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Clinic Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "PV No"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "PV Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cheque No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Cheque Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Cheque Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Email"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Bill Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "FFS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "CoPay"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmEmailNotificationProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_blnDirection(1 To 17) As Boolean

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboStatus.Text, cboStatus.SelStart) + Chr(KeyAscii) + Mid(cboStatus.Text, cboStatus.SelStart + cboStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboStatus.ListCount - 1
 
        If NewText = LCase(Left(cboStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboStatus.Text = ValidValue
                cboStatus.SelStart = 0
                cboStatus.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub CheckAll_Click()
Const METHOD_NAME = "CheckAll_Click"
On Error GoTo debugs
Dim LstCnt As Integer
    
    If lstEmailListing.ListItems.Count <> 0 And checkall <> 0 Then
        LstCnt = lstEmailListing.ListItems.Count
        For LstCnt = 1 To LstCnt
            If lstEmailListing.ListItems(LstCnt).SubItems(9) = "" Then
                lstEmailListing.ListItems(LstCnt).Checked = False
            ElseIf lstEmailListing.ListItems(LstCnt).SubItems(9) <> "" Then
                lstEmailListing.ListItems(LstCnt).Checked = True
            End If
        Next LstCnt
    ElseIf lstEmailListing.ListItems.Count <> 0 And checkall = 0 Then
         LstCnt = lstEmailListing.ListItems.Count
         For LstCnt = 1 To LstCnt
         lstEmailListing.ListItems(LstCnt).Checked = False
         Next LstCnt
    ElseIf checkall = 1 Then
        MsgBox "Please search a record.", vbOKOnly + vbCritical
        checkall = 0
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelAnalysis1(lstEmailListing)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
Dim strAppend As String
    lstEmailListing.ListItems.Clear
    Screen.MousePointer = vbHourglass

        If txtPaymentDateFr = "__/__/____" Then
            MsgBox "Please Enter PV Date From!", vbCritical
            txtPaymentDateFr.SetFocus
        ElseIf txtPaymentDateTo = "__/__/____" Then
            MsgBox "Please Enter PV Date To!", vbCritical
            txtPaymentDateTo.SetFocus
        ElseIf Trim(cboStatus) = "" Then
            MsgBox "Please Enter Status!", vbCritical
            cboStatus.SetFocus
        Else
        
            If Len(Trim(txtPaymentDateFr)) > 0 And IsDate(txtPaymentDateFr) = True Then
                strAppend = strAppend + " And PVCLNDate >= '" & Format(txtPaymentDateFr, "yyyy/mm/dd") & "'"
            End If
        
            If Len(Trim(txtPaymentDateTo)) > 0 And IsDate(txtPaymentDateTo) = True Then
                strAppend = strAppend + " And PVCLNDate <= '" & Format(txtPaymentDateTo, "yyyy/mm/dd") & "'"
            End If
    
            If Trim(TxtPVNoFr) <> "" Then
                strAppend = strAppend + " And PVCLNNo >= '" & Trim(TxtPVNoFr) & "'"
            End If
        
            If Trim(TxtPVNoTo) <> "" Then
                strAppend = strAppend + " And PVCLNNo <= '" & TxtPVNoTo & "'"
            End If
            
            If Mid(cboStatus, 1, 1) = "R" Then
                strAppend = strAppend + " And PVEmailStatus = 'Y'"
            Else
                strAppend = strAppend + " And ( PVEmailStatus  = '' or PVEmailStatus is null)"
            'and (PVRecMNRBStatus = '' or PVRecMNRBStatus is null) "
            End If
            
            If cboEntityType.ListIndex = 0 Then
                strAppend = strAppend + " AND Substring(PVCLNNo, 1, 2) <>'CS' "
            ElseIf cboEntityType.ListIndex = 1 Then
                strAppend = strAppend + " AND Substring(PVCLNNo, 1, 2) ='CS' "
            End If
            Call PaymentListing(strAppend)
        End If
    Screen.MousePointer = vbDefault
        
End Sub



Private Sub CmdSend_Click()
Const METHOD_NAME = "CmdSend_Click"
On Error GoTo debugs
On Err GoTo ErrorSave
Dim olApp As Outlook.Application
Dim objMail As Outlook.MailItem
Dim i As Integer
Dim ItemChecked As Boolean
Dim update
Dim StrSql As String
Dim rst As New ADODB.Recordset
Dim tmpvisit As String
Dim tmpCLMTotalActual As Double
Dim tmpCLMBordxAmt As Double
Dim rstMBM As New ADODB.Recordset
Dim strsqlMBM As String
Dim tmpPatientName As String
Dim strsqlupdate As String
Dim emailupdate As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Mailcmd As ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
Dim selectedEntity As String
    i = 0
    Screen.MousePointer = vbHourglass
    
    'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If txtSendDate = "__/__/____" Then
        MsgBox "Please enter send date", vbOKOnly + vbCritical
    Else
        ItemChecked = False
        For i = 1 To lstEmailListing.ListItems.Count
            If lstEmailListing.ListItems(i).Checked = True Then
                ItemChecked = True
                Exit For
            End If
        Next i
        
        If ItemChecked = False Then
            MsgBox "Please select the record", vbOKOnly + vbCritical
        
        Else
            update = MsgBox("Do you want to start sending email(s) ?", vbYesNo + vbExclamation)
            If update = vbYes Then
            
   
            Dim IsFirst As Boolean
            IsFirst = True
            
            
                For i = 1 To lstEmailListing.ListItems.Count
               
                    If lstEmailListing.ListItems(i).Checked = True Then
                    
                        Dim strTableBeg As String
                        Dim strTableEnd As String
                        Dim strFntNormal As String
                        Dim strFntHeader As String
                        Dim strFntEnd As String
                        Dim strMessage As String
                        Dim StrBodyMsg As String
                        
                        strTableBeg = "<table border=1>"
                        strTableEnd = "</table>"
                        strFntHeader = "<font size=2 face=" & Chr(34) & "Arial" & Chr(34) & "><b> <tr bgcolor=lightblue> <td>No</td> <td>Invoice No</td> <td>Invoice Date</td> <td>Cheque Date</td> <td>Bill Amount</td> <td>Co-Payment</td> <td>FFS</td> <td>Amount Paid</td></tr></b></font>"
                        strFntNormal = "<font color=black face=" & Chr(34) & "Arial" & Chr(34) & " size=1>"
                        strFntEnd = "</font>"
                        strMessage = strTableBeg & strFntNormal & strFntHeader
                        StrBodyMsg = ""
                        
                        tmpCLMTotalActual = 0
                        tmpCLMBordxAmt = 0
                        StrSql = ""
                        
                            Set Mailcmd = New ADODB.Command
                            Set Mailcmd.ActiveConnection = MediConnCISPayment
                            Mailcmd.CommandText = "sp_ProviderInv"
                            Mailcmd.CommandType = adCmdStoredProc
                            Mailcmd.CommandTimeout = 300
                                                    
                            Set Prm1 = Mailcmd.CreateParameter("StrAppend", adVarChar, adParamInput, 30, Trim(lstEmailListing.ListItems(i).SubItems(3)))
                            Mailcmd.Parameters.Append Prm1
                            Set prm2 = Mailcmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 1, "C")
                            Mailcmd.Parameters.Append prm2
                            rst.CursorLocation = adUseClient
                            rst.CursorType = adOpenForwardOnly
                            rst.CacheSize = 100
                            Set rst = Mailcmd.Execute
                            
                            Dim Count As Integer
                            Count = 1
                            Dim BillAmt As Double
                            Dim AmtPaid As Double
                            Dim copay As Double
                            Dim FFSTotal As Double
                            BillAmt = 0
                            AmtPaid = 0
                            copay = 0
                            FFSTotal = 0
                            
                            
                            Do While Not rst.EOF
                            
                            tmpCLMTotalActual = CDbl(tmpCLMTotalActual) + CDbl(rst!CLMTotalActual)
                            StrBodyMsg = StrBodyMsg & "<tr><td>" & Count & "</td> <td>" & rst!InvoiceNo & "</td> <td>" & rst!InvoiceDate & "</td> <td>" & rst!PVCLNDate & "</td> <td>" & Format(rst!CLMTotalActual, "###0.00") & "</td> <td>" & Format(rst!CLMCoPayColl, "###0.00") & "</td> <td>" & Format(rst!CLMMCOProcFee, "###0.00") & "</td> <td>" & Format(rst!CLMPayableByMedix2c, "###0.00") & "</td> </tr>"
                            
                            BillAmt = BillAmt + Format(rst!CLMTotalActual, "###0.00")
                            AmtPaid = AmtPaid + Format(rst!CLMPayableByMedix2c, "###0.00")
                            copay = copay + Format(rst!CLMCoPayColl, "###0.00")
                            FFSTotal = FFSTotal + Format(rst!CLMMCOProcFee, "###0.00")
                            selectedEntity = rst!ENTITYCODE
                            Count = Count + 1
                            rst.MoveNext
                            Loop
                            
                                                                              
                            rst.Close
                            Set rst = Nothing
                                               
                       strMessage = strMessage & StrBodyMsg & "<tr></tr> <tr><td>Total </td> <td>----</td> <td>----</td> <td>----</td> <td>" & Format(BillAmt, "###0.00") & "</td> <td>" & Format(copay, "###0.00") & "</td> <td>" & Format(FFSTotal, "###0.00") & "</td> <td>" & Format(AmtPaid, "###0.00") & "</td> </tr>"
                       strMessage = strMessage & strFntEnd & strTableEnd
                        
                        'Create e-mail item
                            Set olApp = Outlook.Application
                            If selectedEntity = "SJ" Then
                                selectedEntity = "MAJU SIHAT"
                            Else
                                selectedEntity = "HEALTH CONNECT"
                            End If
                            
                            Set objMail = olApp.CreateItem(olMailItem)
                            Dim Body As String
                        
                                    With objMail
                                    Body = "Dear Sir/Madam,</p></p> " & vbNewLine & vbNewLine & _
                                           "PAYMENT ADVICE </p>  " & vbNewLine & vbNewLine & _
                                           "We are pleased to advise that payment in respect of your invoice(s) as mentioned below has been banked into your Account. </p> Please allow reasonable time for the cheque to be cleared.</p>" & vbNewLine & vbNewLine & _
                                           "PV Number               :       " & lstEmailListing.ListItems(i).SubItems(3) & "</p>" & _
                                           "Cheque Number         :       " & lstEmailListing.ListItems(i).SubItems(5) & "</p>" & _
                                           "PV Amount               :       " & "RM " & Format(lstEmailListing.ListItems(i).SubItems(7), "###0.00") & "</p>" & _
                                           "Date Banked-in         :       " & txtSendDate & "</p>" & _
                                           "DETAILS OF PAYMENT" & "</p>" & _
                                           strMessage & "</p>" & _
                                           "</p>Please contact our following staff if you require any clarification:- </p> " & _
                                           "           General Line    -  03-7884 1919 </p></p>" & _
                                           "Thank you. </p>" & _
                                           "Regards,</p>" & _
                                           selectedEntity + " SDN BHD"
                                          .To = lstEmailListing.ListItems(i).SubItems(9)
                                            If txtemailAdd <> Null Then
                                                .BCC = "elaineyeow@medix.com.my"
                                            Else
                                                .BCC = "elaineyeow@medix.com.my" & ";" & txtemailAdd
                                            End If
                                        .Subject = selectedEntity + " Payment Advice" & " - " & lstEmailListing.ListItems(i).SubItems(2) & " (" & lstEmailListing.ListItems(i).SubItems(1) & ")"
                                        .BodyFormat = olFormatHTML
                                        .HTMLBody = "<HTML><BODY>" & strFntNormal & Body & " </BODY></HTML>"
                                        .Send
                                    End With
                           strsqlupdate = ""
                            strsqlupdate = " update Payment2CLN set  PVEmailStatus = 'Y' " & _
                                     " Where PVCLNNo = '" & Trim(lstEmailListing.ListItems(i).SubItems(3)) & "'"
                            MediConnCISPayment.Execute strsqlupdate
                        
                    End If
                Next i
           End If
        End If
        End If
        
    Screen.MousePointer = vbDefault
    MsgBox "Emails have been sent out successfully!", vbOKOnly + vbCritical
    Exit Sub
ErrorSave:
        Screen.MousePointer = vbDefault
        
        MediConnCISPayment.Close
        Set MediConnCISPayment = Nothing
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        
        MsgBox Err.Description
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub PaymentListing(Optional strAppend As String)
Const METHOD_NAME = "PaymentListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim CasNumber As String
Dim WSVersion As String
Dim Record As ListItem
Dim Check As Boolean
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim ContrSql As String
Dim ContraREc As New ADODB.Recordset
Dim PayCurrRec As New ADODB.Recordset
Dim StrSql As String
Dim Pay2MBMAmt As Double
Dim Paid2MBMAmt As Double
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpReceiptNo As String
Dim tmpReceiptAmt As Double
Dim sqlstr As String
Dim rstMBM As New ADODB.Recordset
Dim tmpPatientName As String
Dim tmpPaymentdate As String


    RecFound = False
    'Set cmd1.ActiveConnection = MediConnCISPayment
    'cmd1.CommandTimeout = 300
    'cmd1.CommandType = adCmdStoredProc
    'cmd1.CommandText = "EmailNotification"
    'Set Prm1 = cmd1.CreateParameter("strAppendCase", adVarChar, adParamInput, 1, "1")
    'cmd1.Parameters.Append Prm1
    'Set Prm2 = cmd1.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strAppend)
    'cmd1.Parameters.Append Prm2
    'rst2.CursorLocation = adUseClient
    'rst2.CursorType = adOpenForwardOnly
    'rst2.CacheSize = 100
    'Set rst2 = cmd1.Execute
    Dim Strappend1 As String
    Strappend1 = " ORDER BY CLNEMAIL DESC "
    sqlstr = ""
    sqlstr = "Select * from View_email_Provider Where 0 = 0"
            
        If Len(Trim(strAppend)) Then
            sqlstr = sqlstr + strAppend + Strappend1
        End If

    rst2.Open sqlstr, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly

    Do
        Do While Not rst2.EOF
        RecFound = True
            With rst2
            
                Set Record = lstEmailListing.ListItems.Add(, , "")
                If Len(!ClinicCode) Then
                    Record.SubItems(1) = !ClinicCode
                End If
                Record.SubItems(2) = !CLNName
                Record.SubItems(3) = !PVCLNNo
                Record.SubItems(4) = !PVCLNDate
                Record.SubItems(5) = !PVCLNChequeNo
                Record.SubItems(6) = !PVCLNChequeDate
                Record.SubItems(7) = Format(!PVCLNChequeAmt, "###0.00")
                Record.SubItems(8) = ""
                Record.SubItems(9) = IIf(Len(!CLNEMAIL), !CLNEMAIL, "")
            End With
            rst2.MoveNext
            
            DoEvents
            'If intExit = 1 Then
            '   Check = False
            '   Exit Do
            'End If
        Loop
    Loop Until Check = False
    rst2.Close
    Set rst2 = Nothing
    
    If RecFound = False Then
       MsgBox "No record found !", vbOKOnly + vbCritical, DialogBoxTitle
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "frmEmailNotificationProvider")
    cboStatus.AddItem "N - New"
    cboStatus.AddItem "R - Resend"
    cboEntityType.ListIndex = 0
End Sub




Private Sub lstEmailListing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstEmailListing, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstEmailListing, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstEmailListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstEmailListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(12)
    Case 13
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstEmailListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(14)
    Case 15
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(17)
    End Select

End Sub
Private Sub lstEmailListing_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ProcessCheckedItem(Item)
End Sub
Private Sub txtPaymentDateFr_LostFocus()
    If txtPaymentDateFr <> "__/__/____" Then
        If IsDate(txtPaymentDateFr) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtPaymentDateFr.SetFocus
        End If
    End If
End Sub
Private Sub txtPaymentDateTo_LostFocus()
    If txtPaymentDateTo <> "__/__/____" Then
        If IsDate(txtPaymentDateTo) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtPaymentDateTo.SetFocus
        End If
    End If
End Sub
Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
Dim RecordSaved
    If Item.Checked = True Then
        If Item.SubItems(9) = "" Then
            MsgBox "Please check member's email!", vbCritical
            Item.Checked = False
        ElseIf Item.SubItems(8) = "Y" Then
            RecordSaved = MsgBox("Email has been sent to the member! Do you want to send again?", vbYesNo + vbExclamation, DialogBoxTitle)
                If RecordSaved = vbNo Then
                    Item.Checked = False
                End If
        End If
    End If
End Function

