VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmailRejectedClaim 
   Caption         =   "Rejected Claims Notification"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   6480
      Width           =   11655
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print to Excel"
         Height          =   375
         Left            =   9240
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   2640
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   7440
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   4440
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10440
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSend 
         Caption         =   "Email"
         Height          =   375
         Left            =   8280
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Rejected Claims Notification"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   240
         TabIndex        =   23
         Top             =   0
         Width           =   11445
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11655
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   6480
         TabIndex        =   7
         Top             =   180
         Width           =   3495
      End
      Begin VB.CheckBox ChkDFMNRB 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5280
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtRecMNRB1 
         Height          =   285
         Left            =   7800
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
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
      Begin VB.ComboBox cboGroupCompany 
         Height          =   315
         Left            =   6480
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   460
         Width           =   3495
      End
      Begin MSMask.MaskEdBox TxtDFMNRB2 
         Height          =   285
         Left            =   3480
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDFMNRB1 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPaymentDateFr 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   165
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPaymentDateTo 
         Height          =   285
         Left            =   3360
         TabIndex        =   11
         Top             =   165
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtPVNoFr 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   3
         Top             =   420
         Width           =   1440
      End
      Begin VB.TextBox TxtPVNoTo 
         Height          =   285
         Left            =   3360
         MaxLength       =   15
         TabIndex        =   2
         Top             =   420
         Width           =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "Group Company"
         Height          =   255
         Index           =   1
         Left            =   5280
         TabIndex        =   21
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label66 
         Caption         =   "Claim No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   19
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Status"
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Date Fr MNRB"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Reg Date Fr"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   200
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   195
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Rec. No (MNRB)fr."
         Height          =   255
         Left            =   6360
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   9480
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView lstEmailListing 
      Height          =   5415
      Left            =   120
      TabIndex        =   24
      Top             =   1120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9551
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
      NumItems        =   24
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Employee Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IC Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Company Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Bill Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Paid Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Visit Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Provider Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Claim Number"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Plan Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Payment Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Account Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Date Sumitted"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "MBMNumber"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Email"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Group Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Reason"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Invoice No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Action"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Rej Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "PolicyNumber"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Employee Id"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmEmailRejectedClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_blnDirection(1 To 24) As Boolean

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

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelAnalysis1(lstEmailListing)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
Dim StrAppend As String
    lstEmailListing.ListItems.Clear
    Screen.MousePointer = vbHourglass

        If txtPaymentDateFr = "__/__/____" Then
            MsgBox "Please Enter Registration Date From!", vbCritical
            txtPaymentDateFr.SetFocus
        ElseIf txtPaymentDateTo = "__/__/____" Then
            MsgBox "Please Enter Registration Date To!", vbCritical
            txtPaymentDateTo.SetFocus
        ElseIf Trim(cboStatus) = "" Then
            MsgBox "Please Enter Status!", vbCritical
            cboStatus.SetFocus
        Else
        
            If Len(Trim(txtPaymentDateFr)) > 0 And IsDate(txtPaymentDateFr) = True Then
                StrAppend = StrAppend + " And RegisterDate >= '" & Format(txtPaymentDateFr, "yyyy/mm/dd") & "'"
            End If
        
            If Len(Trim(txtPaymentDateTo)) > 0 And IsDate(txtPaymentDateTo) = True Then
                StrAppend = StrAppend + " And RegisterDate <= '" & Format(txtPaymentDateTo, "yyyy/mm/dd") & "'"
            End If
    
            If Trim(TxtPVNoFr) <> "" Then
                StrAppend = StrAppend + " And CISCas.CasNumber >= '" & Trim(TxtPVNoFr) & "'"
            End If
        
            If Trim(TxtPVNoTo) <> "" Then
                StrAppend = StrAppend + " And CISCas.CasNumber <= '" & TxtPVNoTo & "'"
            End If
            
            If Mid(cboStatus, 1, 1) = "R" Then
                StrAppend = StrAppend + " And CASEmailRejStatus = 'Y'"
            Else
                StrAppend = StrAppend + " And ( CASEmailRejStatus  = '' or CASEmailRejStatus is null)"
            'and (PVRecMNRBStatus = '' or PVRecMNRBStatus is null) "
            End If
            
            If Trim(cboGroupCompany) <> "" Then
                StrAppend = StrAppend + " And MBMGroupCompany = '" & cboGroupCompany & "'"
            End If

            
            'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            
                Call PaymentListing(StrAppend)
            
            'MediConnCISPayment.Close
            'Set MediConnCISPayment = Nothing
            'medixmasterconn.Close
            'Set medixmasterconn = Nothing
        End If
        
    Screen.MousePointer = vbDefault
        
End Sub

Private Sub CmdSend_Click()
Const METHOD_NAME = "CmdSend_Click"
On Error GoTo debugs
On Err GoTo ErrorSave
Dim oOApp As Outlook.Application
Dim oOMail As Outlook.MailItem
Dim i As Integer
Dim ItemChecked As Boolean
Dim update
Dim strsql As String
Dim rst As New ADODB.Recordset
Dim tmpvisit As String
Dim tmpCLMTotalActual As Double
Dim tmpCLMBordxAmt As Double
Dim rstMBM As New ADODB.Recordset
Dim strsqlMBM As String
Dim tmpPatientName As String
Dim strsqlupdate As String
Dim emailupdate As New ADODB.Recordset

    i = 0
    Screen.MousePointer = vbHourglass
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
            
                For i = 1 To lstEmailListing.ListItems.Count
                    If lstEmailListing.ListItems(i).Checked = True Then
                        tmpvisit = ""
                        tmpCLMTotalActual = 0
                        tmpCLMBordxAmt = 0
                        
                      
                        
                        'check that record is BSN or not
                        If Trim(lstEmailListing.ListItems(i).SubItems(3)) = Trim("BANK SIMPANAN NASIONAL") Then
                                If (Len(lstEmailListing.ListItems(i).SubItems(16)) = 0) Or (Trim(Left(lstEmailListing.ListItems(i).SubItems(16), 3)) = Trim("bsn")) Then
                                    'Save doc
                                    Dim word_app As Word.Application
                                    Dim word_doc As Word.Document
                                    Dim objSelection
                                    Dim file_name As String
                        
                                    If Dir("C:\Template\RejectedMails\", vbDirectory) = "" Then
                                         MkDir ("C:\Template\RejectedMails\")
                                    End If
                        
                                    If InStr(1, lstEmailListing.ListItems(i).SubItems(15), "*") Then
                                        Dim sArr() As String
                                        sArr = Split(lstEmailListing.ListItems(i).SubItems(15), "*")
                                        file_name = "C:\Template\RejectedMails\" + sArr(0) + " _ " + Trim(lstEmailListing.ListItems(i).SubItems(1))
                                    Else
                                        file_name = "C:\Template\RejectedMails\" + lstEmailListing.ListItems(i).SubItems(15) + " _ " + Trim(lstEmailListing.ListItems(i).SubItems(1))
                                    End If
    
                                    'Print details with word
                                    Set word_app = New Word.Application
                                    Set word_doc = word_app.Documents.Add(DocumentType:=wdNewBlankDocument)
                        
                        
                                    word_app.Selection.TypeParagraph
                                    word_app.Selection.TypeParagraph
                                    word_app.Selection.TypeText "Dear Sir/Madam" & vbNewLine & _
                                                                "CLAIM DETAILS" & vbNewLine & _
                                                                "Member Name            :" & Trim(lstEmailListing.ListItems(i).SubItems(1)) & vbNewLine & _
                                                                "Policy Number          :" & Trim(lstEmailListing.ListItems(i).SubItems(22)) & vbNewLine & _
                                                                "IC Number              :" & Trim(lstEmailListing.ListItems(i).SubItems(2)) & vbNewLine & _
                                                                "Claim No               :" & Trim(lstEmailListing.ListItems(i).SubItems(9)) & vbNewLine & _
                                                                "InvoiceNo              :" & Trim(lstEmailListing.ListItems(i).SubItems(19)) & vbNewLine & _
                                                                "Membership No          :" & lstEmailListing.ListItems(i).SubItems(15) & vbNewLine & _
                                                                "Date Submitted         :" & lstEmailListing.ListItems(i).SubItems(13) & vbNewLine & _
                                                                "Visit Date             :" & lstEmailListing.ListItems(i).SubItems(7) & vbNewLine & _
                                                                "Incurred Amount        :" & "RM " & Format(lstEmailListing.ListItems(i).SubItems(21), "###0.00") & vbNewLine & vbNewLine & _
                                                                "Thank you for submitting the above claim to us. Please see the following reasons(s) and action(s) required(if any) :-" & vbNewLine & _
                                                                "Reason(s) - " & lstEmailListing.ListItems(i).SubItems(18) & vbNewLine & vbNewLine & _
                                                                "Action(s) - " & lstEmailListing.ListItems(i).SubItems(20) & vbNewLine & vbNewLine & _
                                                                "All original documents will be retunred to your HR department within Fourteen(14) working days" & vbNewLine & vbNewLine & _
                                                                "Kindly reply within 30 days or claim will be rejected and will not be entertained." & vbNewLine & vbNewLine & _
                                                                "Please contact our following staff if you require any clarification:- " & vbNewLine & _
                                                                "           General Line    -  03-7884 1919" & vbNewLine & vbNewLine & _
                                                                "Thank you."
                                                   
                                                
                                    Set objSelection = word_app.Selection
                                    objSelection.WholeStory
                                    objSelection.Font.Name = "Times New Roman"
                                    objSelection.Font.Size = 10
                                    objSelection.Font.Bold = False
                                    word_doc.SaveAs Filename:=file_name
                                    word_doc.Close False
                                    word_app.Quit False
    
                                Else
                                   'Send Email
                                     Set oOApp = CreateObject("Outlook.Application")
                                     Set oOMail = oOApp.CreateItem(olMailItem)
                        
                                   With oOMail
                                   .To = lstEmailListing.ListItems(i).SubItems(16)
                                   .BCC = "foo@medix.com.my;"
                                   .Subject = "CLAIM - " & lstEmailListing.ListItems(i).SubItems(4) & "(Ref No : " & lstEmailListing.ListItems(i).SubItems(9) & ")"
                                   .Body = "Dear Sir/Madam" & vbNewLine & vbNewLine & _
                                        "CLAIM DETAILS" & vbNewLine & _
                                        "Member Name            :" & Trim(lstEmailListing.ListItems(i).SubItems(1)) & vbNewLine & _
                                        "Policy Number          :" & Trim(lstEmailListing.ListItems(i).SubItems(22)) & vbNewLine & _
                                        "IC Number              :" & Trim(lstEmailListing.ListItems(i).SubItems(2)) & vbNewLine & _
                                        "Claim No               :" & Trim(lstEmailListing.ListItems(i).SubItems(9)) & vbNewLine & _
                                        "InvoiceNo              :" & Trim(lstEmailListing.ListItems(i).SubItems(19)) & vbNewLine & _
                                        "Membership No          :" & lstEmailListing.ListItems(i).SubItems(15) & vbNewLine & _
                                        "Date Submitted         :" & lstEmailListing.ListItems(i).SubItems(13) & vbNewLine & _
                                        "Visit Date             :" & lstEmailListing.ListItems(i).SubItems(7) & vbNewLine & _
                                        "Incurred Amount        :" & "RM " & Format(lstEmailListing.ListItems(i).SubItems(21), "###0.00") & vbNewLine & vbNewLine & _
                                        "Thank you for submitting the above claim to us. Please see the following reasons(s) and action(s) required(if any) :-" & vbNewLine & _
                                        "Reason(s) - " & lstEmailListing.ListItems(i).SubItems(18) & vbNewLine & vbNewLine & _
                                        "Action(s) - " & lstEmailListing.ListItems(i).SubItems(20) & vbNewLine & vbNewLine & _
                                        "All original documents will be retunred to your HR department within Fourteen(14) working days" & vbNewLine & vbNewLine & _
                                        "Kindly reply within 30 days or claim will be rejected and will not be entertained." & vbNewLine & vbNewLine & _
                                        "Please contact our following staff if you require any clarification:- " & vbNewLine & _
                                        "           General Line    -  03-7884 1919" & vbNewLine & vbNewLine & _
                                        "Thank you." & vbNewLine & _
                                        "Regards, " & vbNewLine & _
                                        "MediExpress (Malaysia) Sdn Bhd"
                                .Send
                                '.Display
                             End With
                            End If
                        Else
                                Set oOApp = CreateObject("Outlook.Application")
                                Set oOMail = oOApp.CreateItem(olMailItem)
                                With oOMail
                                .To = lstEmailListing.ListItems(i).SubItems(16)
                                .BCC = "foo@medix.com.my;"
                                .Subject = "CLAIM - " & lstEmailListing.ListItems(i).SubItems(4) & "(Ref No : " & lstEmailListing.ListItems(i).SubItems(9) & ")"
                                .Body = "Dear Sir/Madam" & vbNewLine & vbNewLine & _
                                        "CLAIM DETAILS" & vbNewLine & _
                                        "Member Name            :" & Trim(lstEmailListing.ListItems(i).SubItems(1)) & vbNewLine & _
                                        "Policy Number          :" & Trim(lstEmailListing.ListItems(i).SubItems(22)) & vbNewLine & _
                                        "IC Number              :" & Trim(lstEmailListing.ListItems(i).SubItems(2)) & vbNewLine & _
                                        "Claim No               :" & Trim(lstEmailListing.ListItems(i).SubItems(9)) & vbNewLine & _
                                        "InvoiceNo              :" & Trim(lstEmailListing.ListItems(i).SubItems(19)) & vbNewLine & _
                                        "Membership No          :" & lstEmailListing.ListItems(i).SubItems(15) & vbNewLine & _
                                        "Date Submitted         :" & lstEmailListing.ListItems(i).SubItems(13) & vbNewLine & _
                                        "Visit Date             :" & lstEmailListing.ListItems(i).SubItems(7) & vbNewLine & _
                                        "Incurred Amount        :" & "RM " & Format(lstEmailListing.ListItems(i).SubItems(21), "###0.00") & vbNewLine & vbNewLine & _
                                        "Thank you for submitting the above claim to us. Please see the following reasons(s) and action(s) required(if any) :-" & vbNewLine & _
                                        "Reason(s) - " & lstEmailListing.ListItems(i).SubItems(18) & vbNewLine & vbNewLine & _
                                        "Action(s) - " & lstEmailListing.ListItems(i).SubItems(20) & vbNewLine & vbNewLine & _
                                        "All original documents will be retunred to your HR department within Fourteen(14) working days" & vbNewLine & vbNewLine & _
                                        "Kindly reply within 30 days or claim will be rejected and will not be entertained." & vbNewLine & vbNewLine & _
                                        "Please contact our following staff if you require any clarification:- " & vbNewLine & _
                                        "           General Line    -  03-7884 1919" & vbNewLine & vbNewLine & _
                                        "Thank you." & vbNewLine & _
                                        "Regards, " & vbNewLine & _
                                        "MediExpress (Malaysia) Sdn Bhd"
                            .Send
                            '.Display
                             End With
                        End If
                            strsqlupdate = ""
                            strsqlupdate = " update CISCAS set  CASEmailRejStatus = 'Y' " & _
                                     " Where CASNumber = '" & Trim(lstEmailListing.ListItems(i).SubItems(9)) & "'"
                            MediConnCISDB.Execute strsqlupdate
                       
                    End If
                Next i
           End If
        End If
    Screen.MousePointer = vbDefault
    MsgBox "Emails have been sent out successfully!", vbOKOnly + vbCritical
    Exit Sub
ErrorSave:
        Screen.MousePointer = vbDefault
        MsgBox Err.Description
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub PaymentListing(Optional StrAppend As String)
Const METHOD_NAME = "PaymentListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim CasNumber As String
Dim WSVersion As String
Dim Record As ListItem
Dim Check As Boolean
Dim RECSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim ContrSql As String
Dim ContraREc As New ADODB.Recordset
Dim PayCurrRec As New ADODB.Recordset
Dim strsql As String
Dim Pay2MBMAmt As Double
Dim Paid2MBMAmt As Double
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpReceiptNo As String
Dim tmpReceiptAmt As Double
Dim sqlstr As String
Dim rstMBM As New ADODB.Recordset
Dim tmpPatientName As String
Dim tmpPaymentdate As String


    RecFound = False
    Set cmd1.ActiveConnection = MediConnCISPayment
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "EmailNotification"
    Set Prm1 = cmd1.CreateParameter("strAppend", adVarChar, adParamInput, 2000, StrAppend)
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("strAppendCase", adVarChar, adParamInput, 1, "2")
    cmd1.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd1.Execute
    
'    sqlstr = ""
'    sqlstr = "Select * from View_EmailRejection_Listing Where 0 = 0"
'
'        If Len(Trim(StrAppend)) Then
'            sqlstr = sqlstr + StrAppend
'        End If
'
'    rst2.Open sqlstr, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly

    Do
        Do While Not rst2.EOF
        RecFound = True
            With rst2
            
''                If Trim(!coverid) <> "00" Then
''                    RECSql = ""
''                    RECSql = " Select MBMCName FROM MBMCoveredpersons Where MBMCNumber = '" & Trim(!MBMNumber) & "' And MBMCCoverID = '" & !coverid & "'"
''                    rstMBM.Open RECSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
''
''                    Do While Not rstMBM.EOF
''                        tmpPatientName = rstMBM!MBMCName
''                        rstMBM.MoveNext
''                    Loop
''                    rstMBM.Close
''                    Set rstMBM = Nothing
''                Else
''                    tmpPatientName = !MBMName
''                End If
                Set Record = lstEmailListing.ListItems.Add(, , "")
                If Len(!MBMName) Then
                    Record.SubItems(1) = !MBMName
                End If
                Record.SubItems(2) = !MBMIcBcPp
                Record.SubItems(3) = !MBMGroupCompany
                Record.SubItems(4) = !PatientName
                Record.SubItems(5) = Format(!CLMTotalActual, "###0.00")
                Record.SubItems(6) = Format(!CLMBordxAmt, "###0.00")
                Record.SubItems(7) = !DateVisit
                Record.SubItems(8) = IIf(Len(!NPClinic), !NPClinic, "")
                Record.SubItems(9) = !CasNumber
                Record.SubItems(11) = Date
                'Record.SubItems(12) = !MBMBankACNo
                Record.SubItems(13) = !InvoiceReceiptDate
                Record.SubItems(14) = IIf(Len(!CASEmailRejStatus), !CASEmailRejStatus, "")
                Record.SubItems(15) = !MBMNumber
                Record.SubItems(16) = IIf(Len(!MBMEmail), !MBMEmail, "")
                Record.SubItems(17) = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
                Record.SubItems(18) = IIf(Len(!RepudiationReason), !RepudiationReason, "")
                Record.SubItems(19) = IIf(Len(!InvoiceNo), !InvoiceNo, "")
                Record.SubItems(20) = IIf(Len(!CASRepAction), !CASRepAction, "")
                Record.SubItems(21) = IIf(Len(!CASRepAmt), !CASRepAmt, "")
                Record.SubItems(22) = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
                Record.SubItems(23) = IIf(Len(!MBMEmployeeNo), !MBMEmployeeNo, "")
                'tmpPaymentdate = CDate(!PVDate) + 3
                
                
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
    Call ActionLogUpdate(12, "frmEmailRejectedClaim")
    cboStatus.AddItem "N - New"
    cboStatus.AddItem "R - Resend"
    Call PayorComboListing
End Sub


Private Sub lstEmailListing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(1)
    Case 2
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(2)
    Case 3
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(3)
    Case 4
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(4)
    Case 5
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(5)
    Case 6
        SortListView lstEmailListing, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstEmailListing, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstEmailListing, ColumnHeader.Index, ldtDatetime, m_blnDirection(8)
    Case 9
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(9)
    Case 10
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(10)
    Case 11
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(11)
    Case 12
        SortListView lstEmailListing, ColumnHeader.Index, ldtDatetime, m_blnDirection(12)
    Case 13
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(13)
    Case 14
        SortListView lstEmailListing, ColumnHeader.Index, ldtDatetime, m_blnDirection(14)
    Case 15
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(15)
    Case 16
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(16)
    Case 17
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(17)
    Case 18
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(18)
    Case 19
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(19)
    Case 20
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(20)
    Case 21
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(21)
    Case 22
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(22)
    Case 23
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(23)
    Case 24
        SortListView lstEmailListing, ColumnHeader.Index, ldtstring, m_blnDirection(24)
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
        'If Item.SubItems(16) = "" Then
        '    MsgBox "Please check member's email!", vbCritical
        '    Item.Checked = False
        'Else
        If Item.SubItems(14) = "Y" Then
        
            RecordSaved = MsgBox("Email has been sent to the member! Do you want to send again?", vbYesNo + vbExclamation, DialogBoxTitle)
                If RecordSaved = vbNo Then
                    Item.Checked = False
                End If
        End If
    End If
End Function

Private Sub PayorComboListing()
Const METHOD_NAME = "PayorComboListing"
On Error GoTo debugs
Dim Company As New ADODB.Recordset

  '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    cboGroupCompany.Clear
    
    Company.Open "Select GRPCoName from GRPCompany ", MediConnCISDB, adOpenKeyset, adLockOptimistic
       
        Do Until Company.EOF
            cboGroupCompany.AddItem Trim(Company!GrpCoName)
            Company.MoveNext
        Loop
        Company.Close
        Set Company = Nothing
   ' MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


