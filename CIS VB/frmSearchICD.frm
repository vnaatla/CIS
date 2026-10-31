VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchICD 
   Caption         =   "Search ICD"
   ClientHeight    =   5148
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9036
   LinkTopic       =   "Form1"
   ScaleHeight     =   5148
   ScaleWidth      =   9036
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8685
      Begin VB.TextBox txtReasonDesc 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   720
         Width           =   3375
      End
      Begin VB.ComboBox cboCriteria 
         Height          =   315
         Left            =   5040
         TabIndex        =   6
         Text            =   "AND"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtReasonDesc2 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox TxtReasonCode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label33 
         Caption         =   "Reason Desc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label32 
         Caption         =   "Reason Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Top             =   4560
      Width           =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   315
      Left            =   5040
      TabIndex        =   1
      Top             =   4560
      Width           =   840
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   315
      Left            =   4080
      TabIndex        =   0
      Top             =   4560
      Width           =   840
   End
   Begin MSComctlLib.ListView lstReasonList 
      Height          =   2715
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   4784
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ICD Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ICD Description"
         Object.Width           =   8996
      EndProperty
   End
End
Attribute VB_Name = "frmSearchICD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer
Private m_blnDirection(1 To 2) As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '--- Enable PrintScreen feature
    '--- IMPORTANT - Don't forget to set .KeyPreview = True for this form at design time!!
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
TxtReasonCode.Text = ""
txtReasonDesc.Text = ""
txtReasonDesc2.Text = ""
cboCriteria.Text = "AND"
lstReasonList.ListItems.Clear

End Sub

Private Sub Form_Activate()
    txtReasonDesc.SetFocus
End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "frmSearchICD")
    cboCriteria.AddItem ("AND")
    cboCriteria.AddItem ("OR")
End Sub


Private Sub ReasonListing(Optional strAppend As String)
Const METHOD_NAME = "ReasonListing"
On Error GoTo debugs

    Dim rst2 As New ADODB.Recordset
    Dim ReasonMaster As ListItem
    Dim sql As String
     
    lstReasonList.ListItems.Clear

        sql = " SELECT ReasonCode, ReasonDesc " & _
        " from TreatmentReason where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set ReasonMaster = lstReasonList.ListItems.Add(, , rst2!ReasonCode)
                If Len(rst2!ReasonDesc) Then
                    ReasonMaster.SubItems(1) = Trim(rst2!ReasonDesc)
                End If
                rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub cmdSearch_Click()
Const METHOD_NAME = "cmdSearch_Click"
On Error GoTo debugs

    Dim strCriteria As String
    Screen.MousePointer = vbHourglass
    strCriteria = ""
    If Len(Trim(TxtReasonCode)) Then
            strCriteria = " AND ReasonCOde like '" & Trim(TxtReasonCode) & "%'"
    End If
    If Len(Trim(txtReasonDesc)) Then
            strCriteria = strCriteria & " AND ReasonDesc like '%" & Trim(txtReasonDesc) & "%' "
    End If
    If Len(Trim(txtReasonDesc2)) Then
            strCriteria = strCriteria & cboCriteria & " ReasonDesc like '%" & Trim(txtReasonDesc2) & "%'"
    End If
    Call ReasonListing(strCriteria)
    Screen.MousePointer = vbDefault
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'**********************Start Sorted ListView *******************************
Private Sub lstReasonList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstReasonList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstReasonList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    End Select
End Sub
'**********************End Sorted ListView *******************************
Private Sub lstReasonList_DblClick()
    
    If lstReasonList.ListItems.Count <> 0 Then
        If Source = 1 Then
            frmInvoice.txtDiagnosis1 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoice.TxtDiagDesc1 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 2 Then
            frmInvoice.txtDiagnosis2 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoice.TxtDiagDesc2 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 3 Then
            frmInvoice.txtDiagnosis3 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoice.TxtDiagDesc3 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 4 Then
            frmInvoice.txtDiagnosis4 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoice.TxtDiagDesc4 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 5 Then
            frmInvoice.TxtDiagnosis5 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoice.TxtDiagDesc5 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 6 Then
            frmInvoiceMBM.txtDiagnosis1 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceMBM.TxtDiagDesc1 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 7 Then
            frmInvoiceMBM.txtDiagnosis2 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceMBM.TxtDiagDesc2 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 8 Then
            frmInvoiceMBM.txtDiagnosis3 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceMBM.TxtDiagDesc3 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 9 Then
            frmInvoiceMBM.txtDiagnosis4 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceMBM.TxtDiagDesc4 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 10 Then
            frmInvoiceMBM.TxtDiagnosis5 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceMBM.TxtDiagDesc5 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 11 Then
            frmInvoiceAdj.txtDiagnosis1 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceAdj.TxtDiagDesc1 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 12 Then
            frmInvoiceAdj.txtDiagnosis2 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceAdj.TxtDiagDesc2 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 13 Then
            frmInvoiceAdj.txtDiagnosis3 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceAdj.TxtDiagDesc3 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 14 Then
            frmInvoiceAdj.txtDiagnosis4 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceAdj.TxtDiagDesc4 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 15 Then
            frmInvoiceAdj.TxtDiagnosis5 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceAdj.TxtDiagDesc5 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
            
        ElseIf Source = 16 Then
            frmInvoiceAdjAftBordx.txtDiagnosis1 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceAdjAftBordx.TxtDiagDesc1 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 17 Then
            frmInvoiceAdjAftBordx.txtDiagnosis2 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceAdjAftBordx.TxtDiagDesc2 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 18 Then
            frmInvoiceAdjAftBordx.txtDiagnosis3 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceAdjAftBordx.TxtDiagDesc3 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 19 Then
            frmInvoiceAdjAftBordx.txtDiagnosis4 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceAdjAftBordx.TxtDiagDesc4 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        ElseIf Source = 20 Then
            frmInvoiceAdjAftBordx.TxtDiagnosis5 = lstReasonList.SelectedItem.Text '& " - " & lstReasonList.SelectedItem.SubItems(1)
            frmInvoiceAdjAftBordx.TxtDiagDesc5 = Trim(lstReasonList.SelectedItem.SubItems(1))
            Unload Me
        End If
    End If
End Sub


