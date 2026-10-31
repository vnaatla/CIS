VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSUNSystemSearchAcc 
   Caption         =   "Search Account Code"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   8745
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
      Height          =   795
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8685
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   840
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   315
         Left            =   3840
         TabIndex        =   9
         Top             =   360
         Width           =   840
      End
      Begin VB.TextBox txtReasonDesc2 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   2040
         Width           =   3375
      End
      Begin VB.ComboBox cboCriteria 
         Height          =   315
         Left            =   5040
         TabIndex        =   6
         Text            =   "AND"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtReasonDesc 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox TxtReasonCode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label33 
         Caption         =   "Reason Desc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label Label32 
         Caption         =   "Account Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstReasonList 
      Height          =   4275
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   7541
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Acc Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "AccDescription"
         Object.Width           =   8996
      EndProperty
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   315
      Left            =   3720
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   840
   End
End
Attribute VB_Name = "frmSUNSystemSearchAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer
Private m_blnDirection(1 To 2) As Boolean

Private Sub CmdSearch_Click()

    Dim strCriteria As String
    Screen.MousePointer = vbHourglass
    strCriteria = ""
    If Len(Trim(TxtReasonCode)) Then
        If Len(Trim(TxtReasonCode)) = 2 Then
            strCriteria = strCriteria & " AND substring(ACCNT_CODE,1,2) = '" & Trim(TxtReasonCode) & "'"
        ElseIf Len(Trim(TxtReasonCode)) = 3 Then
            strCriteria = strCriteria & " AND substring(ACCNT_CODE,1,3) = '" & Trim(TxtReasonCode) & "'"
        ElseIf Len(Trim(TxtReasonCode)) = 4 Then
            strCriteria = strCriteria & " AND substring(ACCNT_CODE,1,4) = '" & Trim(TxtReasonCode) & "'"
        Else
            strCriteria = strCriteria & " AND ACCNT_CODE like '%" & Trim(TxtReasonCode) & "' "
        End If
    End If
    
    medixmasterSUNSys.Open "File Name=" & BOSPath & "Conn\MediConnSUNSystem.UDL"
    
        Call ReasonListing(strCriteria)
        
    medixmasterSUNSys.Close
    Set medixmasterSUNSys = Nothing
    
    Screen.MousePointer = vbDefault
End Sub

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
    TxtReasonCode.SetFocus
End Sub

Private Sub ReasonListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim ReasonMaster As ListItem
    Dim sql As String
     
    lstReasonList.ListItems.Clear
'SUN_DB
'ACCNT_CODE
'ACCNT_NAME
    If Trim(frmSUNSystemAllocation.txtDB) = "UHC" Then
        sql = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'UHC' "
    ElseIf Trim(frmSUNSystemAllocation.txtDB) = "FW" Then
        sql = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'FW' "
    ElseIf Trim(frmSUNSystemAllocation.txtDB) = "MED" Then
        sql = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'MED' "
    ElseIf Trim(frmSUNSystemAllocation.txtDB) = "GV" Then
        sql = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'GV' "
    ElseIf Trim(frmSUNSystemAllocation.txtDB) = "TL" Then
        sql = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'TL' "
    ElseIf Trim(frmSUNSystemAllocation.txtDB) = "UHI" Then
        sql = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'UHI' "
    End If
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
    
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set ReasonMaster = lstReasonList.ListItems.Add(, , Trim(rst2!ACCNT_CODE))
                If Len(rst2!ACCNT_NAME) Then
                    ReasonMaster.SubItems(1) = Trim(rst2!ACCNT_NAME)
                End If
                rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    End If

End Sub


'**********************End Sorted ListView *******************************
Private Sub lstReasonList_DblClick()
    
    If lstReasonList.ListItems.Count <> 0 Then
        frmSUNSystemAllocation.txtACCCodeFr = lstReasonList.SelectedItem.Text
        frmSUNSystemAllocation.txtAccountname = Trim(lstReasonList.SelectedItem.SubItems(1))
        'frmInvoice.TxtDiagDesc1 = Trim(lstReasonList.SelectedItem.SubItems(1))
        Unload Me
    End If
End Sub

