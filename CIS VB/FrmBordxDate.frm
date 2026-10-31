VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBordxDate 
   Caption         =   "ChangeBordxDate"
   ClientHeight    =   6624
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   11688
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6624
   ScaleWidth      =   11688
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   11415
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   10200
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   8880
         MaskColor       =   &H80000012&
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   7320
         MaskColor       =   &H80000012&
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11565
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "BordxDate Module"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   3525
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11415
      Begin VB.TextBox txtBordxNo 
         Height          =   350
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   1
         Top             =   357
         Width           =   2775
      End
      Begin MSMask.MaskEdBox txtBordxDateC 
         Height          =   345
         Left            =   5040
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   614
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Bordx No:*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   435
         Width           =   810
      End
      Begin VB.Label Label30 
         Caption         =   "BordxDate"
         Height          =   255
         Left            =   4080
         TabIndex        =   3
         Top             =   405
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstBordxRefNo 
      Height          =   3735
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   11415
      _ExtentX        =   20130
      _ExtentY        =   6583
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "BordxRefNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "BordxDate"
         Object.Width           =   5469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Postingstatus"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "NoOfCases"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FrmBordxDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClnAccessMaster As ListItem
Dim postingstatus As Boolean

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Const METHOD_NAME = "Form_Load"
On Error GoTo debugs

Call ActionLogUpdate(12, "FrmBordxDate")


debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub CmdSearch_Click()
Screen.MousePointer = vbHourglass

    lstBordxRefNo.ListItems.Clear
    Call SearchBordxDate
End Sub

Private Sub SearchBordxDate()
    Screen.MousePointer = vbHourglass
    Call BordxDateListing(Trim(txtBordxNo))
    Screen.MousePointer = Default
End Sub


Private Sub BordxDateListing(strAppend As String)
Const METHOD_NAME = "BordxDateListing"
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim ClnAccessMaster As ListItem
Dim TotalCACount As Integer
    
     postingstatus = False
    Set cmd.ActiveConnection = MediConnCISWS
                cmd.CommandText = "BordxDateChange"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd.Parameters.Append Prm1
                rst2.CursorLocation = adUseClient
                rst2.CursorType = adOpenForwardOnly
                rst2.CacheSize = 100
                Set rst2 = cmd.Execute
                
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
    
    Do While Not rst2.EOF
        TotalCACount = 0
            Set ClnAccessMaster = lstBordxRefNo.ListItems.Add(, , rst2!Bordxrefno)
                ClnAccessMaster.SubItems(1) = rst2!BordxDate
                ClnAccessMaster.SubItems(2) = rst2!postingstatus
                    If rst2!postingstatus = "YES" Then
                        postingstatus = True
                    End If
                ClnAccessMaster.SubItems(3) = rst2!NoOfCases
                BordxDate = rst2!BordxDate
                rst2.MoveNext
        TotalCACount = TotalCACount + 1
        txtTotalRecord = TotalCACount
        Loop
    rst2.Close
    End If
    
    If postingstatus = True Then
        MsgBox "Posting is Done. Update is not allowed", vbOKOnly + vbInformation, DialogBoxTitle
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdUpdate_Click()
Const METHOD_NAME = "CmdUpdate_Click"
On Error GoTo debugs

    Screen.MousePointer = vbHourglass
    Call UpdateBordxDate
    Screen.MousePointer = vbDefault
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub UpdateBordxDate()
    Dim StrSql As String
    StrSql = ""
    
    If txtBordxDateC = "__/__/____" Then
            MsgBox "Please Enter Bordx Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtBordxDateC.SetFocus
            Exit Sub
    ElseIf postingstatus = True Then
            MsgBox "Posting is Done. Update is not allowed", vbOKOnly + vbInformation, DialogBoxTitle
            Exit Sub
    Else
            StrSql = "Update CISWorksheet..Bordx SET BordxBy='" & ChkStr(Logon) & "',BordxDate = '" & Format(txtBordxDateC, "yyyy/mm/dd") & "' WHERE BordxRefNo = '" & Trim(txtBordxNo) & "'"
            MediConnCISWS.Execute StrSql
            StrSql = ""
            StrSql = "Update CISWorksheet..Claim set " & _
                     " CLMBordxMonth = '" & Format(txtBordxDateC, "MM") & "'," & _
                     " CLMBordxYear = '" & Format(txtBordxDateC, "yyyy") & "'" & _
                     " WHERE claim.BordxRefNo = '" & Trim(txtBordxNo) & "'"
            MediConnCISWS.Execute StrSql
            MsgBox "BordxDate Updated Successfully..!!", vbOKOnly + vbInformation, DialogBoxTitle
    End If
End Sub
