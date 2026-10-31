VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBSNPrintSummary 
   Caption         =   "BSN-Summary Location"
   ClientHeight    =   10284
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   13428
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10284
   ScaleWidth      =   13428
   WindowState     =   2  'Maximized
   Begin VB.Frame framelisting 
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   10215
      Begin MSComctlLib.ListView lstBordxTrackList 
         Height          =   3015
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   9615
         _ExtentX        =   16955
         _ExtentY        =   5313
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Bordx No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NO Of Claim"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Bordx Amt"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Bordx Date"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.CommandButton CmdPrintSummary 
      Caption         =   "&Print Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8640
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Cmdclear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7680
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6720
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5760
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdincreasebutton 
      Caption         =   "Increase Bordx &Text Box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox txtBordxrefno 
      Height          =   405
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblBordxrefno 
      Caption         =   "Bordx No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmBSNPrintSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmptop As Integer
Dim tmpid As Integer
Dim PrevTextControl As TextBox
Dim TextControl As TextBox
Dim frmvisibletop As Integer
Private Sub Cmdclear_Click()
    tmpid = 0
    tmptop = 0
    frmvisibletop = 0
    frmBSNPrintSummary.Show
    Set PrevTextControl = TextControl
End Sub
Private Sub cmdexit_Click()
Unload Me
tmpid = 0
tmptop = 0
frmvisibletop = 0
End Sub
Private Sub Form_Load()
Call ActionLogUpdate(12, "frmBSNPrintSummary")
    tmpid = 0
    tmptop = 0
    frmvisibletop = 0
End Sub
Private Sub cmdincreasebutton_Click()
    tmpid = tmpid + 1
    tmptop = tmptop + 600
    Set TextControl = frmBSNPrintSummary.Controls.Add("Vb.TextBox", "Text" & tmpid)
    If (tmpid > 1) Then
        Set PrevTextControl = frmBSNPrintSummary.Controls("Text" & (tmpid - 1))
    Else
        Set PrevTextControl = TextControl
        tmptop = 840
    End If
    With TextControl
        .Left = 1080
        .Top = tmptop
        .Width = 1935
        .Height = 405
        .Visible = True
    End With
    frmvisibletop = frmvisibletop + 600 + tmptop
End Sub
Private Sub cmdsearch_Click()
Dim rst2 As New ADODB.Recordset
Dim strAppendCase As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
StrAppend = ""
  If Len(txtBordxrefno) Then
    StrAppend = StrAppend + " And Bordx.BordxRefNo = '" & Trim(txtBordxrefno) & "'"
  End If
    
    
 If Len(txtBordxrefno) = 0 Then
        MsgBox "Please Enter Bordx Refno..! ", vbOKOnly + vbCritical, DialogBoxTitle
        txtBordxrefno.SetFocus
        Exit Sub
 Else
    lstBordxTrackList.ListItems.Clear
    strAppendCase = "1"
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "SP_Bordxdetails"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    Do
    Do Until rst2.EOF
        With rst2
            Set Record = lstBordxTrackList.ListItems.Add(, , IIf(Len(rst2!Bordxrefno), rst2!Bordxrefno, ""))
                If Len(rst2!NOOFCLAIMS) Then
                    Record.SubItems(1) = Trim(rst2!NOOFCLAIMS)
                End If
                
                If Len(rst2!CLMBordxAmt) Then
                    Record.SubItems(2) = Trim(rst2!CLMBordxAmt)
                End If
                
                If Len(rst2!BordxDate) Then
                   Record.SubItems(3) = rst2!BordxDate
                End If
        End With
        rst2.MoveNext
         DoEvents
       Loop
    Loop Until Check = False
    rst2.Close
    Set rst2 = Nothing
    If lstBordxTrackList.ListItems.Count <> 0 Then
        If frmvisibletop = 0 Then
            framelisting.Top = 840
        Else
            framelisting.Top = frmvisibletop
        End If
        framelisting.Visible = True
        cmdincreasebutton.Enabled = False
    End If

 End If
    
End Sub
