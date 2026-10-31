VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCasePending 
   Caption         =   "Case Pending Options"
   ClientHeight    =   8844
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10260
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8844
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save && &Exit"
      Height          =   375
      Left            =   8400
      TabIndex        =   15
      Top             =   8400
      Width           =   1692
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10215
      _ExtentX        =   18013
      _ExtentY        =   13991
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Pending Reasons"
      TabPicture(0)   =   "frmCasePending.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Question"
      TabPicture(1)   =   "frmCasePending.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Old Pending Reason"
         Height          =   975
         Left            =   240
         TabIndex        =   13
         Top             =   6840
         Width           =   9855
         Begin VB.Label lblOldPendingReason 
            Height          =   615
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   9615
         End
      End
      Begin VB.Frame Frame4 
         Height          =   5412
         Left            =   -74760
         TabIndex        =   7
         Top             =   480
         Width           =   9612
         Begin VB.ListBox lstquestions 
            Height          =   2928
            ItemData        =   "frmCasePending.frx":0038
            Left            =   120
            List            =   "frmCasePending.frx":003A
            TabIndex        =   11
            Top             =   1920
            Width           =   9372
         End
         Begin VB.CommandButton cmdaddreason 
            Caption         =   "Add"
            Height          =   552
            Index           =   0
            Left            =   8160
            TabIndex        =   10
            Top             =   720
            Width           =   1332
         End
         Begin VB.CommandButton cmdaddreason 
            Caption         =   "Remove"
            Height          =   552
            Index           =   1
            Left            =   8160
            TabIndex        =   9
            Top             =   1260
            Width           =   1332
         End
         Begin VB.TextBox txtquestion 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   120
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   240
            Width           =   9372
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Pending Reasons"
         Height          =   2655
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   9855
         Begin MSComctlLib.ListView lstpendinglist 
            Height          =   2292
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   9612
            _ExtentX        =   16955
            _ExtentY        =   4043
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Check"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Pending Description"
               Object.Width           =   9349
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Pending ID"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Selected Pending Reasons"
         Height          =   3375
         Left            =   240
         TabIndex        =   1
         Top             =   3360
         Width           =   9855
         Begin VB.TextBox txtAddPendRemark 
            Height          =   615
            Left            =   1560
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   2640
            Width           =   8175
         End
         Begin VB.Label Label1 
            Caption         =   "Additional Remark"
            Height          =   495
            Left            =   120
            TabIndex        =   4
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label lblSelectedReasons 
            BackColor       =   &H00DCEFF3&
            Height          =   2295
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   9615
         End
      End
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Reset"
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   8400
      Width           =   1092
   End
   Begin VB.Label lblsplremark 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   10212
   End
End
Attribute VB_Name = "frmCasePending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public selectedPendingReasons As String

Dim varpipe() As String

Private Sub cmdaddreason_Click(Index As Integer)
    If cmdaddreason(0) Then
        Dim chkdup As Boolean
        chkdup = False
        Dim i As Long
        For i = 1 To lstquestions.ListCount
        If txtquestion.Text = lstquestions.List(i - 1) Then
            chkdup = True
            txtquestion.Text = ""
        Exit For
        End If
        Next
        If chkdup = False And Len(txtquestion.Text) >= 1 Then
              lstquestions.AddItem (txtquestion.Text)
              txtquestion.Text = ""
        End If
    ElseIf cmdaddreason(1) Then
        If (lstquestions.SelCount > 0) Then
            lstquestions.RemoveItem (lstquestions.ListIndex)
        End If
    End If
End Sub

Private Sub cmdcancel_Click()
Const METHOD_NAME = "cmdcancel_Click"
On Error GoTo debugs
Dim i As Integer
Dim j As Integer
Dim arrexistpendIds() As String

If (MsgBox("Your current selection of Pending reasons and Questions will be reset to original values, are you sure?", vbYesNo + vbExclamation)) = vbYes Then
    lstquestions.Clear
    If Len(frmInvoiceAdj.existstrQAtoHos) Then
        Dim strQAarr() As String
        strQAarr = Split(strQAtoHos, "|")
        Dim strQAItem As Integer
        For strQAItem = 0 To UBound(strQAarr())
            lstquestions.AddItem (strQAarr(strQAItem))
        Next
        frmInvoiceAdj.existstrQAtoHos = strQAtoHos
    Else
        frmInvoiceAdj.existstrQAtoHos = ""
        lstquestions.Clear
    End If

    If IsNumeric(Replace(frmInvoiceAdj.existPendingReasonID, ",", "")) Then
        For i = 1 To frmCasePending.lstpendinglist.ListItems.Count
            frmCasePending.lstpendinglist.ListItems(i).Checked = False
        Next i
        lblSelectedReasons = ""
        If Len(frmInvoiceAdj.existPendingReasonID) > 0 Then
            If InStr(frmInvoiceAdj.existPendingReasonID, ",") Then
                arrexistpendIds = Split(frmInvoiceAdj.existPendingReasonID, ",")
            Else
                ReDim arrexistpendIds(0)
                arrexistpendIds(0) = frmInvoiceAdj.existPendingReasonID
            End If
            For i = 1 To frmCasePending.lstpendinglist.ListItems.Count
                For j = 0 To UBound(arrexistpendIds)
                 If (frmCasePending.lstpendinglist.ListItems(i).SubItems(2) = arrexistpendIds(j)) Then
                    frmCasePending.lstpendinglist.ListItems(i).Checked = True
                    frmCasePending.selectedPendingReasons = frmCasePending.lblSelectedReasons + IIf(Len(frmCasePending.lblSelectedReasons) > 0, vbNewLine, "") + frmCasePending.lstpendinglist.ListItems(i).SubItems(1)
                    frmCasePending.lblSelectedReasons = frmCasePending.selectedPendingReasons
                 End If
                Next
            Next
        End If
    Else
         For i = 1 To frmCasePending.lstpendinglist.ListItems.Count
             frmCasePending.lstpendinglist.ListItems(i).Checked = False
         Next i
         lblSelectedReasons = ""
         lblOldPendingReason = ""
    End If
    
    If Len(frmInvoiceAdj.ExistAddPendRemark) <> 0 Then
        txtAddPendRemark = frmInvoiceAdj.ExistAddPendRemark
    Else
        txtAddPendRemark = ""
    End If
    
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdSave_Click()
Const METHOD_NAME = "CmdSave_Click"
On Error GoTo debugs
    strQAtoHos = ""
    If Len(Trim(txtquestion)) And cmdaddreason(0).Enabled = True Then
        MsgBox "You have typed the question but not added to the list. Please add it or remove it!", vbInformation
        txtquestion.SetFocus
        Exit Sub
    ElseIf lstquestions.ListCount > 0 Then
        For i = 0 To lstquestions.ListCount - 1
            strQAtoHos = strQAtoHos + lstquestions.List(i) + "|"
        Next
        strQAtoHos = Mid(strQAtoHos, 1, Len(strQAtoHos) - 1)
    Else
        strQAtoHos = ""
    End If

    
    strAddPendRemark = Trim(txtAddPendRemark)
    
    strPendingReasonId = ""
        For j = 1 To lstpendinglist.ListItems.Count
            If lstpendinglist.ListItems(j).Checked = True Then
                strPendingReasonId = strPendingReasonId + lstpendinglist.ListItems(j).SubItems(2) + ","
            End If
        Next
        
        If Len(strPendingReasonId) > 0 Then
            strPendingReasonId = Mid(strPendingReasonId, 1, Len(strPendingReasonId) - 1)
        End If
        
    frmCasePending.Visible = False
    txtquestion = ""
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub Form_Load()

Call ActionLogUpdate(12, "frmCasePending")
SSTab1.Tab = 0
txtAddPendRemark = strAddPendRemark
End Sub



Private Sub lstpendinglist_ItemCheck(ByVal Item As MSComctlLib.ListItem)
selectedPendingReasons = ""
lblSelectedReasons = ""
Dim i As Integer
For i = 1 To lstpendinglist.ListItems.Count
    
    
            If lstpendinglist.ListItems(i).Checked = True Then
                selectedPendingReasons = lblSelectedReasons + IIf(Len(lblSelectedReasons) > 0, vbNewLine, "") + lstpendinglist.ListItems(i).SubItems(1)
                lblSelectedReasons = selectedPendingReasons
            Else
                lstpendinglist.ListItems(i).Checked = False
            End If
        Next
End Sub




