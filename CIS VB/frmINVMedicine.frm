VERSION 5.00
Begin VB.Form frmINVMedicine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Claim Invoice Actual Others"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   3300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   3135
      Begin VB.CommandButton CmdOK 
         Caption         =   "&OK"
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   285
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.TextBox txtMEdication1 
         Height          =   615
         Left            =   480
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtMEdication2 
         Height          =   615
         Left            =   480
         TabIndex        =   9
         Top             =   1200
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtMEdication3 
         Height          =   615
         Left            =   480
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox CboProc7 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   7
         Top             =   2040
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.ComboBox CboProc6 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   6
         Top             =   1800
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.ComboBox CboProc5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   5
         Top             =   1560
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.ComboBox CboProc4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.ComboBox CboProc3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.ComboBox CboProc2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.ComboBox CboProc1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.Label lblINvOthTitle 
         Caption         =   "Medication"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmINVMedicine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer

Private Sub Cboproc4_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboProc4.Text, CboProc4.SelStart) + Chr(KeyAscii) + Mid(CboProc4.Text, CboProc4.SelStart + CboProc4.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboProc4.ListCount - 1
       
        If NewText = LCase(Left(CboProc4.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboProc4.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboProc4.Text = ValidValue
            CboProc4.SelStart = 0
            CboProc4.SelLength = Len(ValidValue)
        End If
    End If
End Sub
Private Sub Cboproc5_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboProc5.Text, CboProc5.SelStart) + Chr(KeyAscii) + Mid(CboProc5.Text, CboProc5.SelStart + CboProc5.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboProc5.ListCount - 1
       
        If NewText = LCase(Left(CboProc5.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboProc5.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboProc5.Text = ValidValue
            CboProc5.SelStart = 0
            CboProc5.SelLength = Len(ValidValue)
        End If
    End If
End Sub
Private Sub Cboproc6_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboProc6.Text, CboProc6.SelStart) + Chr(KeyAscii) + Mid(CboProc6.Text, CboProc6.SelStart + CboProc6.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboProc6.ListCount - 1
       
        If NewText = LCase(Left(CboProc6.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboProc6.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboProc6.Text = ValidValue
            CboProc6.SelStart = 0
            CboProc6.SelLength = Len(ValidValue)
        End If
    End If
End Sub
Private Sub Cboproc7_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboProc7.Text, CboProc7.SelStart) + Chr(KeyAscii) + Mid(CboProc7.Text, CboProc7.SelStart + CboProc7.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboProc7.ListCount - 1
       
        If NewText = LCase(Left(CboProc7.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboProc7.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboProc7.Text = ValidValue
            CboProc7.SelStart = 0
            CboProc7.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboProc1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboProc1.Text, CboProc1.SelStart) + Chr(KeyAscii) + Mid(CboProc1.Text, CboProc1.SelStart + CboProc1.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboProc1.ListCount - 1
       
        If NewText = LCase(Left(CboProc1.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboProc1.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboProc1.Text = ValidValue
            CboProc1.SelStart = 0
            CboProc1.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboProc2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboProc2.Text, CboProc2.SelStart) + Chr(KeyAscii) + Mid(CboProc2.Text, CboProc2.SelStart + CboProc2.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboProc2.ListCount - 1
       
        If NewText = LCase(Left(CboProc2.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboProc2.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboProc2.Text = ValidValue
            CboProc2.SelStart = 0
            CboProc2.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboProc3_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboProc3.Text, CboProc3.SelStart) + Chr(KeyAscii) + Mid(CboProc3.Text, CboProc3.SelStart + CboProc3.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboProc3.ListCount - 1
       
        If NewText = LCase(Left(CboProc3.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboProc3.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboProc3.Text = ValidValue
            CboProc3.SelStart = 0
            CboProc3.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CmdClear_Click()
Dim RecordExit
    RecordExit = MsgBox("Are you sure ?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
    End If
End Sub


Private Sub CmdOK_Click()
    Me.Visible = False
End Sub
Private Sub Form_Load()
Dim ProcRec As New ADODB.Recordset
 
    CboProc1.Clear
    CboProc2.Clear
    CboProc3.Clear
    CboProc4.Clear
    CboProc5.Clear
    CboProc6.Clear
    CboProc7.Clear
    ProcRec.Open "SELECT ProcedureCode, ProcedureDesc FROM Procedures", MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do Until ProcRec.EOF
        CboProc1.AddItem ProcRec!ProcedureCode & " - " & ProcRec!ProcedureDesc
        CboProc2.AddItem ProcRec!ProcedureCode & " - " & ProcRec!ProcedureDesc
        CboProc3.AddItem ProcRec!ProcedureCode & " - " & ProcRec!ProcedureDesc
        CboProc4.AddItem ProcRec!ProcedureCode & " - " & ProcRec!ProcedureDesc
        CboProc5.AddItem ProcRec!ProcedureCode & " - " & ProcRec!ProcedureDesc
        CboProc6.AddItem ProcRec!ProcedureCode & " - " & ProcRec!ProcedureDesc
        CboProc7.AddItem ProcRec!ProcedureCode & " - " & ProcRec!ProcedureDesc
        ProcRec.MoveNext
    Loop
    ProcRec.Close
    Set ProcRec = Nothing
End Sub

Private Sub Form_Activate()
    
    If Source = 1 Or Source = 3 Then
        lblINvOthTitle = "Medication"
        Call EntriesMedication(True, False)
    ElseIf Source = 2 Or Source = 4 Then
        lblINvOthTitle = "Procedures"
        Call EntriesMedication(False, True)
    End If
End Sub

Private Sub EntriesMedication(MEdStatus As Boolean, ProcStat As Boolean)
    
    txtMEdication1.Visible = MEdStatus
    txtMEdication2.Visible = MEdStatus
    txtMEdication3.Visible = MEdStatus
    
    CboProc1.Visible = ProcStat
    CboProc2.Visible = ProcStat
    CboProc3.Visible = ProcStat
    CboProc4.Visible = ProcStat
    CboProc5.Visible = ProcStat
    CboProc6.Visible = ProcStat
    CboProc7.Visible = ProcStat
    
End Sub


