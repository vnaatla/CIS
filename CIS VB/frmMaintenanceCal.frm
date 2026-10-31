VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaintenanceCal 
   Caption         =   "Calculation Maintenance"
   ClientHeight    =   8004
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11076
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8004
   ScaleWidth      =   11076
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11805
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Calculation Maintenance"
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
         TabIndex        =   16
         Top             =   120
         Width           =   3285
      End
   End
   Begin VB.Frame Frame10 
      Height          =   705
      Left            =   360
      TabIndex        =   0
      Top             =   6720
      Width           =   11055
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   315
         Left            =   3240
         TabIndex        =   5
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   720
         TabIndex        =   6
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   8
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   8160
         TabIndex        =   4
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   9000
         TabIndex        =   11
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Default         =   -1  'True
         Height          =   315
         Left            =   7320
         TabIndex        =   10
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   4080
         TabIndex        =   9
         Top             =   285
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7005
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   11685
      _ExtentX        =   20616
      _ExtentY        =   12361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Co-Payment Cal."
      TabPicture(0)   =   "frmMaintenanceCal.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label21"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstCoPaymentList"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Fee Payable 2 Clinic"
      TabPicture(1)   =   "frmMaintenanceCal.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label31"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstFeeList"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   -74760
         TabIndex        =   36
         Top             =   360
         Width           =   11175
         Begin VB.TextBox txtFeeCode 
            Height          =   285
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtFeeDesc 
            Height          =   600
            Left            =   2160
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   500
            Width           =   8685
         End
         Begin VB.TextBox txtFeeRemark 
            Height          =   960
            Left            =   2160
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   1070
            Width           =   8685
         End
         Begin VB.Label Label6 
            Caption         =   "Remark                     *"
            Height          =   255
            Left            =   405
            TabIndex        =   39
            Top             =   1200
            Width           =   1650
         End
         Begin VB.Label Label5 
            Caption         =   "Fee Code            *"
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
            Left            =   360
            TabIndex        =   38
            Top             =   285
            Width           =   1740
         End
         Begin VB.Label Label2 
            Caption         =   "Description                 *"
            Height          =   255
            Left            =   360
            TabIndex        =   37
            Top             =   645
            Width           =   1695
         End
      End
      Begin VB.Frame Frame 
         Height          =   1215
         Left            =   -74640
         TabIndex        =   24
         Top             =   960
         Width           =   6375
         Begin VB.TextBox txtOCPcode1 
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
            Left            =   1995
            MaxLength       =   2
            TabIndex        =   26
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtOCPDesc1 
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
            Left            =   1995
            MaxLength       =   50
            TabIndex        =   25
            Top             =   675
            Width           =   3300
         End
         Begin VB.Label Label1 
            Caption         =   "Occupation Code"
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
            Left            =   255
            TabIndex        =   28
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label Label4 
            Caption         =   "Occupation Description"
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
            TabIndex        =   27
            Top             =   675
            Width           =   1770
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   240
         TabIndex        =   18
         Top             =   420
         Width           =   11175
         Begin VB.TextBox txtCoPayment 
            Height          =   285
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   1
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtCoDesc 
            Height          =   600
            Left            =   2160
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   500
            Width           =   8685
         End
         Begin VB.TextBox txtCoRemark 
            Height          =   960
            Left            =   2160
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1070
            Width           =   8685
         End
         Begin VB.Label Label14 
            Caption         =   "Description                 *"
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   645
            Width           =   1695
         End
         Begin VB.Label Label15 
            Caption         =   "Co-Payment Code *"
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
            Left            =   360
            TabIndex        =   20
            Top             =   285
            Width           =   1740
         End
         Begin VB.Label Label16 
            Caption         =   "Remark                     *"
            Height          =   255
            Left            =   405
            TabIndex        =   19
            Top             =   1200
            Width           =   1650
         End
      End
      Begin MSComctlLib.ListView lstCoPaymentList 
         Height          =   3180
         Left            =   240
         TabIndex        =   22
         Top             =   2880
         Width           =   11115
         _ExtentX        =   19600
         _ExtentY        =   5609
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Co-Payment Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Remark"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView lstOccupationList1 
         Height          =   3120
         Left            =   -74640
         TabIndex        =   23
         Top             =   2640
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   5503
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
            Text            =   "Occupation Code"
            Object.Width           =   2611
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descriptions"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstFeeList 
         Height          =   3240
         Left            =   -74760
         TabIndex        =   29
         Top             =   2880
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   5715
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fee Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fee Description"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Remark"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "CO-PAYMENT LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   240
         TabIndex        =   32
         Top             =   2640
         Width           =   1515
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "OCCUPATIOP LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74640
         TabIndex        =   31
         Top             =   2400
         Width           =   1635
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "FEE LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74760
         TabIndex        =   30
         Top             =   2580
         Width           =   855
      End
   End
   Begin VB.Label Label18 
      Caption         =   "Bold - Searchable        * - Mandatory Fields"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   7680
      Width           =   4215
   End
   Begin VB.Label Label32 
      Caption         =   "TOTAL RECORDS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5160
      TabIndex        =   34
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label txtTotalRecord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   33
      Top             =   7680
      Width           =   1815
   End
End
Attribute VB_Name = "frmMaintenanceCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim EditFlag As Boolean
Public COPCriteria As String
Public FEECriteria As String
Private m_blnDirection(1 To 4) As Boolean
Private Sub Form_Activate()
    SSTab1.Tab = 0
End Sub

Private Sub cmdAdd_Click()
    EditFlag = False
    Call ClearForm(Me)
    Call DisableCmdButton(Me)
    Call EntriesEnable(True)
    CmdSave.Enabled = True
    CmdAdd.Enabled = True
   
    Select Case SSTab1.Tab
        Case 0
            txtCoPayment.SetFocus
        Case 1
            txtFeeCode.SetFocus
    End Select
End Sub
Private Sub Form_Load()
Const METHOD_NAME = "Form_Load"
On Error GoTo debugs

Call ActionLogUpdate(12, "frmMaintenanceCal")


debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    Call EntriesEnable(True)
    CmdSave.Enabled = True
    CmdDelete.Enabled = False
     
    If txtCoPayment <> "" Then
       txtCoPayment.Enabled = False
       txtCoDesc.Enabled = True
       txtCoDesc.SetFocus
       txtCoRemark.Enabled = True
       txtCoRemark.SetFocus
    End If

    If txtFeeCode <> "" Then
       txtFeeCode.Enabled = False
       txtFeeDesc.Enabled = True
       txtFeeDesc.SetFocus
       txtFeeRemark.Enabled = True
       txtFeeRemark.SetFocus
    End If

End Sub

Private Sub cmdprint_Click()
    Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call PrintCOP
        Case 1
            Call PrintFEE
     End Select
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    Screen.MousePointer = Default
End Sub

Private Sub cmdSave_Click()
  Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call SaveCOP
        Case 1
            Call SaveFEE
    End Select
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False

    Screen.MousePointer = Default
End Sub

Private Sub cmdDelete_Click()
Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
       Case 0
           Call DeleteCOP
       Case 1
           Call DeleteFEE
    End Select
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
       
    Screen.MousePointer = Default
End Sub

Private Sub cmdSearch_Click()
Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call SearchCOP
            Call ClearForm(Me)
            txtCoPayment.Enabled = True
            txtCoDesc.Enabled = True
            txtCoRemark.Enabled = True
            txtCoPayment.SetFocus
        Case 1
            Call SearchFEE
            Call ClearForm(Me)
            txtFeeCode.Enabled = True
            txtFeeDesc.Enabled = True
            txtFeeCode.SetFocus
        End Select
Screen.MousePointer = Default
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    CmdAdd.Enabled = True
    lstCoPaymentList.ListItems.Clear
    lstFeeList.ListItems.Clear
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    Call EntriesEnable(True)
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub


'************************Start EntriesEnable*********************************

Private Sub EntriesEnable(status As Boolean)
     
    txtCoPayment.Enabled = status
    txtCoDesc.Enabled = status
    txtCoRemark.Enabled = status
    txtFeeCode.Enabled = status
    txtFeeDesc.Enabled = status
    CmdSearch.Enabled = status
    
End Sub
'************************End EntriesEnable*********************************
    
'####### @@@ START SAVE, EDIT, ADD, DELETE CO-PAYMENT @@@ #######'

Private Sub DeleteCOP()
Const METHOD_NAME = "DeleteCOP"
On Error GoTo debugs

Dim rsDeleteCOP As New ADODB.Recordset
Dim rsCheckCOP As New ADODB.Recordset
Dim strsql, sqlCOP As String
Dim DeleteRecord
    
EditFlag = False
       
If txtCoPayment = "" Then
   MsgBox "Please Enter Co-Payment Code", vbOKOnly + vbCritical, DialogBoxTitle
ElseIf lstCoPaymentList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlCOP = "Select COPCode From CoPayment where COPCode = '" & Trim(lstCoPaymentList.SelectedItem.Text) & "'"
    
'   rsCheckCOP.MaxRecords = 1
'   rsCheckCOP.Open sqlCOP, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
'    If rsCheckCOP.RecordCount Then
'        MsgBox "This record cannot be deleted because it is used in Calculation Table!", vbCritical, "Error in deleting record..."
'        Call ClearForm(Me)
'        CmdDelete.Enabled = False
'        CmdEdit.Enabled = False
    
'    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
             strsql = "Delete From CoPayment Where COPCode = '" & ChkStr(lstCoPaymentList.SelectedItem.Text) & "'"
             rsDeleteCOP.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call ClearForm(Me)
             Call COPListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             txtCoPayment.SetFocus
        Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

    
Private Sub SearchCOP()
Dim strAppend As String
    
  Screen.MousePointer = vbHourglass
  strAppend = ""
  
  If Len(Trim(txtCoPayment)) Then
    strAppend = strAppend + " And COPCode = '" & Trim(Mid(txtCoPayment, 1, 3)) & "'"
  'HLTCode like '%" & Mid(ChkStr(txtHLTcode), 1, 2) & "%'"
  End If
    
  COPCriteria = strAppend
  Call COPListing(strAppend)
  Screen.MousePointer = Default
  
End Sub

Private Sub PrintCOP()
    RptCoPaymentList.Show
End Sub

Private Sub SaveCOP()
Const METHOD_NAME = "SaveCOP"
On Error GoTo debugs
   On Error Resume Next
   Dim RecordSaved
   Dim COP As New ADODB.Recordset
   Dim SaveCOP As New ADODB.Recordset
   Dim strsql As String
   
    If txtCoPayment = "" Then
        MsgBox "Please Enter Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtCoPayment.SetFocus
    ElseIf Len(txtCoPayment) < 3 Then
        MsgBox "Co-Payment Code must be 3 integer!", vbOKOnly + vbCritical, DialogBoxTitle
        txtCoPayment.SetFocus
    ElseIf txtCoDesc = "" Then
        MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtCoDesc.SetFocus
'    ElseIf txtCoRemark = "" Then
'        MsgBox "Please Enter Remarks", vbOKOnly + vbCritical, DialogBoxTitle
'        txtCoRemark.SetFocus
    Else
        strsql = "Select * From CoPayment where COPCode =  '" & Trim(txtCoPayment) & "'"
        COP.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If COP.RecordCount > 0 And EditFlag = False Then
           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           COP.Close
           txtCoPayment.SetFocus
           txtCoPayment.SelStart = 0
           txtCoPayment.SelLength = Len(txtCoPayment)
           Set COP = Nothing
           Exit Sub
        Else
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           If RecordSaved = vbYes Then
              If EditFlag = False Then
                 SaveCOP.Open "CoPayment", MediConnCISDB, adOpenKeyset, adLockOptimistic
                 SaveCOP.AddNew
              Else
                 strsql = "Select * From CoPayment Where COPCode = '" & ChkStr(txtCoPayment) & "'"
                 SaveCOP.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
              End If
              
              With SaveCOP
                !COPCode = ChkStr(txtCoPayment)
                !COPDesc = ChkStr(txtCoDesc)
                !COPRemark = ChkStr(txtCoRemark)
                !COPLastUpdateUser = Logon
                !COPLastUpdateDate = Date
                
              End With
              SaveCOP.update
            
             Call ClearForm(Me)
             'Call SearchCOP
             Call COPListing(COPCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             COP.Close
             Set SaveCOP = Nothing
        End If
        EditFlag = False
     End If
   End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub COPListing(Optional strAppend As String)
Const METHOD_NAME = "COPListing"
On Error GoTo debugs
 'On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim COP As ListItem
    Dim sql As String
    Dim TotalCOPCount As Integer
   
    TotalCOPCount = 0
    lstCoPaymentList.ListItems.Clear
    
    sql = " select COPCode, COPDesc, COPRemark " & _
      " From CoPayment where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            With rst2
            Set COP = lstCoPaymentList.ListItems.Add(, , rst2!COPCode)
                COP.SubItems(1) = rst2!COPDesc
                COP.SubItems(2) = rst2!COPRemark
                rst2.MoveNext
            End With
        TotalCOPCount = TotalCOPCount + 1
        txtTotalRecord = TotalCOPCount
        Loop
    End If
    COPCriteria = ""
    rst2.Close
    Set rst2 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstCoPaymentList_Click()
Dim strsql As String
    
    If lstCoPaymentList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtCoPayment = lstCoPaymentList.SelectedItem.Text
        txtCoPayment.Enabled = False
        txtCoDesc = lstCoPaymentList.SelectedItem.SubItems(1)
        txtCoDesc.Enabled = False
        txtCoRemark = lstCoPaymentList.SelectedItem.SubItems(2)
        txtCoRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstCoPaymentList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstCoPaymentList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtCoPayment = lstCoPaymentList.SelectedItem.Text
        txtCoPayment.Enabled = False
        txtCoDesc = lstCoPaymentList.SelectedItem.SubItems(1)
        txtCoDesc.Enabled = False
        txtCoRemark = lstCoPaymentList.SelectedItem.SubItems(2)
        txtCoRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
    

End Sub

Private Sub lstCoPaymentList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstCoPaymentList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtCoPayment = lstCoPaymentList.SelectedItem.Text
        txtCoPayment.Enabled = False
        txtCoDesc = lstCoPaymentList.SelectedItem.SubItems(1)
        txtCoDesc.Enabled = False
        txtCoRemark = lstCoPaymentList.SelectedItem.SubItems(2)
        txtCoRemark.Enabled = False
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
    

End Sub

Private Sub lstCoPaymentList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstCoPaymentList.SortKey = ColumnHeader.Index - 1
    lstCoPaymentList.Sorted = True
    If lstCoPaymentList.SortOrder = lvwAscending Then
       lstCoPaymentList.SortOrder = lvwDescending
    Else
       lstCoPaymentList.SortOrder = lvwAscending
    End If
    
    txtCoPayment.Enabled = False
End Sub
'####### @@@ END SAVE, EDIT , ADD, DELETE CO-PAYMENT @@@ #######'

'&&&&&&& *** START SAVE, EDIT, ADD, DELETE FEE *** &&&&&&&'

Private Sub DeleteFEE()
Const METHOD_NAME = "DeleteFEE"
On Error GoTo debugs
Dim rsDeleteFEE As New ADODB.Recordset
Dim rsCheckFEE As New ADODB.Recordset
Dim strsql, sqlFEE As String
Dim DeleteRecord
    
EditFlag = False
       
If txtFeeCode = "" Then
   MsgBox "Please Enter Fee Code", vbOKOnly + vbCritical, DialogBoxTitle
ElseIf lstFeeList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlFEE = "Select FeeCode From FEE where FeeCode = '" & Trim(lstFeeList.SelectedItem.Text) & "'"
    
'   rsCheckFEE.MaxRecords = 1
'   rsCheckFEE.Open sqlFEE, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
'    If rsCheckFEE.RecordCount Then
'        MsgBox "This record cannot be deleted because it is used in FEE Table!", vbCritical, "Error in deleting record..."
'        Call ClearForm(Me)
'        CmdDelete.Enabled = False
'        CmdEdit.Enabled = False
    
'    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
        'And rsCheckCOP.RecordCount = 0 Then
             strsql = "Delete From FEE Where FeeCode = '" & ChkStr(lstFeeList.SelectedItem.Text) & "'"
             rsDeleteFEE.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call ClearForm(Me)
             Call FEEListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             txtFeeCode.SetFocus
        Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
    
Private Sub SearchFEE()
Const METHOD_NAME = "SearchFEE"
On Error GoTo debugs

Dim strAppend As String
    
  Screen.MousePointer = vbHourglass
  strAppend = ""
  
  If Len(Trim(txtFeeCode)) Then
  strAppend = strAppend + " And FeeCode = '" & Trim(Mid(txtFeeCode, 1, 3)) & "'"
  'HLTCode like '%" & Mid(ChkStr(txtHLTcode), 1, 2) & "%'"
  End If
    
  FEECriteria = strAppend
  Call FEEListing(strAppend)
  Screen.MousePointer = Default
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub PrintFEE()
    RptFeeList.Show
End Sub

Private Sub SaveFEE()
Const METHOD_NAME = "SaveFEE"
On Error GoTo debugs

   On Error Resume Next
   Dim RecordSaved
   Dim FEE As New ADODB.Recordset
   Dim SaveFEE As New ADODB.Recordset
   Dim strsql As String
   
    If txtFeeCode = "" Then
        MsgBox "Please Enter Fee Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtFeeCode.SetFocus
    ElseIf Len(txtFeeCode) < 3 Then
        MsgBox "Fee Code must be 3 integer!", vbOKOnly + vbCritical, DialogBoxTitle
        txtFeeCode.SetFocus
    ElseIf txtFeeDesc = "" Then
        MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtFeeDesc.SetFocus
    Else
        strsql = "Select * From FEE where FeeCode =  '" & Trim(txtFeeCode) & "'"
        FEE.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If FEE.RecordCount > 0 And EditFlag = False Then
           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           FEE.Close
           txtFeeCode.SetFocus
           txtFeeCode.SelStart = 0
           txtFeeCode.SelLength = Len(txtFeeCode)
           Set FEE = Nothing
           Exit Sub
        Else
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           If RecordSaved = vbYes Then
              If EditFlag = False Then
                 SaveFEE.Open "FEE", MediConnCISDB, adOpenKeyset, adLockOptimistic
                 SaveFEE.AddNew
              Else
                 strsql = "Select * From FEE Where FeeCode = '" & ChkStr(txtFeeCode) & "'"
                 SaveFEE.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
              End If
              
              With SaveFEE
                !FeeCode = ChkStr(txtFeeCode)
                !FeeDesc = ChkStr(txtFeeDesc)
                !FeeRemark = ChkStr(txtFeeRemark)
                !COPLastUpdateUser = Logon
                !COPLastUpdateDate = Date
                
              End With
              SaveFEE.update
            
             Call ClearForm(Me)
             Call FEEListing(FEECriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             FEE.Close
             Set SaveFEE = Nothing
        End If
        EditFlag = False
     End If
   End If
   
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub FEEListing(Optional strAppend As String)
Const METHOD_NAME = "FEEListing"
On Error GoTo debugs
 'On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim FEE As ListItem
    Dim sql As String
    Dim TotalFEECount As Integer
   
    TotalFEECount = 0
    lstFeeList.ListItems.Clear
    
    sql = " select FeeCode, FeeDesc, FeeRemark " & _
      " From FEE where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            With rst2
            Set FEE = lstFeeList.ListItems.Add(, , rst2!FeeCode)
                FEE.SubItems(1) = rst2!FeeDesc
                FEE.SubItems(2) = rst2!FeeRemark
                rst2.MoveNext
            End With
        TotalFEECount = TotalFEECount + 1
        txtTotalRecord = TotalFEECount
        Loop
    End If
    FEECriteria = ""
    rst2.Close
    Set rst2 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub lstFeeList_Click()
Dim strsql As String
    
    If lstFeeList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtFeeCode = lstFeeList.SelectedItem.Text
        txtFeeCode.Enabled = False
        txtFeeDesc = lstFeeList.SelectedItem.SubItems(1)
        txtFeeDesc.Enabled = False
        CmdDelete.Enabled = True
        txtFeeRemark = lstFeeList.SelectedItem.SubItems(2)
        txtFeeRemark.Enabled = False
        'CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstFeeList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstFeeList.SortKey = ColumnHeader.Index - 1
    lstFeeList.Sorted = True
    If lstFeeList.SortOrder = lvwAscending Then
       lstFeeList.SortOrder = lvwDescending
    Else
       lstFeeList.SortOrder = lvwAscending
    End If
    
    txtFeeCode.Enabled = False
End Sub

Private Sub lstFeeList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstFeeList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtFeeCode = lstFeeList.SelectedItem.Text
        txtFeeCode.Enabled = False
        txtFeeDesc = lstFeeList.SelectedItem.SubItems(1)
        txtFeeDesc.Enabled = False
        CmdDelete.Enabled = True
        txtFeeRemark = lstFeeList.SelectedItem.SubItems(2)
        txtFeeRemark.Enabled = False
        'CmdSave.Enabled = False
    End If

End Sub

Private Sub lstFeeList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstFeeList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtFeeCode = lstFeeList.SelectedItem.Text
        txtFeeCode.Enabled = False
        txtFeeDesc = lstFeeList.SelectedItem.SubItems(1)
        txtFeeDesc.Enabled = False
        CmdDelete.Enabled = True
        txtFeeRemark = lstFeeList.SelectedItem.SubItems(2)
        txtFeeRemark.Enabled = False
        'CmdSave.Enabled = False
    End If

End Sub

'&&&&&&& *** END SAVE, EDIT, ADD, DELETE FEE *** &&&&&&&'


Private Sub SSTab1_Click(PreviousTab As Integer)
    'lstFeeList.ListItems.Clear
    'Call ClearForm(Me)

End Sub

Private Sub txtCoPayment_Keypress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtFeeCode_Keypress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

