VERSION 5.00
Begin VB.Form frmShowInvestigation 
   Caption         =   "Investigation"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   7080
      TabIndex        =   29
      Top             =   6120
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   5820
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      Begin VB.TextBox TxtCPTCode 
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
         Left            =   1680
         TabIndex        =   14
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox TxtCPTDesc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   1680
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   4545
         Width           =   6015
      End
      Begin VB.TextBox TxtMedixDesc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   1680
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1560
         Width           =   6015
      End
      Begin VB.TextBox TxtMMADesc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   1680
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   2820
         Width           =   6015
      End
      Begin VB.TextBox TxtBodyPartMMA 
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
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtBodyPartMedix 
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
         Left            =   1680
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtOPCS 
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
         Left            =   1680
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtSurgicalCat 
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
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtAnaeCat 
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
         Left            =   1680
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtMMAPart 
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
         Left            =   5280
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtPage 
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
         Left            =   5280
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtPartAFees 
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
         Left            =   5280
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtSurgicalFee 
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
         Left            =   5280
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtAnaeFees 
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
         Left            =   5280
         TabIndex        =   1
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "CPT Code"
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
         TabIndex        =   28
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "CPT Description"
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
         TabIndex        =   27
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Medix Description"
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
         TabIndex        =   26
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Body Part - Medix"
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
         TabIndex        =   25
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "MMA Description"
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
         TabIndex        =   24
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Body Part - MMA"
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
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "OPCS "
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
         TabIndex        =   22
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Surgical Cat."
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
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Anae's Cat."
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
         TabIndex        =   20
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Surgical Fees"
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
         Left            =   3840
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Anae's Fees"
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
         Left            =   3840
         TabIndex        =   18
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Part A Fees"
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
         Left            =   3840
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Page"
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
         Left            =   3840
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "MMA Part"
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
         Left            =   3840
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Investigation"
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
      Height          =   225
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   7845
   End
End
Attribute VB_Name = "frmShowInvestigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer
Public tmpOPCS As String

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Source = 1 Then
        If frmMMASchedule.lstMMAListing.ListItems.Count > 0 Then
            If Trim(frmMMASchedule.lstMMAListing.SelectedItem.SubItems(1)) <> "" Then
                TxtBodyPartMMA = frmMMASchedule.lstMMAListing.SelectedItem.SubItems(1)
                txtOPCS = frmMMASchedule.lstMMAListing.SelectedItem.SubItems(2)
                txtSurgicalCat = frmMMASchedule.lstMMAListing.SelectedItem.SubItems(4)
                txtSurgicalFee = frmMMASchedule.lstMMAListing.SelectedItem.SubItems(6)
                txtAnaeCat = frmMMASchedule.lstMMAListing.SelectedItem.SubItems(5)
                txtAnaeFees = frmMMASchedule.lstMMAListing.SelectedItem.SubItems(7)
                txtPartAFees = frmMMASchedule.lstMMAListing.SelectedItem.SubItems(8)
                txtPage = frmMMASchedule.lstMMAListing.SelectedItem.SubItems(9)
                txtMMAPart = frmMMASchedule.lstMMAListing.SelectedItem.SubItems(10)
                TxtMMADesc = frmMMASchedule.lstMMAListing.SelectedItem.SubItems(3)
            End If
        Else
            Unload Me
        End If
    ElseIf Source = 2 Then
        If Len(tmpOPCS) Then
            Call MMAListing
        Else
            Unload Me
        End If
    End If
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

Private Sub MMAListing(Optional strAppend As String)
Dim rst As New ADODB.Recordset
Dim TotalPLNCount As Integer
Dim MMAMaster As ListItem
Dim StrSql As String
Dim strsqlFees As String
Dim rstFees As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
    
    'lstMMAListing.ListItems.Clear
    'StrSql = "select ProcMainBPCode, OPCSCode, OPCSNarrative , PartACost, CPTCode,CPTDesc," & _
             "ProcSurgCat, ProcAnaesCat, ProcPageNo, ProcPart, SurgeonsFees, AnaesthetistsFees " & _
             " From VIEW_MMA_Schedule where 0=0 "
    
    'If Len(Trim(StrAppend)) Then
    '    StrSql = StrSql + StrAppend
    'End If
    
    strAppend = ""
    strAppend = strAppend + " And OPCSCode = '" & ChkStr(tmpOPCS) & "'"

    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    'rst.Open StrSql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "1"
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "Case_MMASchedule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    rst.CacheSize = 100
    Set rst = cmd.Execute
    
    If rst.EOF Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        TxtBodyPartMMA = IIf(Trim(rst!ProcMainBPCode) <> "", Trim(rst!ProcMainBPCode), "")
        txtOPCS = IIf(Trim(rst!OPCSCode) <> "", Trim(rst!OPCSCode), "")
        txtSurgicalCat = IIf(Trim(rst!ProcSurgCat) <> "", Mid(rst!ProcSurgCat, 2, 8), "")
        txtSurgicalFee = IIf(Trim(rst!SurgeonsFees) <> "", rst!SurgeonsFees, "")
        txtAnaeCat = IIf(Trim(rst!ProcAnaesCat) <> "", Mid(rst!ProcAnaesCat, 2, 8), "")
        txtAnaeFees = IIf(Trim(rst!AnaesthetistsFees) <> "", rst!AnaesthetistsFees, "")
        txtPartAFees = IIf(Trim(rst!PartACost) <> "", rst!PartACost, "")
        txtPage = IIf(Trim(rst!ProcPageNo) <> "", rst!ProcPageNo, "")
        txtMMAPart = IIf(Trim(rst!ProcPart) <> "", rst!ProcPart, "")
        TxtMMADesc = IIf(Trim(rst!OPCSNarrative) <> "", rst!OPCSNarrative, "")
    End If
    rst.Close
    Set rst = Nothing
   ' medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
End Sub




